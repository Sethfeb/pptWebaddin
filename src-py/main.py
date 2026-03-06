"""
main.py
EquipSpec Tool 진입점 + 시스템 트레이 상주

실행 흐름:
  1. tkinter 루트 창 생성 (숨김 상태)
  2. 시스템 트레이 아이콘 생성 (별도 스레드)
  3. 설정 로드 → AuthService / SharePointService / CacheService 초기화
  4. 캐시에 토큰 있으면 자동 로그인 상태 유지
  5. 토큰 없으면 첫 검색/새로고침 시 로그인 다이얼로그 표시

스레드 구조:
  - 메인 스레드: tkinter 이벤트 루프 (root.mainloop)
  - 트레이 스레드: pystray (daemon=True)
  - 프리페치/검색 스레드: threading.Thread (daemon=True)
  - Device Code Flow: 백그라운드 스레드에서 블로킹 대기

COM 주의사항:
  - PptService 호출은 반드시 메인 스레드(tkinter after())에서 실행.
  - 트레이 콜백에서 직접 COM 호출 금지.
"""
import threading
import tkinter as tk
import webbrowser
from tkinter import messagebox, simpledialog
from typing import List, Optional

import pystray
from PIL import Image, ImageDraw

import settings as cfg
from models import SpecRecord
from services.auth_service import AuthError, AuthService
from services.cache_service import CacheService
from services.ppt_service import PptNotRunningError, PptService
from services.sharepoint_service import SharePointService
from views.search_window import SearchWindow
from views.settings_window import SettingsWindow


# ──────────────────────────────────────────────────────────────────────
# 트레이 아이콘 이미지 생성 (Pillow, 외부 파일 불필요)
# ──────────────────────────────────────────────────────────────────────

def _make_tray_icon(logged_in: bool = True) -> Image.Image:
    color = (0, 112, 192) if logged_in else (160, 160, 160)
    img = Image.new("RGBA", (64, 64), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)
    draw.ellipse((4, 4, 60, 60), fill=color)
    draw.text((18, 18), "ES", fill="white")
    return img


# ──────────────────────────────────────────────────────────────────────
# 로그인 안내 다이얼로그
# ──────────────────────────────────────────────────────────────────────

class LoginDialog(tk.Toplevel):
    """Device Code Flow 로그인 안내 창."""

    def __init__(self, parent: tk.Misc, url: str, code: str) -> None:
        super().__init__(parent)
        self.title("Microsoft 로그인 필요")
        self.geometry("460x220")
        self.resizable(False, False)
        self.grab_set()

        tk.Label(self, text="SharePoint 접근을 위해 Microsoft 계정 로그인이 필요합니다.",
                 wraplength=420, justify=tk.LEFT, pady=8).pack(padx=16)

        frame = tk.Frame(self, bd=1, relief=tk.SUNKEN, bg="#f0f0f0")
        frame.pack(fill=tk.X, padx=16, pady=4)

        tk.Label(frame, text="1. 아래 코드를 복사하세요:", anchor=tk.W, bg="#f0f0f0").pack(
            fill=tk.X, padx=8, pady=(6, 0))

        code_frame = tk.Frame(frame, bg="#f0f0f0")
        code_frame.pack(fill=tk.X, padx=8, pady=4)
        code_var = tk.StringVar(value=code)
        tk.Entry(code_frame, textvariable=code_var, font=("Consolas", 14, "bold"),
                 state="readonly", width=16).pack(side=tk.LEFT)
        tk.Button(code_frame, text="복사", command=lambda: self._copy(code)).pack(
            side=tk.LEFT, padx=6)

        tk.Label(frame, text="2. 브라우저에서 로그인 후 코드를 입력하세요:", anchor=tk.W,
                 bg="#f0f0f0").pack(fill=tk.X, padx=8)
        tk.Button(frame, text="브라우저 열기", command=lambda: webbrowser.open(url),
                  bg="#0078d4", fg="white", relief=tk.FLAT, padx=8, pady=4).pack(
            anchor=tk.W, padx=8, pady=(4, 8))

        tk.Label(self, text="로그인 완료 후 이 창은 자동으로 닫힙니다.",
                 foreground="gray", pady=4).pack()

        self.protocol("WM_DELETE_WINDOW", lambda: None)  # 닫기 버튼 비활성화

    def _copy(self, text: str) -> None:
        self.clipboard_clear()
        self.clipboard_append(text)

    def close(self) -> None:
        self.grab_release()
        self.destroy()


# ──────────────────────────────────────────────────────────────────────
# 애플리케이션 컨트롤러
# ──────────────────────────────────────────────────────────────────────

class App:
    def __init__(self) -> None:
        self._root = tk.Tk()
        self._root.withdraw()

        self._cache: CacheService = CacheService(cfg.get("cache_ttl_seconds"))
        self._auth_service: Optional[AuthService] = None
        self._sp_service: Optional[SharePointService] = None
        self._search_win: Optional[SearchWindow] = None
        self._settings_win: Optional[SettingsWindow] = None
        self._tray: Optional[pystray.Icon] = None
        self._login_dialog: Optional[LoginDialog] = None

        self._init_services()
        self._init_windows()
        self._init_tray()

    # ------------------------------------------------------------------
    # 초기화
    # ------------------------------------------------------------------

    def _init_services(self) -> None:
        if not cfg.is_configured():
            return
        site_url: str = cfg.get("sharepoint_url")
        self._auth_service = AuthService(site_url)
        self._sp_service = SharePointService(
            site_url=site_url,
            token_provider=self._get_token_blocking,
        )
        # 캐시에 토큰이 있으면 조용히 프리페치
        if self._auth_service.is_logged_in and cfg.get("prefetch_on_start"):
            self._start_prefetch()

    def _init_windows(self) -> None:
        self._search_win = SearchWindow(
            parent=self._root,
            on_search=self._search,
            on_insert=self._insert,
            on_replace_all=self._replace_all,
            on_refresh=self._refresh,
        )
        self._settings_win = SettingsWindow(
            parent=self._root,
            on_save=self._on_settings_saved,
        )

    def _init_tray(self) -> None:
        logged_in = self._auth_service is not None and self._auth_service.is_logged_in
        menu = pystray.Menu(
            pystray.MenuItem("검색 패널 열기", self._tray_open_search, default=True),
            pystray.MenuItem("단축어 일괄치환 (현재 슬라이드)", self._tray_replace),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem("설정", self._tray_open_settings),
            pystray.MenuItem("캐시 새로고침", self._tray_refresh),
            pystray.MenuItem("로그아웃", self._tray_logout),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem("종료", self._tray_quit),
        )
        self._tray = pystray.Icon(
            "EquipSpec",
            _make_tray_icon(logged_in),
            "EquipSpec Tool",
            menu,
        )
        t = threading.Thread(target=self._tray.run, daemon=True)
        t.start()

    # ------------------------------------------------------------------
    # 인증
    # ------------------------------------------------------------------

    def _get_token_blocking(self) -> str:
        """
        토큰 획득 (블로킹, 백그라운드 스레드에서 호출).
        캐시 히트 시 즉시 반환.
        미스 시 Device Code Flow → UI 다이얼로그 표시 후 대기.
        """
        if self._auth_service is None:
            raise AuthError("AuthService가 초기화되지 않았습니다.")

        # silent 시도
        token = self._auth_service.acquire_token_silent()
        if token:
            return token

        # Device Code Flow 필요 — 로그인 다이얼로그를 메인 스레드에서 열기
        dialog_ready = threading.Event()
        dialog_ref: list = []  # mutable container for cross-thread access

        def _show_dialog(url: str, code: str) -> None:
            def _open() -> None:
                dlg = LoginDialog(self._root, url, code)
                self._login_dialog = dlg
                dialog_ref.append(dlg)
                dialog_ready.set()
            self._root.after(0, _open)

        try:
            token = self._auth_service.acquire_token_device_flow(_show_dialog)
        finally:
            # 로그인 완료 후 다이얼로그 닫기
            def _close_dialog() -> None:
                if self._login_dialog is not None:
                    try:
                        self._login_dialog.close()
                    except Exception:
                        pass
                    self._login_dialog = None
                # 트레이 아이콘 색상 업데이트
                if self._tray is not None:
                    self._tray.icon = _make_tray_icon(True)
            self._root.after(0, _close_dialog)

        return token

    # ------------------------------------------------------------------
    # 트레이 콜백
    # ------------------------------------------------------------------

    def _tray_open_search(self, *_: object) -> None:
        self._root.after(0, self._show_search)

    def _tray_open_settings(self, *_: object) -> None:
        self._root.after(0, self._show_settings)

    def _tray_replace(self, *_: object) -> None:
        self._root.after(0, self._replace_all)

    def _tray_refresh(self, *_: object) -> None:
        self._root.after(0, self._start_prefetch)

    def _tray_logout(self, *_: object) -> None:
        self._root.after(0, self._do_logout)

    def _tray_quit(self, *_: object) -> None:
        self._root.after(0, self._quit)

    # ------------------------------------------------------------------
    # 창 표시
    # ------------------------------------------------------------------

    def _show_search(self) -> None:
        if not cfg.is_configured():
            messagebox.showwarning(
                "설정 필요",
                "SharePoint URL을 먼저 설정하세요.\n[설정] 창을 엽니다.",
            )
            self._show_settings()
            return
        assert self._search_win is not None
        self._search_win.show()

    def _show_settings(self) -> None:
        assert self._settings_win is not None
        self._settings_win.show()

    # ------------------------------------------------------------------
    # 서비스 콜백
    # ------------------------------------------------------------------

    def _search(self, keyword: str) -> List[SpecRecord]:
        cached = self._cache.search(keyword)
        if cached:
            return cached
        if self._sp_service is None:
            return []
        results = self._sp_service.search(keyword)
        if not results and self._sp_service.last_error:
            err_msg = self._sp_service.last_error
            self._root.after(0, lambda: messagebox.showerror(
                "SharePoint 연결 오류", err_msg
            ))
        for rec in results:
            self._cache.put(rec)
        return results

    def _insert(self, rec: SpecRecord) -> bool:
        return PptService.insert_text(rec.display_value)

    def _replace_all(self) -> None:
        if not self._cache.is_fully_loaded:
            self._start_prefetch_sync()

        lookup = self._cache.build_lookup()
        if not lookup:
            messagebox.showinfo("알림", "캐시가 비어 있습니다. 먼저 새로고침을 실행하세요.")
            return

        try:
            count, not_found = PptService.replace_shortcodes(lookup, scope="active_slide")
        except PptNotRunningError:
            messagebox.showwarning("PPT 미실행", "PowerPoint가 실행 중이지 않습니다.")
            return

        msg = f"{count}개 단축어 치환 완료."
        if not_found:
            msg += f"\n미발견 단축어: {', '.join(not_found)}"
        messagebox.showinfo("치환 완료", msg)

    def _refresh(self) -> None:
        self._start_prefetch()

    def _do_logout(self) -> None:
        if self._auth_service is None:
            return
        if messagebox.askyesno("로그아웃", "Microsoft 계정에서 로그아웃하시겠습니까?"):
            self._auth_service.logout()
            self._cache.clear()
            if self._tray is not None:
                self._tray.icon = _make_tray_icon(False)
            messagebox.showinfo("로그아웃", "로그아웃되었습니다.\n다음 검색 시 다시 로그인이 필요합니다.")

    def _on_settings_saved(self) -> None:
        if self._sp_service is not None:
            self._sp_service.close()
        self._cache.clear()
        self._cache.set_ttl(cfg.get("cache_ttl_seconds"))
        if cfg.is_configured():
            site_url: str = cfg.get("sharepoint_url")
            self._auth_service = AuthService(site_url)
            self._sp_service = SharePointService(
                site_url=site_url,
                token_provider=self._get_token_blocking,
            )
            if cfg.get("prefetch_on_start"):
                self._start_prefetch()
        else:
            self._auth_service = None
            self._sp_service = None

    # ------------------------------------------------------------------
    # 프리페치
    # ------------------------------------------------------------------

    def _start_prefetch(self) -> None:
        t = threading.Thread(target=self._prefetch_worker, daemon=True)
        t.start()

    def _start_prefetch_sync(self) -> None:
        self._prefetch_worker()

    def _prefetch_worker(self) -> None:
        if self._sp_service is None:
            return
        try:
            items = self._sp_service.get_all()
            self._cache.put_all(items)
            print(f"[EquipSpec] 프리페치 완료: {len(items)}건")
        except Exception as exc:
            print(f"[EquipSpec][ERROR] 프리페치 실패: {exc}")

    # ------------------------------------------------------------------
    # 종료
    # ------------------------------------------------------------------

    def _quit(self) -> None:
        if self._tray is not None:
            self._tray.stop()
        if self._sp_service is not None:
            self._sp_service.close()
        self._root.destroy()

    def run(self) -> None:
        self._root.mainloop()


# ──────────────────────────────────────────────────────────────────────
# 진입점
# ──────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    try:
        import pythoncom  # type: ignore
        pythoncom.CoInitialize()
    except ImportError:
        pass

    app = App()
    app.run()
