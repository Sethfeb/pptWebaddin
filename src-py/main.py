"""
main.py
EquipSpec Tool 진입점 + 시스템 트레이 상주

실행 흐름:
  1. tkinter 루트 창 생성 (숨김 상태)
  2. 시스템 트레이 아이콘 생성 (별도 스레드)
  3. 설정 로드 → SharePointService / CacheService 초기화
  4. 설정에 prefetch_on_start=True 이면 백그라운드 프리페치
  5. 트레이 메뉴 → 검색 패널 / 단축어 치환 / 설정 / 종료

스레드 구조:
  - 메인 스레드: tkinter 이벤트 루프 (root.mainloop)
  - 트레이 스레드: pystray (daemon=True)
  - 프리페치 스레드: threading.Thread (daemon=True)

COM 주의사항:
  - PptService 호출은 반드시 메인 스레드(tkinter after())에서 실행.
  - 트레이 콜백에서 직접 COM 호출 금지.
"""
import sys
import threading
import tkinter as tk
from tkinter import messagebox
from typing import List, Optional

import pystray
from PIL import Image, ImageDraw

import settings as cfg
from models import SpecRecord
from services.cache_service import CacheService
from services.ppt_service import PptNotRunningError, PptService
from services.sharepoint_service import SharePointService
from views.search_window import SearchWindow
from views.settings_window import SettingsWindow


# ──────────────────────────────────────────────────────────────────────
# 트레이 아이콘 이미지 생성 (Pillow, 외부 파일 불필요)
# ──────────────────────────────────────────────────────────────────────

def _make_tray_icon() -> Image.Image:
    img = Image.new("RGBA", (64, 64), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)
    draw.ellipse((4, 4, 60, 60), fill=(0, 112, 192))
    draw.text((18, 18), "ES", fill="white")
    return img


# ──────────────────────────────────────────────────────────────────────
# 애플리케이션 컨트롤러
# ──────────────────────────────────────────────────────────────────────

class App:
    def __init__(self) -> None:
        self._root = tk.Tk()
        self._root.withdraw()  # 메인 창 숨김

        self._cache: CacheService = CacheService(cfg.get("cache_ttl_seconds"))
        self._sp_service: Optional[SharePointService] = None
        self._search_win: Optional[SearchWindow] = None
        self._settings_win: Optional[SettingsWindow] = None
        self._tray: Optional[pystray.Icon] = None

        self._init_services()
        self._init_windows()
        self._init_tray()

    # ------------------------------------------------------------------
    # 초기화
    # ------------------------------------------------------------------

    def _init_services(self) -> None:
        if cfg.is_configured():
            self._sp_service = SharePointService(cfg.get("sharepoint_url"))
            if cfg.get("prefetch_on_start"):
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
        menu = pystray.Menu(
            pystray.MenuItem("검색 패널 열기", self._tray_open_search, default=True),
            pystray.MenuItem("단축어 일괄치환 (현재 슬라이드)", self._tray_replace),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem("설정", self._tray_open_settings),
            pystray.MenuItem("캐시 새로고침", self._tray_refresh),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem("종료", self._tray_quit),
        )
        self._tray = pystray.Icon(
            "EquipSpec",
            _make_tray_icon(),
            "EquipSpec Tool",
            menu,
        )
        t = threading.Thread(target=self._tray.run, daemon=True)
        t.start()

    # ------------------------------------------------------------------
    # 트레이 콜백 (트레이 스레드에서 호출 → after()로 메인 스레드 위임)
    # ------------------------------------------------------------------

    def _tray_open_search(self, *_: object) -> None:
        self._root.after(0, self._show_search)

    def _tray_open_settings(self, *_: object) -> None:
        self._root.after(0, self._show_settings)

    def _tray_replace(self, *_: object) -> None:
        self._root.after(0, self._replace_all)

    def _tray_refresh(self, *_: object) -> None:
        self._root.after(0, self._start_prefetch)

    def _tray_quit(self, *_: object) -> None:
        self._root.after(0, self._quit)

    # ------------------------------------------------------------------
    # 창 표시 (메인 스레드)
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
    # 서비스 콜백 (SearchWindow에서 호출)
    # ------------------------------------------------------------------

    def _search(self, keyword: str) -> List[SpecRecord]:
        """캐시 우선 검색 → 미스 시 SharePoint 검색."""
        cached = self._cache.search(keyword)
        if cached:
            return cached
        if self._sp_service is None:
            return []
        results = self._sp_service.search(keyword)
        for rec in results:
            self._cache.put(rec)
        return results

    def _insert(self, rec: SpecRecord) -> bool:
        """선택된 레코드를 PPT에 삽입."""
        return PptService.insert_text(rec.display_value)

    def _replace_all(self) -> None:
        """현재 슬라이드의 단축어 일괄 치환."""
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
        """전체 데이터 새로고침 (백그라운드)."""
        self._start_prefetch()

    def _on_settings_saved(self) -> None:
        """설정 저장 후 서비스 재초기화."""
        if self._sp_service is not None:
            self._sp_service.close()
        self._cache.clear()
        self._cache.set_ttl(cfg.get("cache_ttl_seconds"))
        if cfg.is_configured():
            self._sp_service = SharePointService(cfg.get("sharepoint_url"))
            if cfg.get("prefetch_on_start"):
                self._start_prefetch()
        else:
            self._sp_service = None

    # ------------------------------------------------------------------
    # 프리페치
    # ------------------------------------------------------------------

    def _start_prefetch(self) -> None:
        t = threading.Thread(target=self._prefetch_worker, daemon=True)
        t.start()

    def _start_prefetch_sync(self) -> None:
        """동기 프리페치 (replace_all 호출 전 보장용)."""
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

    # ------------------------------------------------------------------
    # 실행
    # ------------------------------------------------------------------

    def run(self) -> None:
        self._root.mainloop()


# ──────────────────────────────────────────────────────────────────────
# 진입점
# ──────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    # COM STA 초기화 (pywin32)
    try:
        import pythoncom  # type: ignore
        pythoncom.CoInitialize()
    except ImportError:
        pass

    app = App()
    app.run()
