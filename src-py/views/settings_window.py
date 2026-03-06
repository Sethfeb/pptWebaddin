"""
settings_window.py
설정 창 (tkinter)

설정 항목:
  - SharePoint 사이트 URL
  - 캐시 TTL (초)
  - 시작 시 전체 데이터 프리페치 여부
"""
import re
import tkinter as tk
from tkinter import messagebox, ttk
from typing import Callable

import settings as cfg

# SharePoint 공유 링크 패턴 (/:l:/, /:f:/, /:b:/ 등)
_SHARE_LINK_RE = re.compile(r"sharepoint\.com/:[a-z]:/")
# 올바른 사이트 URL 패턴: .../sites/xxx 또는 .../teams/xxx
_SITE_URL_RE = re.compile(r"sharepoint\.com/(sites|teams)/[^/?#]+", re.IGNORECASE)


def _validate_sharepoint_url(url: str) -> str:
    """
    URL 형식 검증.
    반환값: "" (정상) 또는 오류 메시지 문자열
    """
    if not url:
        return ""  # 빈 값은 허용 (미설정 상태)
    if not url.startswith("https://"):
        return "URL은 https:// 로 시작해야 합니다."
    if _SHARE_LINK_RE.search(url):
        return (
            "공유 링크(/:l:/ 등)는 사용할 수 없습니다.\n\n"
            "SharePoint 사이트 루트 URL을 입력하세요.\n"
            "예) https://ati5344.sharepoint.com/sites/atimarketing"
        )
    if "sharepoint.com" in url and not _SITE_URL_RE.search(url):
        return (
            "SharePoint 사이트 URL 형식이 올바르지 않습니다.\n\n"
            "올바른 형식: https://{테넌트}.sharepoint.com/sites/{사이트명}\n"
            "예) https://ati5344.sharepoint.com/sites/atimarketing"
        )
    return ""


class SettingsWindow(tk.Toplevel):
    """설정 창."""

    def __init__(
        self,
        parent: tk.Misc,
        on_save: Callable[[], None],
    ) -> None:
        """
        Parameters
        ----------
        parent  : 부모 tk 위젯
        on_save : 설정 저장 완료 후 호출되는 콜백 (캐시/서비스 재초기화 등)
        """
        super().__init__(parent)
        self._on_save = on_save

        self.title("EquipSpec - 설정")
        self.geometry("480x260")
        self.resizable(False, False)

        self._build_ui()
        self._load_current()

        self.protocol("WM_DELETE_WINDOW", self.withdraw)

    # ------------------------------------------------------------------
    # UI 구성
    # ------------------------------------------------------------------

    def _build_ui(self) -> None:
        pad = {"padx": 12, "pady": 6}
        frame = ttk.Frame(self, padding=12)
        frame.pack(fill=tk.BOTH, expand=True)

        # SharePoint URL
        ttk.Label(frame, text="SharePoint 사이트 URL:").grid(
            row=0, column=0, sticky=tk.W, **pad
        )
        self._url_var = tk.StringVar()
        ttk.Entry(frame, textvariable=self._url_var, width=42).grid(
            row=0, column=1, sticky=tk.EW, **pad
        )
        ttk.Label(
            frame,
            text="예) https://ati5344.sharepoint.com/sites/atimarketing  (사이트 루트 URL)",
            foreground="gray",
        ).grid(row=1, column=1, sticky=tk.W, padx=12, pady=0)

        # 캐시 TTL
        ttk.Label(frame, text="캐시 유효 시간 (초):").grid(
            row=2, column=0, sticky=tk.W, **pad
        )
        self._ttl_var = tk.StringVar()
        ttk.Entry(frame, textvariable=self._ttl_var, width=10).grid(
            row=2, column=1, sticky=tk.W, **pad
        )

        # 프리페치
        self._prefetch_var = tk.BooleanVar()
        ttk.Checkbutton(
            frame,
            text="시작 시 전체 데이터 미리 불러오기",
            variable=self._prefetch_var,
        ).grid(row=3, column=1, sticky=tk.W, **pad)

        frame.columnconfigure(1, weight=1)

        ttk.Separator(self, orient=tk.HORIZONTAL).pack(fill=tk.X, padx=8)

        btn_frame = ttk.Frame(self, padding=(12, 8))
        btn_frame.pack(fill=tk.X)
        ttk.Button(btn_frame, text="저장", command=self._do_save).pack(side=tk.RIGHT, padx=(4, 0))
        ttk.Button(btn_frame, text="취소", command=self.withdraw).pack(side=tk.RIGHT)

    # ------------------------------------------------------------------
    # 로직
    # ------------------------------------------------------------------

    def _load_current(self) -> None:
        data = cfg.load()
        self._url_var.set(data.get("sharepoint_url", ""))
        self._ttl_var.set(str(data.get("cache_ttl_seconds", 300)))
        self._prefetch_var.set(bool(data.get("prefetch_on_start", True)))

    def _do_save(self) -> None:
        url = self._url_var.get().strip()
        ttl_str = self._ttl_var.get().strip()

        # URL 형식 검증
        url_err = _validate_sharepoint_url(url)
        if url_err:
            messagebox.showerror("URL 오류", url_err, parent=self)
            return

        try:
            ttl = int(ttl_str)
            if ttl < 10:
                raise ValueError
        except ValueError:
            messagebox.showerror(
                "입력 오류",
                "캐시 유효 시간은 10 이상의 정수로 입력하세요.",
                parent=self,
            )
            return

        cfg.save({
            "sharepoint_url": url,
            "cache_ttl_seconds": ttl,
            "prefetch_on_start": self._prefetch_var.get(),
            "shortcode_prefix": "!",
        })
        self._on_save()
        messagebox.showinfo("저장 완료", "설정이 저장되었습니다.", parent=self)
        self.withdraw()

    def show(self) -> None:
        """창을 앞으로 가져오고 표시."""
        self._load_current()
        self.deiconify()
        self.lift()
        self.focus_force()
