"""
search_window.py
설비 사양 검색 패널 (tkinter)

레이아웃:
  ┌─────────────────────────────────────────────────┐
  │  [검색어 입력창]  [검색]  [전체 새로고침]         │
  │  ─────────────────────────────────────────────  │
  │  결과 목록 (Treeview)                            │
  │  EquipID | ShortCode | SpecName | Value | Unit  │
  │  ─────────────────────────────────────────────  │
  │  [단축어 일괄치환]  [선택 삽입]  [닫기]           │
  │  상태바                                          │
  └─────────────────────────────────────────────────┘

스레드 정책:
  - SharePoint 검색/프리페치는 threading.Thread 로 백그라운드 실행.
  - UI 업데이트는 반드시 after() 로 메인 스레드에서 실행.
"""
import threading
import tkinter as tk
from tkinter import messagebox, ttk
from typing import Callable, List, Optional

from models import SpecRecord


class SearchWindow(tk.Toplevel):
    """설비 사양 검색 패널 창."""

    def __init__(
        self,
        parent: tk.Misc,
        on_search: Callable[[str], List[SpecRecord]],
        on_insert: Callable[[SpecRecord], bool],
        on_replace_all: Callable[[], None],
        on_refresh: Callable[[], None],
    ) -> None:
        """
        Parameters
        ----------
        parent        : 부모 tk 위젯
        on_search     : (keyword) -> List[SpecRecord]
        on_insert     : (SpecRecord) -> bool  (PPT 삽입)
        on_replace_all: () -> None  (단축어 일괄치환)
        on_refresh    : () -> None  (캐시 전체 새로고침)
        """
        super().__init__(parent)
        self._on_search = on_search
        self._on_insert = on_insert
        self._on_replace_all = on_replace_all
        self._on_refresh = on_refresh
        self._results: List[SpecRecord] = []
        self._searching = False

        self.title("EquipSpec - 설비 사양 검색")
        self.geometry("780x480")
        self.resizable(True, True)
        self.minsize(600, 360)

        self._build_ui()
        self._bind_keys()

        # 창 닫기 시 숨기기 (destroy 대신)
        self.protocol("WM_DELETE_WINDOW", self.withdraw)

    # ------------------------------------------------------------------
    # UI 구성
    # ------------------------------------------------------------------

    def _build_ui(self) -> None:
        # ── 검색 바 ──────────────────────────────────────────────────
        top_frame = ttk.Frame(self, padding=(8, 6))
        top_frame.pack(fill=tk.X)

        ttk.Label(top_frame, text="검색어:").pack(side=tk.LEFT)

        self._search_var = tk.StringVar()
        self._entry = ttk.Entry(top_frame, textvariable=self._search_var, width=30)
        self._entry.pack(side=tk.LEFT, padx=(4, 4))

        self._btn_search = ttk.Button(top_frame, text="검색", command=self._do_search)
        self._btn_search.pack(side=tk.LEFT, padx=(0, 4))

        self._btn_refresh = ttk.Button(top_frame, text="전체 새로고침", command=self._do_refresh)
        self._btn_refresh.pack(side=tk.LEFT)

        ttk.Separator(self, orient=tk.HORIZONTAL).pack(fill=tk.X, padx=8)

        # ── 결과 목록 ─────────────────────────────────────────────────
        list_frame = ttk.Frame(self)
        list_frame.pack(fill=tk.BOTH, expand=True, padx=8, pady=4)

        columns = ("equip_id", "short_code", "spec_name", "spec_value", "unit", "revision")
        self._tree = ttk.Treeview(
            list_frame,
            columns=columns,
            show="headings",
            selectmode="browse",
        )
        col_cfg = [
            ("equip_id",   "설비 ID",   90),
            ("short_code", "단축어",   130),
            ("spec_name",  "사양명",   180),
            ("spec_value", "값",        90),
            ("unit",       "단위",      60),
            ("revision",   "Rev",       40),
        ]
        for col, heading, width in col_cfg:
            self._tree.heading(col, text=heading, anchor=tk.W)
            self._tree.column(col, width=width, minwidth=40, anchor=tk.W)

        vsb = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self._tree.yview)
        self._tree.configure(yscrollcommand=vsb.set)
        self._tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)

        ttk.Separator(self, orient=tk.HORIZONTAL).pack(fill=tk.X, padx=8)

        # ── 하단 버튼 ─────────────────────────────────────────────────
        btn_frame = ttk.Frame(self, padding=(8, 4))
        btn_frame.pack(fill=tk.X)

        self._btn_replace = ttk.Button(
            btn_frame, text="단축어 일괄치환 (현재 슬라이드)",
            command=self._do_replace_all,
        )
        self._btn_replace.pack(side=tk.LEFT, padx=(0, 6))

        self._btn_insert = ttk.Button(
            btn_frame, text="선택 삽입",
            command=self._do_insert,
        )
        self._btn_insert.pack(side=tk.LEFT, padx=(0, 6))

        ttk.Button(btn_frame, text="닫기", command=self.withdraw).pack(side=tk.RIGHT)

        # ── 상태바 ────────────────────────────────────────────────────
        self._status_var = tk.StringVar(value="준비")
        ttk.Label(self, textvariable=self._status_var, anchor=tk.W, relief=tk.SUNKEN).pack(
            fill=tk.X, padx=0, pady=(2, 0)
        )

    def _bind_keys(self) -> None:
        self._entry.bind("<Return>", lambda _: self._do_search())
        self._tree.bind("<Double-1>", lambda _: self._do_insert())

    # ------------------------------------------------------------------
    # 이벤트 핸들러
    # ------------------------------------------------------------------

    def _do_search(self) -> None:
        if self._searching:
            return
        keyword = self._search_var.get().strip()
        if not keyword:
            self._set_status("검색어를 입력하세요.")
            return

        self._searching = True
        self._set_status("검색 중...")
        self._btn_search.config(state=tk.DISABLED)

        def _worker() -> None:
            results = self._on_search(keyword)
            self.after(0, lambda: self._update_results(results))

        threading.Thread(target=_worker, daemon=True).start()

    def _do_refresh(self) -> None:
        self._set_status("전체 데이터 새로고침 중...")
        self._btn_refresh.config(state=tk.DISABLED)

        def _worker() -> None:
            self._on_refresh()
            self.after(0, self._on_refresh_done)

        threading.Thread(target=_worker, daemon=True).start()

    def _do_insert(self) -> None:
        rec = self._selected_record()
        if rec is None:
            self._set_status("삽입할 항목을 선택하세요.")
            return
        ok = self._on_insert(rec)
        if ok:
            self._set_status(f"삽입 완료: {rec.display_value}")
        else:
            messagebox.showwarning(
                "삽입 실패",
                "PowerPoint가 실행 중이지 않거나 텍스트박스가 선택되지 않았습니다.",
                parent=self,
            )

    def _do_replace_all(self) -> None:
        self._set_status("단축어 치환 중...")
        self._on_replace_all()

    # ------------------------------------------------------------------
    # 내부 유틸
    # ------------------------------------------------------------------

    def _update_results(self, results: List[SpecRecord]) -> None:
        self._results = results
        self._tree.delete(*self._tree.get_children())
        for rec in results:
            self._tree.insert(
                "", tk.END,
                values=(rec.equip_id, rec.short_code, rec.spec_name,
                        rec.spec_value, rec.unit, rec.revision),
            )
        self._searching = False
        self._btn_search.config(state=tk.NORMAL)
        self._set_status(f"검색 결과: {len(results)}건")

    def _on_refresh_done(self) -> None:
        self._btn_refresh.config(state=tk.NORMAL)
        self._set_status("새로고침 완료")

    def _selected_record(self) -> Optional[SpecRecord]:
        sel = self._tree.selection()
        if not sel:
            return None
        idx = self._tree.index(sel[0])
        if 0 <= idx < len(self._results):
            return self._results[idx]
        return None

    def _set_status(self, msg: str) -> None:
        self._status_var.set(msg)

    def show(self) -> None:
        """창을 앞으로 가져오고 표시."""
        self.deiconify()
        self.lift()
        self.focus_force()
        self._entry.focus_set()
