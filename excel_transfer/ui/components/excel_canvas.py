# ============================================
# excel_transfer/ui/components/excel_canvas.py
# ============================================
from __future__ import annotations

import tkinter as tk
from tkinter import ttk
from typing import Optional, Set, Tuple, ClassVar, List, Callable

import openpyxl
from services.excel_view_service import ExcelViewService, ExcelViewport

CELL_W = 80
CELL_H = 24
HEADER_H = 22
GRID = "#d0d0d0"
SEL = "#cde8ff"
HDR_BG = "#f0f0f0"


class ExcelCanvas(ttk.Frame):
    """
    Excel 表示 Canvas（UI部品）

    - Excel 最大行列に自動追従
    - 同期スクロール（クラス内）
    - on_change コールバック互換維持
    """

    # ---- 同期グループ（クラス変数） ----
    _GROUP: ClassVar[List["ExcelCanvas"]] = []
    _SYNCING: ClassVar[bool] = False

    def __init__(
        self,
        parent,
        service: ExcelViewService,
        on_change: Optional[Callable[[Set[Tuple[int, int]]], None]] = None,
        logger=None,
    ):
        super().__init__(parent)
        self.service = service
        self.on_change = on_change
        self.logger = logger

        self.cells: Set[Tuple[int, int]] = set()
        self.anchor: Optional[Tuple[int, int]] = None
        self._vp: Optional[ExcelViewport] = None

        self._build()
        self._bind()

        ExcelCanvas._GROUP.append(self)

    # -------------------------------------------------
    # logging
    # -------------------------------------------------
    def _log(self, level: str, msg: str) -> None:
        try:
            if self.logger:
                fn = getattr(self.logger, level.lower(), None)
                if fn:
                    fn(msg)
                else:
                    self.logger.info(msg)
        except Exception:
            pass

    # -------------------------------------------------
    # UI
    # -------------------------------------------------
    def _build(self) -> None:
        self.rowconfigure(0, weight=1)
        self.columnconfigure(0, weight=1)

        self.canvas = tk.Canvas(self, bg="white", highlightthickness=0)
        self.vsb = ttk.Scrollbar(self, orient="vertical", command=self._yview)
        self.hsb = ttk.Scrollbar(self, orient="horizontal", command=self._xview)

        self.canvas.configure(
            yscrollcommand=self.vsb.set,
            xscrollcommand=self.hsb.set,
        )

        self.canvas.grid(row=0, column=0, sticky="nsew")
        self.vsb.grid(row=0, column=1, sticky="ns")
        self.hsb.grid(row=1, column=0, sticky="ew")

    def _bind(self) -> None:
        self.canvas.bind("<Button-1>", self._click)
        self.canvas.bind("<Control-Button-1>", self._ctrl_click)
        self.canvas.bind("<Shift-Button-1>", self._shift_click)

        self.canvas.bind("<MouseWheel>", self._wheel_y)
        self.canvas.bind("<Shift-MouseWheel>", self._wheel_x)

    # -------------------------------------------------
    # 互換 API（呼ばれても無視）
    # -------------------------------------------------
    def set_view_size(self, _rows: int, _cols: int) -> None:
        return

    # -------------------------------------------------
    # Scroll（同期）
    # -------------------------------------------------
    def _yview(self, *args) -> None:
        self.canvas.yview(*args)
        self._sync_from_self()

    def _xview(self, *args) -> None:
        self.canvas.xview(*args)
        self._sync_from_self()

    def _wheel_y(self, e) -> str:
        self.canvas.yview_scroll(-1 if e.delta > 0 else 1, "units")
        self._sync_from_self()
        return "break"

    def _wheel_x(self, e) -> str:
        self.canvas.xview_scroll(-1 if e.delta > 0 else 1, "units")
        self._sync_from_self()
        return "break"

    def _sync_from_self(self) -> None:
        if ExcelCanvas._SYNCING:
            return
        try:
            ExcelCanvas._SYNCING = True
            x0, _ = self.canvas.xview()
            y0, _ = self.canvas.yview()

            for c in ExcelCanvas._GROUP:
                if c is not self:
                    c.canvas.xview_moveto(x0)
                    c.canvas.yview_moveto(y0)
        finally:
            ExcelCanvas._SYNCING = False

    # -------------------------------------------------
    # Public
    # -------------------------------------------------
    def refresh(self) -> None:
        max_r, max_c = self.service.get_sheet_size()
        if max_r <= 0 or max_c <= 0:
            return

        self._vp = ExcelViewport(1, 1, max_r, max_c)

        w = max_c * CELL_W
        h = HEADER_H + max_r * CELL_H
        self.canvas.configure(scrollregion=(0, 0, w, h))

        self.redraw()

    def clear_selection(self) -> None:
        self.cells.clear()
        self.anchor = None
        self.redraw()
        if self.on_change:
            self.on_change(self.cells)

    # -------------------------------------------------
    # Drawing
    # -------------------------------------------------
    def redraw(self) -> None:
        if not self._vp:
            return

        self.canvas.delete("all")

        # header
        for c in range(1, self._vp.cols + 1):
            x = (c - 1) * CELL_W
            self.canvas.create_rectangle(
                x, 0, x + CELL_W, HEADER_H,
                fill=HDR_BG, outline=GRID
            )
            self.canvas.create_text(
                x + CELL_W / 2,
                HEADER_H / 2,
                text=openpyxl.utils.get_column_letter(c),
            )

        # cells
        for r in range(1, self._vp.rows + 1):
            y = HEADER_H + (r - 1) * CELL_H
            row_vals = self.service.get_row_texts(r, 1, self._vp.cols)
            for c_i, txt in enumerate(row_vals):
                c = c_i + 1
                x = c_i * CELL_W
                self.canvas.create_rectangle(
                    x, y, x + CELL_W, y + CELL_H,
                    outline=GRID,
                    fill=SEL if (r, c) in self.cells else "white",
                )
                if txt:
                    self.canvas.create_text(
                        x + 4, y + CELL_H / 2,
                        anchor="w", text=txt
                    )

    # -------------------------------------------------
    # Selection
    # -------------------------------------------------
    def _cell_from_event(self, e) -> Optional[Tuple[int, int]]:
        x = self.canvas.canvasx(e.x)
        y = self.canvas.canvasy(e.y)
        if y < HEADER_H:
            return None

        c = int(x // CELL_W) + 1
        r = int((y - HEADER_H) // CELL_H) + 1
        return (r, c)

    def _click(self, e) -> None:
        cell = self._cell_from_event(e)
        if not cell:
            return
        self.cells = {cell}
        self.anchor = cell
        self.redraw()
        if self.on_change:
            self.on_change(self.cells)

    def _ctrl_click(self, e) -> None:
        cell = self._cell_from_event(e)
        if not cell:
            return
        if cell in self.cells:
            self.cells.remove(cell)
        else:
            self.cells.add(cell)
        self.anchor = cell
        self.redraw()
        if self.on_change:
            self.on_change(self.cells)

    def _shift_click(self, e) -> None:
        if not self.anchor:
            return
        cell = self._cell_from_event(e)
        if not cell:
            return

        r1, c1 = self.anchor
        r2, c2 = cell
        rr1, rr2 = (r1, r2) if r1 <= r2 else (r2, r1)
        cc1, cc2 = (c1, c2) if c1 <= c2 else (c2, c1)

        self.cells = {
            (r, c)
            for r in range(rr1, rr2 + 1)
            for c in range(cc1, cc2 + 1)
        }
        self.redraw()
        if self.on_change:
            self.on_change(self.cells)
