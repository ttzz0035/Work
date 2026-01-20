# ============================================
# excel_transfer/outputs/transfer_csv_builder.py
# ============================================
from __future__ import annotations

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path
from typing import List, Set, Tuple

import openpyxl

from services.excel_view_service import ExcelViewService
from ui.components.excel_canvas import ExcelCanvas


# =====================================================
# Transfer CSV Builder Dialog
# =====================================================
class TransferCsvBuilderDialog(tk.Toplevel):

    HEADERS = [
        "source_file", "source_sheet", "source_cell",
        "source_row_offset", "source_col_offset",
        "destination_file", "destination_sheet", "destination_cell",
        "destination_row_offset", "destination_col_offset",
    ]

    def __init__(self, parent, ctx, logger, on_created):
        super().__init__(parent)

        self.ctx = ctx
        self.logger = logger
        self.on_created = on_created

        self.src_cells: Set[Tuple[int, int]] = set()
        self.dst_cells: Set[Tuple[int, int]] = set()

        # サービス（A/Bで独立）
        self.src_service = ExcelViewService(logger=self.logger)
        self.dst_service = ExcelViewService(logger=self.logger)

        self.title("転記CSV作成")
        self.geometry("1500x850")

        self._build()

    # ---------- focus restore ----------
    def _restore_focus(self):
        self.after(0, self.lift)
        self.after(0, self.focus_force)

    # ---------- logging ----------
    def _log(self, msg: str) -> None:
        try:
            if self.logger:
                self.logger.info(msg)
        except Exception:
            pass

    # =================================================
    def _build(self):
        root = ttk.Frame(self, padding=6)
        root.pack(fill="both", expand=True)
        root.columnconfigure(0, weight=1)
        root.columnconfigure(1, weight=1)
        root.rowconfigure(2, weight=1)

        # ---- 候補 ----
        cand = ttk.LabelFrame(root, text="マッピング候補")
        cand.grid(row=0, column=0, columnspan=2, sticky="we")
        self.cand_var = tk.StringVar()
        ttk.Entry(cand, textvariable=self.cand_var, state="readonly").pack(
            fill="x", padx=4, pady=4
        )

        # ---- 確定テーブル ----
        conf = ttk.LabelFrame(root, text="確定済み（CSV内容）")
        conf.grid(row=1, column=0, columnspan=2, sticky="nsew")
        conf.rowconfigure(0, weight=1)
        conf.columnconfigure(0, weight=1)

        tv = ttk.Frame(conf)
        tv.grid(row=0, column=0, sticky="nsew")

        self.table = ttk.Treeview(
            tv,
            columns=self.HEADERS,
            show="headings",
            selectmode="extended",
        )
        for h in self.HEADERS:
            self.table.heading(h, text=h)
            self.table.column(h, width=140, anchor="center")

        vsb = ttk.Scrollbar(tv, orient="vertical", command=self.table.yview)
        hsb = ttk.Scrollbar(tv, orient="horizontal", command=self.table.xview)
        self.table.configure(
            yscrollcommand=vsb.set,
            xscrollcommand=hsb.set,
        )

        self.table.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")

        tv.rowconfigure(0, weight=1)
        tv.columnconfigure(0, weight=1)

        self.table.bind("<Control-a>", self._table_select_all)
        self.table.bind("<Control-A>", self._table_select_all)
        self.table.bind("<Delete>", self._table_delete_selected)

        btns = ttk.Frame(conf)
        btns.grid(row=1, column=0, sticky="we", pady=4)
        ttk.Button(btns, text="確定", command=self.confirm).pack(side="left")
        ttk.Button(btns, text="削除", command=self.remove).pack(side="left", padx=6)
        ttk.Button(btns, text="CSVを作成...", command=self.create_csv).pack(side="right")

        # ---- Excel 表（左右）----
        paned = ttk.Panedwindow(root, orient="horizontal")
        paned.grid(row=2, column=0, columnspan=2, sticky="nsew")

        self._build_side(paned, True)
        self._build_side(paned, False)

        self.refresh_candidate()

    # =================================================
    def _build_side(self, paned, is_src: bool):
        frm = ttk.Frame(paned)
        paned.add(frm, weight=1)
        frm.rowconfigure(1, weight=1)
        frm.columnconfigure(0, weight=1)

        title = "ファイルA（転記元）" if is_src else "ファイルB（転記先）"
        lf = ttk.LabelFrame(frm, text=title)
        lf.grid(row=0, column=0, sticky="we")
        lf.columnconfigure(1, weight=1)

        ttk.Label(lf, text="ブック").grid(row=0, column=0, sticky="w")
        cmb_book = ttk.Combobox(lf, state="readonly", width=60)
        cmb_book.grid(row=0, column=1, sticky="we", padx=4)

        ttk.Button(
            lf,
            text="追加...",
            command=lambda s=is_src: self._choose_books(s)
        ).grid(row=0, column=2, padx=(0, 4))

        ttk.Label(lf, text="シート").grid(row=1, column=0, sticky="w", pady=(4, 0))
        cmb_sheet = ttk.Combobox(lf, state="readonly", width=40)
        cmb_sheet.grid(row=1, column=1, sticky="w", padx=4, pady=(4, 0))

        service = self.src_service if is_src else self.dst_service
        canvas = ExcelCanvas(
            frm,
            service=service,
            on_change=self.on_src_change if is_src else self.on_dst_change,
            logger=self.logger,
        )
        canvas.grid(row=1, column=0, sticky="nsew", pady=(6, 0))

        cmb_book.bind("<<ComboboxSelected>>", lambda _e, s=is_src: self._on_book_selected(s))
        cmb_sheet.bind("<<ComboboxSelected>>", lambda _e, s=is_src: self._on_sheet_selected(s))

        if is_src:
            self.src_book = cmb_book
            self.src_sheet = cmb_sheet
            self.src_canvas = canvas
        else:
            self.dst_book = cmb_book
            self.dst_sheet = cmb_sheet
            self.dst_canvas = canvas

    # -------------------------
    # table key handlers
    # -------------------------
    def _table_select_all(self, _e=None):
        items = self.table.get_children()
        if items:
            self.table.selection_set(items)
        return "break"

    def _table_delete_selected(self, _e=None):
        self.remove()
        return "break"

    # -------------------------
    # book/sheet change
    # -------------------------
    def _choose_books(self, is_src: bool):
        paths = filedialog.askopenfilenames(
            filetypes=[("Excel", "*.xlsx;*.xlsm;*.xlsb;*.xls")]
        )
        self._restore_focus()
        if not paths:
            return

        svc = self.src_service if is_src else self.dst_service
        svc.add_books(list(paths))

        cmb_book = self.src_book if is_src else self.dst_book
        cmb_sheet = self.src_sheet if is_src else self.dst_sheet
        canvas = self.src_canvas if is_src else self.dst_canvas

        cmb_book["values"] = svc.get_book_paths()
        if svc.get_current_book_path():
            cmb_book.set(svc.get_current_book_path())

        cmb_sheet["values"] = svc.get_sheet_names()
        if svc.get_current_sheet_name():
            cmb_sheet.set(svc.get_current_sheet_name())

        canvas.refresh()

    def _on_book_selected(self, is_src: bool):
        svc = self.src_service if is_src else self.dst_service
        cmb_book = self.src_book if is_src else self.dst_book
        cmb_sheet = self.src_sheet if is_src else self.dst_sheet
        canvas = self.src_canvas if is_src else self.dst_canvas

        p = cmb_book.get().strip()
        if not p:
            return

        svc.select_book(p)
        cmb_sheet["values"] = svc.get_sheet_names()
        if svc.get_current_sheet_name():
            cmb_sheet.set(svc.get_current_sheet_name())

        canvas.refresh()

    def _on_sheet_selected(self, is_src: bool):
        svc = self.src_service if is_src else self.dst_service
        cmb_sheet = self.src_sheet if is_src else self.dst_sheet
        canvas = self.src_canvas if is_src else self.dst_canvas

        s = cmb_sheet.get().strip()
        if not s:
            return

        svc.select_sheet(s)
        canvas.refresh()

    # -------------------------
    # selection callbacks
    # -------------------------
    def on_src_change(self, cells):
        self.src_cells = set(cells or [])
        self.refresh_candidate()

    def on_dst_change(self, cells):
        self.dst_cells = set(cells or [])
        self.refresh_candidate()

    def refresh_candidate(self):
        sc = len(self.src_cells)
        dc = len(self.dst_cells)

        if sc:
            names = self._cell_name(self.src_cells)
            s_txt = f"{names[0]}〜{names[-1]} ({sc})"
        else:
            s_txt = "(未選択)"

        if dc:
            names = self._cell_name(self.dst_cells)
            d_txt = f"{names[0]}〜{names[-1]} ({dc})"
        else:
            d_txt = "(未選択)"

        self.cand_var.set(f"A: {s_txt}  →  B: {d_txt}")

    def _cell_name(self, cells: Set[Tuple[int, int]]) -> List[str]:
        return [
            f"{openpyxl.utils.get_column_letter(c)}{r}"
            for (r, c) in sorted(cells)
        ]

    # -------------------------
    # confirm/remove/csv
    # -------------------------
    def confirm(self):
        if not self.src_cells or not self.dst_cells:
            return

        if not self.src_service.get_current_book_path() or not self.src_service.get_current_sheet_name():
            return
        if not self.dst_service.get_current_book_path() or not self.dst_service.get_current_sheet_name():
            return

        src = self._cell_name(self.src_cells)
        dst = self._cell_name(self.dst_cells)

        if len(src) == len(dst):
            pairs = list(zip(src, dst))
        elif len(src) == 1:
            pairs = [(src[0], d) for d in dst]
        elif len(dst) == 1:
            pairs = [(s, dst[0]) for s in src]
        else:
            messagebox.showerror("転記", "セル数が一致しません")
            self._restore_focus()
            return

        s_book = self.src_service.get_current_book_path()
        s_sheet = self.src_service.get_current_sheet_name()
        d_book = self.dst_service.get_current_book_path()
        d_sheet = self.dst_service.get_current_sheet_name()

        for s, d in pairs:
            self.table.insert("", "end", values=[
                s_book, s_sheet, s, "0", "0",
                d_book, d_sheet, d, "0", "0"
            ])

        self.src_cells.clear()
        self.dst_cells.clear()
        self.src_canvas.clear_selection()
        self.dst_canvas.clear_selection()
        self.refresh_candidate()

    def remove(self):
        for iid in list(self.table.selection()):
            try:
                self.table.delete(iid)
            except Exception:
                pass

    def create_csv(self):
        if not self.table.get_children():
            return

        path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV Files", "*.csv")]
        )
        self._restore_focus()
        if not path:
            return

        lines = [",".join(self.HEADERS)]
        for iid in self.table.get_children():
            values = list(self.table.item(iid)["values"])
            for idx in (3, 4, 8, 9):
                if idx < len(values):
                    v = values[idx]
                    s = str(v).strip() if v is not None else ""
                    values[idx] = s if s else "0"
            lines.append(",".join(map(str, values)))

        Path(path).write_text("\n".join(lines) + "\n", encoding="utf-8-sig")

        if self.on_created:
            self.on_created(path)
        self.destroy()
