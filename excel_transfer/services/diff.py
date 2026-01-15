# excel_transfer/services/diff.py
from __future__ import annotations

import os
import sys
import shutil
from typing import Tuple, Dict, Any, List

# =====================================================
# project root 探索（models/dto.py を基準に決定）
# =====================================================
_THIS_DIR = os.path.dirname(os.path.abspath(__file__))


def _resolve_project_root(start: str) -> str:
    cur = start
    for _ in range(10):
        cand = os.path.join(cur, "models", "dto.py")
        if os.path.isfile(cand):
            return cur
        parent = os.path.dirname(cur)
        if parent == cur:
            break
        cur = parent
    raise RuntimeError("project root not found (models/dto.py not found)")


_PROJECT_ROOT = _resolve_project_root(_THIS_DIR)
if _PROJECT_ROOT not in sys.path:
    sys.path.insert(0, _PROJECT_ROOT)

# =====================================================
# imports
# =====================================================
import xlwings as xw
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Border, Side

from models.dto import DiffRequest, LogFn

# =====================================================
# Excel Diff Service
# =====================================================
class ExcelDiffService:
    def __init__(self, req: DiffRequest, logger, append_log: LogFn):
        self.req = req
        self.logger = logger
        self.append_log = append_log

        self.diff_cells: List[Tuple[int, int]] = []
        self.diff_shapes: List[str] = []
        self.out_file: str = ""

    # -------------------------------------------------
    # public
    # -------------------------------------------------
    def run(self) -> str:
        self.append_log("=== Diff開始 ===")

        if not os.path.exists(self.req.file_a) or not os.path.exists(self.req.file_b):
            raise ValueError("比較ファイルが存在しません。")

        self.out_file = os.path.splitext(self.req.file_b)[0] + "_DIFF.xlsx"
        shutil.copyfile(self.req.file_b, self.out_file)
        self.append_log(f"[INFO] 出力ファイル: {self.out_file}")

        self._diff_excel()

        if self.diff_cells:
            self._mark_cells_red()
            self.append_log(f"[OK] セル差分 {len(self.diff_cells)} 件を赤色反映")

        if self.diff_shapes:
            self._mark_shapes_red()
            self.append_log(f"[OK] 図形差分 {len(self.diff_shapes)} 件を赤枠反映")

        self.append_log("=== Diff完了 ===")
        return self.out_file

    # -------------------------------------------------
    # core diff
    # -------------------------------------------------
    def _diff_excel(self) -> None:
        app_a = app_b = None
        book_a = book_b = None

        try:
            app_a = xw.App(visible=False, add_book=False)
            app_b = xw.App(visible=False, add_book=False)

            book_a = app_a.books.open(self.req.file_a, read_only=True)
            book_b = app_b.books.open(self.req.file_b, read_only=True)

            sht_a = self._sheet_or_first(book_a, getattr(self.req, "sheet_a", ""))
            sht_b = self._sheet_or_first(book_b, getattr(self.req, "sheet_b", ""))

            self.append_log(f"比較シート: {sht_a.name} ↔ {sht_b.name}")

            range_a = getattr(self.req, "range_a", "").strip()
            range_b = getattr(self.req, "range_b", "").strip()

            a_dict = self._read_sheet_to_dict(sht_a, range_a)
            b_dict = self._read_sheet_to_dict(sht_b, range_b)

            a_keys = set(a_dict.keys())
            b_keys = set(b_dict.keys())

            for k in sorted(a_keys - b_keys):
                self.append_log(f"[DEL] {k}")
            for k in sorted(b_keys - a_keys):
                self.append_log(f"[ADD] {k}")

            for k in sorted(a_keys & b_keys):
                ra = a_dict[k]
                rb = b_dict[k]
                for col in ra.keys():
                    va = ra[col]["value"]
                    vb = rb[col]["value"]
                    if va != vb:
                        self.append_log(f"[MOD] {k} col={col} A={va} B={vb}")
                        self.diff_cells.append((rb[col]["row"], rb[col]["col"]))

            if self.req.compare_shapes:
                self.append_log("[INFO] 図形・画像比較開始")
                sa = self._read_shapes(sht_a)
                sb = self._read_shapes(sht_b)

                for name in sa.keys() - sb.keys():
                    self.append_log(f"[SHAPE-DEL] {name}")
                for name in sb.keys() - sa.keys():
                    self.append_log(f"[SHAPE-ADD] {name}")
                    self.diff_shapes.append(name)
                for name in sa.keys() & sb.keys():
                    if sa[name] != sb[name]:
                        self.append_log(f"[SHAPE-MOD] {name}")
                        self.diff_shapes.append(name)

        finally:
            for book in (book_a, book_b):
                try:
                    if book:
                        book.close()
                except Exception:
                    pass
            for app in (app_a, app_b):
                try:
                    if app:
                        app.kill()
                except Exception:
                    pass

    # -------------------------------------------------
    # helpers
    # -------------------------------------------------
    def _sheet_or_first(self, book: xw.Book, name: str):
        return book.sheets[name] if name else book.sheets[0]

    def _read_sheet_to_dict(
        self, sht: xw.Sheet, addr: str
    ) -> Dict[Tuple, Dict[str, Any]]:
        rng = sht.range(addr) if addr else sht.used_range
        vals = rng.value

        if not vals or not isinstance(vals, list):
            return {}

        headers = vals[0]
        if not isinstance(headers, list):
            headers = [headers]

        col_index = {str(h): i for i, h in enumerate(headers)}
        rows = vals[1:]

        start_row = rng.row
        start_col = rng.column

        out: Dict[Tuple, Dict[str, Any]] = {}

        for ridx, row in enumerate(rows, start=start_row + 1):
            if not isinstance(row, list):
                row = [row]

            key = (
                tuple(row[col_index.get(k, -1)] for k in self.req.key_cols)
                if self.req.key_cols
                else (ridx,)
            )

            record: Dict[str, Any] = {}
            for h, cidx in col_index.items():
                cell = sht.range((ridx, start_col + cidx))
                v = cell.formula if self.req.compare_formula else cell.value
                record[h] = {
                    "value": v,
                    "row": cell.row,
                    "col": cell.column,
                }

            out[key] = record

        return out

    def _read_shapes(self, sht: xw.Sheet) -> Dict[str, Dict[str, Any]]:
        out: Dict[str, Dict[str, Any]] = {}
        for shp in sht.api.Shapes:
            try:
                out[str(shp.Name)] = {
                    "top": shp.Top,
                    "left": shp.Left,
                    "width": shp.Width,
                    "height": shp.Height,
                }
            except Exception:
                pass
        return out

    # -------------------------------------------------
    # excel mark
    # -------------------------------------------------
    def _mark_cells_red(self) -> None:
        wb = load_workbook(self.out_file)
        ws = wb.active

        red_fill = PatternFill(
            start_color="FFFF6666", end_color="FFFF6666", fill_type="solid"
        )
        red_border = Border(
            left=Side(style="thin", color="FF0000"),
            right=Side(style="thin", color="FF0000"),
            top=Side(style="thin", color="FF0000"),
            bottom=Side(style="thin", color="FF0000"),
        )

        for r, c in self.diff_cells:
            cell = ws.cell(row=int(r), column=int(c))
            cell.fill = red_fill
            cell.border = red_border

        wb.save(self.out_file)

    def _mark_shapes_red(self) -> None:
        app = xw.App(visible=False, add_book=False)
        try:
            book = app.books.open(self.out_file)
            sht = book.sheets[0]
            for shp in sht.api.Shapes:
                if str(shp.Name) in self.diff_shapes:
                    try:
                        shp.Line.ForeColor.RGB = 255
                    except Exception:
                        pass
            book.save()
            book.close()
        finally:
            app.kill()


# =====================================================
# 既存互換API
# =====================================================
def run_diff(req: DiffRequest, ctx, logger, append_log: LogFn) -> str:
    svc = ExcelDiffService(req, logger, append_log)
    return svc.run()
