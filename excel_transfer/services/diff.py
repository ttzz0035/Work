from __future__ import annotations

import json
import os
import shutil
from datetime import datetime
from typing import Any, Dict, List

import xlwings as xw

from models.dto import DiffRequest, LogFn


class ExcelDiffService:
    def __init__(self, req: DiffRequest, logger, append_log: LogFn):
        self.req = req
        self.logger = logger
        self.append_log = append_log

        self.diff_cells: List[Dict[str, Any]] = []
        self.diff_shapes: List[Dict[str, Any]] = []

        self._meta: Dict[str, Any] = {}
        self._summary: Dict[str, Any] = {}

    # -------------------------------------------------
    # logging
    # -------------------------------------------------
    def _log(self, msg: str) -> None:
        try:
            self.append_log(msg)
        except Exception:
            pass
        try:
            if self.logger:
                self.logger.info(msg)
        except Exception:
            pass

    def _log_err(self, msg: str) -> None:
        try:
            self.append_log(msg)
        except Exception:
            pass
        try:
            if self.logger:
                self.logger.error(msg)
        except Exception:
            pass

    # -------------------------------------------------
    # public
    # -------------------------------------------------
    def run(self) -> str:
        self._log("=== Diff開始 ===")

        if not self.req.range_a or not self.req.range_b:
            raise ValueError("range_a / range_b は必須です（空不可）")

        base = str(getattr(self.req, "base_file", "B") or "B").upper()
        if base not in ("A", "B"):
            base = "B"

        sheet_mode = getattr(self.req, "sheet_mode", "index")
        if sheet_mode not in ("index", "name"):
            raise ValueError(f"invalid sheet_mode: {sheet_mode}")

        base_path = self.req.file_a if base == "A" else self.req.file_b
        other_path = self.req.file_b if base == "A" else self.req.file_a

        if not os.path.exists(base_path):
            raise ValueError(f"invalid base_path: {base_path}")
        if not os.path.exists(other_path):
            raise ValueError(f"invalid other_path: {other_path}")

        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        root, ext = os.path.splitext(base_path)
        diff_path = f"{root}_DIFF_{ts}{ext}"
        json_path = f"{root}_DIFF_{ts}.json"

        self._log(f"[COPY] {base_path} -> {diff_path}")
        shutil.copy2(base_path, diff_path)

        self._meta = {
            "file_a": self.req.file_a,
            "file_b": self.req.file_b,
            "range_a": self.req.range_a,
            "range_b": self.req.range_b,
            "base_file": base,
            "sheet_mode": sheet_mode,
            "compare_formula": bool(getattr(self.req, "compare_formula", False)),
            "compare_shapes": bool(getattr(self.req, "compare_shapes", False)),
        }

        app = None
        book_diff = None
        book_other = None

        try:
            app = xw.App(visible=False, add_book=False)

            book_diff = app.books.open(diff_path)
            book_other = app.books.open(other_path, read_only=True)

            sheet_pairs = self._resolve_sheet_pairs(book_diff, book_other, sheet_mode)

            for sheet_name, sht_diff, sht_other in sheet_pairs:
                self._log(f"[SHEET] {sheet_name}")

                self._diff_cells_core(sht_diff, sht_other, sheet_name)
                self._mark_cells_red_xlwings(sht_diff, sheet_name)

                if bool(getattr(self.req, "compare_shapes", False)):
                    self._diff_shapes_core(sht_diff, sht_other, sheet_name)
                    self._mark_shapes_red(book_diff, sheet_name)

            book_diff.save()

        except Exception as ex:
            self._log_err(f"[ERR] Diff: {ex}")
            raise

        finally:
            try:
                if book_other:
                    book_other.close()
            except Exception:
                pass
            try:
                if book_diff:
                    book_diff.close()
            except Exception:
                pass
            try:
                if app:
                    app.quit()
            except Exception:
                pass

        self._summary = {
            "cell_mod_count": len(self.diff_cells),
            "shape_diff_count": len(self.diff_shapes),
            "base_file": base,
            "sheet_mode": sheet_mode,
        }

        self._write_json(json_path)

        self._log(f"[OK] 出力ファイル: {diff_path}")
        self._log("=== Diff完了 ===")
        return diff_path

    # -------------------------------------------------
    # sheet resolve
    # -------------------------------------------------
    def _resolve_sheet_pairs(self, book_diff: xw.Book, book_other: xw.Book, mode: str):
        pairs = []

        if mode == "index":
            n = min(len(book_diff.sheets), len(book_other.sheets))
            for i in range(n):
                pairs.append(
                    (
                        book_diff.sheets[i].name,
                        book_diff.sheets[i],
                        book_other.sheets[i],
                    )
                )

        elif mode == "name":
            other_names = set(book_other.sheets.names)
            for name in book_diff.sheets.names:
                if name in other_names:
                    pairs.append(
                        (
                            name,
                            book_diff.sheets[name],
                            book_other.sheets[name],
                        )
                    )
                else:
                    self.diff_shapes.append({"type": "SHEET_DEL", "sheet": name})

            for name in other_names - set(book_diff.sheets.names):
                self.diff_shapes.append({"type": "SHEET_ADD", "sheet": name})

        return pairs

    # -------------------------------------------------
    # cell diff
    # -------------------------------------------------
    def _diff_cells_core(self, sht_base: xw.Sheet, sht_other: xw.Sheet, sheet: str) -> None:
        base = self._meta["base_file"]

        data_base = self._read_range(
            sht_base,
            self.req.range_b if base == "B" else self.req.range_a,
        )
        data_other = self._read_range(
            sht_other,
            self.req.range_a if base == "B" else self.req.range_b,
        )

        for r in sorted(set(data_base) | set(data_other)):
            row_a = data_base.get(r, {})
            row_b = data_other.get(r, {})
            for c in sorted(set(row_a) | set(row_b)):
                va = row_a.get(c)
                vb = row_b.get(c)
                if va != vb:
                    self.diff_cells.append(
                        {
                            "type": "MOD",
                            "sheet": sheet,
                            "row": r,
                            "col": c,
                            "base": base,
                            "value_a": "" if va is None else str(va),
                            "value_b": "" if vb is None else str(vb),
                        }
                    )

    # -------------------------------------------------
    # shape diff
    # -------------------------------------------------
    def _diff_shapes_core(self, sht_base: xw.Sheet, sht_other: xw.Sheet, sheet: str) -> None:
        shapes_base = self._read_shapes(sht_base)
        shapes_other = self._read_shapes(sht_other)

        for name in sorted(set(shapes_base) | set(shapes_other)):
            a = shapes_base.get(name)
            b = shapes_other.get(name)
            if a != b:
                self.diff_shapes.append(
                    {"type": "SHAPE_MOD", "sheet": sheet, "name": name, "a": a, "b": b}
                )

    # -------------------------------------------------
    # helpers
    # -------------------------------------------------
    def _read_range(self, sht: xw.Sheet, rng: str) -> Dict[int, Dict[int, Any]]:
        area = sht.range(rng)
        values = area.formula if getattr(self.req, "compare_formula", False) else area.value

        out: Dict[int, Dict[int, Any]] = {}
        sr, sc = area.row, area.column

        for ro, row in enumerate(values):
            if not isinstance(row, list):
                row = [row]
            if not any(v is not None for v in row):
                continue
            out[sr + ro] = {sc + c: v for c, v in enumerate(row)}
        return out

    def _read_shapes(self, sht: xw.Sheet) -> Dict[str, Dict[str, Any]]:
        out = {}
        for shp in sht.api.Shapes:
            out[str(shp.Name)] = {
                "top": float(shp.Top),
                "left": float(shp.Left),
                "width": float(shp.Width),
                "height": float(shp.Height),
                "rotation": float(shp.Rotation),
            }
        return out

    # -------------------------------------------------
    # mark
    # -------------------------------------------------
    def _mark_cells_red_xlwings(self, sht: xw.Sheet, sheet: str) -> None:
        for d in self.diff_cells:
            if d["sheet"] != sheet:
                continue
            try:
                cell = sht.cells(d["row"], d["col"])
                cell.api.Interior.Color = 0x6666FF
                cell.api.Borders.Weight = 2
            except Exception:
                pass

    def _mark_shapes_red(self, book: xw.Book, sheet: str) -> None:
        sht = book.sheets[sheet]
        targets = {d["name"] for d in self.diff_shapes if d.get("sheet") == sheet}
        for shp in sht.api.Shapes:
            if str(shp.Name) in targets:
                try:
                    shp.Line.Visible = True
                    shp.Line.ForeColor.RGB = 255
                    shp.Line.Weight = 2
                except Exception:
                    pass

    # -------------------------------------------------
    # json
    # -------------------------------------------------
    def _write_json(self, path: str) -> None:
        payload = {
            "meta": self._meta,
            "summary": self._summary,
            "diff_cells": self.diff_cells,
            "diff_shapes": self.diff_shapes,
        }
        with open(path, "w", encoding="utf-8") as f:
            json.dump(payload, f, ensure_ascii=False, indent=2)


def run_diff(req: DiffRequest, ctx, logger, append_log: LogFn) -> str:
    return ExcelDiffService(req, logger, append_log).run()
