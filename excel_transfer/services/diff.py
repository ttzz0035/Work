from __future__ import annotations

import json
import os
from typing import Any, Dict, List, Tuple

import xlwings as xw

from models.dto import DiffRequest, LogFn
from datetime import datetime


class ExcelDiffService:
    def __init__(self, req: DiffRequest, logger, append_log: LogFn):
        self.req = req
        self.logger = logger
        self.append_log = append_log

        self.diff_cells: List[Tuple[int, int]] = []
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

        # --- base 正規化 ---
        base = str(getattr(self.req, "base_file", "B") or "B").upper()
        if base not in ("A", "B"):
            base = "B"

        # --- paths（常に定義される） ---
        base_path = self.req.file_a if base == "A" else self.req.file_b
        other_path = self.req.file_b if base == "A" else self.req.file_a

        if not base_path or not os.path.exists(base_path):
            raise ValueError(f"invalid base_path: {base_path}")
        if not other_path or not os.path.exists(other_path):
            raise ValueError(f"invalid other_path: {other_path}")

        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        base_root, base_ext = os.path.splitext(base_path)

        diff_path = f"{base_root}_DIFF_{ts}{base_ext}"
        json_path = f"{base_root}_DIFF_{ts}.json"

        self._meta = {
            "file_a": self.req.file_a,
            "file_b": self.req.file_b,
            "range_a": self.req.range_a,
            "range_b": self.req.range_b,
            "base_file": base,
            "compare_formula": bool(getattr(self.req, "compare_formula", False)),
            "compare_shapes": bool(getattr(self.req, "compare_shapes", False)),
        }

        app = None
        book_base = None
        book_other = None

        try:
            self._log(f"[PATH] base_path={base_path}")
            self._log(f"[PATH] other_path={other_path}")

            app = xw.App(visible=False, add_book=False)

            book_base = app.books.open(base_path)
            if book_base is None:
                raise RuntimeError(f"failed to open base excel: {base_path}")

            book_other = app.books.open(other_path, read_only=True)
            if book_other is None:
                raise RuntimeError(f"failed to open other excel: {other_path}")

            sht_base = book_base.sheets[0]
            sht_other = book_other.sheets[0]

            self._diff_cells_core(sht_base, sht_other)
            self._mark_cells_red_xlwings(sht_base)

            if bool(getattr(self.req, "compare_shapes", False)):
                self._diff_shapes_core(sht_base, sht_other)
                self._mark_shapes_red(book_base)

            book_base.save(diff_path)

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
                if book_base:
                    book_base.close()
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
        }

        self._write_json(json_path)

        self._log(f"[OK] 出力ファイル: {diff_path}")
        self._log("=== Diff完了 ===")
        return diff_path

    # -------------------------------------------------
    # cell diff
    # -------------------------------------------------
    def _diff_cells_core(self, sht_base: xw.Sheet, sht_other: xw.Sheet) -> None:
        base = self._meta["base_file"]

        data_base = self._read_range(
            sht_base,
            self.req.range_b if base == "B" else self.req.range_a,
        )
        data_other = self._read_range(
            sht_other,
            self.req.range_a if base == "B" else self.req.range_b,
        )

        rows = sorted(set(data_base) | set(data_other))
        total = len(rows)

        for i, r in enumerate(rows, 1):
            row_a = data_base.get(r, {})
            row_b = data_other.get(r, {})

            for c in sorted(set(row_a) | set(row_b)):
                if row_a.get(c) != row_b.get(c):
                    self.diff_cells.append((r, c))
                    self._log(
                        f"[MOD] cell r={r} c={c} A={row_a.get(c)} B={row_b.get(c)}"
                    )

            if i % 500 == 0 or i == total:
                self._log(f"[PROGRESS] cells {i}/{total}")

    # -------------------------------------------------
    # shape diff (JSON only)
    # -------------------------------------------------
    def _diff_shapes_core(self, sht_base: xw.Sheet, sht_other: xw.Sheet) -> None:
        self._log("[INFO] 図形比較開始")

        shapes_base = self._read_shapes(sht_base)
        shapes_other = self._read_shapes(sht_other)

        for name in sorted(set(shapes_base) | set(shapes_other)):
            a = shapes_base.get(name)
            b = shapes_other.get(name)

            if a is None and b is not None:
                self.diff_shapes.append(
                    {"type": "SHAPE_ADD", "name": name, "a": None, "b": b}
                )
                self._log(f"[SHAPE-ADD] {name}")

            elif a is not None and b is None:
                self.diff_shapes.append(
                    {"type": "SHAPE_DEL", "name": name, "a": a, "b": None}
                )
                self._log(f"[SHAPE-DEL] {name}")

            elif a != b:
                self.diff_shapes.append(
                    {"type": "SHAPE_MOD", "name": name, "a": a, "b": b}
                )
                self._log(f"[SHAPE-MOD] {name}")

    # -------------------------------------------------
    # helpers
    # -------------------------------------------------
    def _read_range(self, sht: xw.Sheet, rng: str) -> Dict[int, Dict[int, Any]]:
        area = sht.range(rng)
        values = area.formula if getattr(self.req, "compare_formula", False) else area.value

        out: Dict[int, Dict[int, Any]] = {}
        start_row = area.row
        start_col = area.column

        self._log(f"[READ] start sheet={sht.name} range={rng}")

        for r_off, row in enumerate(values):
            if not isinstance(row, list):
                row = [row]
            if not any(v is not None for v in row):
                continue

            r = start_row + r_off
            out[r] = {start_col + c: v for c, v in enumerate(row)}

        self._log(f"[READ] done sheet={sht.name}")
        return out

    def _read_shapes(self, sht: xw.Sheet) -> Dict[str, Dict[str, Any]]:
        out: Dict[str, Dict[str, Any]] = {}
        for shp in sht.api.Shapes:
            try:
                out[str(shp.Name)] = {
                    "top": float(shp.Top),
                    "left": float(shp.Left),
                    "width": float(shp.Width),
                    "height": float(shp.Height),
                    "rotation": float(shp.Rotation),
                    "text": getattr(shp.TextFrame.Characters(), "Text", ""),
                }
            except Exception as e:
                self._log(f"[SHAPE-READ-ERR] {e}")
        return out

    # -------------------------------------------------
    # mark (xlwings only)
    # -------------------------------------------------
    def _mark_cells_red_xlwings(self, sht: xw.Sheet) -> None:
        for r, c in self.diff_cells:
            try:
                cell = sht.cells(r, c)
                cell.api.Interior.Color = 0x6666FF  # 赤系
                cell.api.Borders.Weight = 2
            except Exception as e:
                self._log(f"[CELL-MARK-ERR] r={r} c={c} err={e}")

    def _mark_shapes_red(self, book: xw.Book) -> None:
        sht = book.sheets[0]
        targets = {
            d["name"] for d in self.diff_shapes if d.get("type") == "SHAPE_MOD"
        }

        for shp in sht.api.Shapes:
            if str(shp.Name) in targets:
                try:
                    shp.Line.Visible = True
                    shp.Line.ForeColor.RGB = 255  # 赤
                    shp.Line.Weight = 2
                except Exception as e:
                    self._log(f"[SHAPE-MARK-ERR] {e}")

    # -------------------------------------------------
    # json
    # -------------------------------------------------
    def _write_json(self, path: str) -> None:
        payload = {
            "meta": self._meta,
            "summary": self._summary,
            "diff_cells": [
                {
                    "type": "MOD",
                    "mark": {"row": r, "col": c, "base": self._meta["base_file"]},
                }
                for r, c in self.diff_cells
            ],
            "diff_shapes": self.diff_shapes,
        }

        with open(path, "w", encoding="utf-8") as f:
            json.dump(payload, f, ensure_ascii=False, indent=2)

        self._log(f"[OK] JSON出力: {path}")


def run_diff(req: DiffRequest, ctx, logger, append_log: LogFn) -> str:
    return ExcelDiffService(req, logger, append_log).run()
