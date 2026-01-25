from __future__ import annotations

import json
import os
import shutil
from datetime import datetime
from typing import Any, Dict, List, Tuple

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
    # logging (EN only, UI-friendly)
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
        self._log("=== DIFF START ===")

        if not self.req.range_a or not self.req.range_b:
            raise ValueError("range_a / range_b is required")

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
            self._log_err(f"[ERR] DIFF FAILED: {ex}")
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

        self._log(f"[OK] DIFF FILE: {diff_path}")
        self._log("=== DIFF END ===")
        return diff_path

    # -------------------------------------------------
    # sheet resolve
    # -------------------------------------------------
    def _resolve_sheet_pairs(self, book_diff: xw.Book, book_other: xw.Book, mode: str):
        pairs: List[Tuple[str, xw.Sheet, xw.Sheet]] = []

        if mode == "index":
            n = min(len(book_diff.sheets), len(book_other.sheets))
            for i in range(n):
                pairs.append(
                    (
                        str(book_diff.sheets[i].name),
                        book_diff.sheets[i],
                        book_other.sheets[i],
                    )
                )

        elif mode == "name":
            diff_map: Dict[str, xw.Sheet] = {}
            other_map: Dict[str, xw.Sheet] = {}

            for sht in book_diff.sheets:
                try:
                    diff_map[str(sht.name)] = sht
                except Exception as e:
                    self._log_err(f"[SHEET] diff name read failed: {e}")

            for sht in book_other.sheets:
                try:
                    other_map[str(sht.name)] = sht
                except Exception as e:
                    self._log_err(f"[SHEET] other name read failed: {e}")

            common = sorted(set(diff_map) & set(other_map))
            only_diff = sorted(set(diff_map) - set(other_map))
            only_other = sorted(set(other_map) - set(diff_map))

            self._log(
                f"[SHEET] name-mode common={len(common)} "
                f"only_diff={len(only_diff)} only_other={len(only_other)}"
            )

            for name in common:
                pairs.append((name, diff_map[name], other_map[name]))

            for name in only_diff:
                self.diff_shapes.append({"type": "SHEET_DEL", "sheet": name})

            for name in only_other:
                self.diff_shapes.append({"type": "SHEET_ADD", "sheet": name})

        return pairs

    # -------------------------------------------------
    # cell diff
    # -------------------------------------------------
    def _diff_cells_core(self, sht_base: xw.Sheet, sht_other: xw.Sheet, sheet: str) -> None:
        base = self._meta["base_file"]

        rng_base = self.req.range_b if base == "B" else self.req.range_a
        rng_other = self.req.range_a if base == "B" else self.req.range_b

        self._log(
            f"[CELL] start sheet={sheet} base={base} "
            f"range_base={rng_base} range_other={rng_other}"
        )

        area_base = sht_base.range(rng_base)
        area_other = sht_other.range(rng_other)

        rows = area_base.rows.count
        cols = area_base.columns.count

        self._log(f"[CELL] read rows={rows} cols={cols}")

        vals_base = area_base.options(ndim=2, empty=None).value
        vals_other = area_other.options(ndim=2, empty=None).value

        sr = area_base.row
        sc = area_base.column

        hit = 0

        for r_off in range(rows):
            for c_off in range(cols):
                va = vals_base[r_off][c_off]
                vb = vals_other[r_off][c_off]

                if va != vb:
                    hit += 1
                    self.diff_cells.append(
                        {
                            "type": "MOD",
                            "sheet": sheet,
                            "row": sr + r_off,
                            "col": sc + c_off,
                            "base": base,
                            "value_a": "" if va is None else str(va),
                            "value_b": "" if vb is None else str(vb),
                        }
                    )

        self._log(f"[CELL] end sheet={sheet} diff_cells={hit}")

    # -------------------------------------------------
    # shape helpers
    # -------------------------------------------------
    def _read_shapes(self, sht: xw.Sheet) -> Dict[str, Dict[str, Any]]:
        out: Dict[str, Dict[str, Any]] = {}
        for shp in sht.api.Shapes:
            try:
                try:
                    text = shp.TextFrame.Characters().Text
                except Exception:
                    text = ""

                out[str(shp.Name)] = {
                    "top": float(shp.Top),
                    "left": float(shp.Left),
                    "width": float(shp.Width),
                    "height": float(shp.Height),
                    "rotation": float(shp.Rotation),
                    "line_color": int(shp.Line.ForeColor.RGB) if shp.Line.Visible else None,
                    "fill_color": int(shp.Fill.ForeColor.RGB) if shp.Fill.Visible else None,
                    "text": text or "",
                }
            except Exception as e:
                self._log_err(f"[SHAPE] read failed sheet={sht.name}: {e}")
        return out

    def _diff_shapes_core(self, sht_base: xw.Sheet, sht_other: xw.Sheet, sheet: str) -> None:
        base_shapes = self._read_shapes(sht_base)
        other_shapes = self._read_shapes(sht_other)

        names_base = set(base_shapes)
        names_other = set(other_shapes)

        self._log(
            f"[SHAPE] start sheet={sheet} base={len(names_base)} other={len(names_other)}"
        )

        for name in sorted(names_base - names_other):
            self.diff_shapes.append({"type": "SHAPE_DEL", "sheet": sheet, "name": name})

        for name in sorted(names_other - names_base):
            self.diff_shapes.append({"type": "SHAPE_ADD", "sheet": sheet, "name": name})

        for name in sorted(names_base & names_other):
            a = base_shapes[name]
            b = other_shapes[name]

            geom_keys = [
                "top",
                "left",
                "width",
                "height",
                "rotation",
                "line_color",
                "fill_color",
            ]
            if any(a[k] != b[k] for k in geom_keys):
                self.diff_shapes.append(
                    {
                        "type": "SHAPE_GEOM",
                        "sheet": sheet,
                        "name": name,
                        "a": {k: a[k] for k in geom_keys},
                        "b": {k: b[k] for k in geom_keys},
                    }
                )

            if (a.get("text") or "") != (b.get("text") or ""):
                self.diff_shapes.append(
                    {
                        "type": "SHAPE_TEXT",
                        "sheet": sheet,
                        "name": name,
                        "text_a": a.get("text", ""),
                        "text_b": b.get("text", ""),
                    }
                )

        self._log(f"[SHAPE] end sheet={sheet} diff_shapes={len(self.diff_shapes)}")

    # -------------------------------------------------
    # mark
    # -------------------------------------------------
    def _mark_shapes_red(self, book: xw.Book, sheet: str) -> None:
        try:
            sht = book.sheets[sheet]
        except Exception as e:
            self._log_err(f"[SHAPE] sheet not found: {sheet} ({e})")
            return

        geom_targets = {
            d["name"]
            for d in self.diff_shapes
            if d["type"] == "SHAPE_GEOM" and d["sheet"] == sheet
        }
        text_targets = {
            d["name"]
            for d in self.diff_shapes
            if d["type"] == "SHAPE_TEXT" and d["sheet"] == sheet
        }

        for shp in sht.api.Shapes:
            name = str(shp.Name)

            if name in geom_targets:
                try:
                    shp.Line.Visible = True
                    shp.Line.ForeColor.RGB = 255
                    shp.Line.Weight = 2
                except Exception as e:
                    self._log_err(f"[SHAPE] geom mark failed {name}: {e}")

            if name in text_targets:
                try:
                    shp.TextFrame.Characters().Font.Color = 255
                except Exception as e:
                    self._log_err(f"[SHAPE] text mark failed {name}: {e}")

    def _mark_cells_red_xlwings(self, sht: xw.Sheet, sheet: str) -> None:
        for d in self.diff_cells:
            if d.get("sheet") != sheet:
                continue
            try:
                cell = sht.range((int(d["row"]), int(d["col"])))
                cell.api.Interior.Color = 0x6666FF
            except Exception as e:
                self._log_err(f"[CELL] mark failed sheet={sheet}: {e}")

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

        self._log(f"[OK] JSON OUTPUT: {path}")


def run_diff(req: DiffRequest, ctx, logger, append_log: LogFn) -> str:
    return ExcelDiffService(req, logger, append_log).run()
