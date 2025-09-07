# excel_transfer/utils/excel.py
import os, glob
from pathlib import Path
from typing import List
import xlwings as xw

def open_app():
    return xw.App(visible=False, add_book=False)

def safe_kill(app):
    try:
        app.kill()
    except Exception:
        pass

def list_excel_files(root: Path) -> List[Path]:
    exts = ("*.xlsx","*.xlsm","*.xlsb","*.xls")
    files: List[Path] = []
    for ext in exts:
        files += list(root.rglob(ext))
    return files

def used_range_2d_values(sheet: xw.Sheet, as_formula=False):
    if sheet is None:
        return []
    ur = sheet.used_range
    if ur is None:
        return []
    vals = ur.formula if as_formula else ur.value
    # xlwings: 単セルはスカラになることがある→2Dへ
    if not isinstance(vals, list) or (vals and not isinstance(vals[0], list)):
        return [[vals]]
    return vals

def normalize_2d(vals):
    if not vals:
        return []
    # None→""、数値/日付はそのまま、式は文字列に
    out = []
    for row in vals:
        rr = []
        if not isinstance(row, list):
            row = [row]
        for v in row:
            if v is None:
                rr.append("")
            else:
                rr.append(str(v) if isinstance(v, str) else v)
        out.append(rr)
    return out
