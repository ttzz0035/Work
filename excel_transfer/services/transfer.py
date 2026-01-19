# excel_transfer/services/transfer.py
import os
import csv
import shutil
import datetime
import re
import gc
from typing import Dict, Tuple, Optional

import xlwings as xw
from openpyxl.utils import column_index_from_string, get_column_letter

from models.dto import TransferRequest, LogFn
from utils.search_utils import compile_matcher, find_in_column, find_in_row

# A1 / A1:A1
_A1_RE_CELL  = re.compile(r"^\$?([A-Z]+)\$?(\d+)$", re.I)
_A1_RE_RANGE = re.compile(r"^\$?([A-Z]+)\$?(\d+)\s*:\s*\$?([A-Z]+)\$?(\d+)$", re.I)
# 検索式: A{文字列} もしくは 1{文字列}
_SEARCH_RE   = re.compile(r"^([A-Z]+|\d+)\{(.+)\}$")


# ==========
# ユーティリティ
# ==========
def _a1(row: int, col: int) -> str:
    return f"{get_column_letter(col)}{row}"


def _parse_cell(a1: str) -> Tuple[int, int]:
    m = _A1_RE_CELL.match(a1.strip())
    if not m:
        raise ValueError(f"無効なセル形式: {a1}")
    col, row = m.group(1), m.group(2)
    return int(row), column_index_from_string(col.upper())


def _parse_ref(a1_or_range: str):
    s = a1_or_range.strip()
    m2 = _A1_RE_RANGE.match(s)
    if m2:
        r1, c1 = int(m2.group(2)), column_index_from_string(m2.group(1).upper())
        r2, c2 = int(m2.group(4)), column_index_from_string(m2.group(3).upper())
        # 正規化
        r1, r2 = (r1, r2) if r1 <= r2 else (r2, r1)
        c1, c2 = (c1, c2) if c1 <= c2 else (c2, c1)
        return (r1, c1), (r2, c2)
    r, c = _parse_cell(s)
    return (r, c), None


def _log(append_log: LogFn, level: str, code: str, **fields):
    parts = [f"[{level.upper()}][{code}]"] + [f"{k}={v}" for k, v in fields.items()]
    append_log(" ".join(parts))


def _backup_file(file_path: str, logger):
    if os.path.exists(file_path):
        ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        backup_path = f"{os.path.splitext(file_path)[0]}_backup_{ts}.xlsx"
        shutil.copy2(file_path, backup_path)
        if logger:
            logger.info(f"バックアップ作成: {backup_path}")
        return backup_path
    return ""


# ==========
# 検索式の解決（完全一致）
# ==========
def _resolve_search_cell(sht: xw.Sheet, expr: str, r_off: int, c_off: int,
                         role: str, append_log: LogFn, job_id: int, csv_row: int,
                         skip_policy: str) -> Tuple[Tuple[int, int], bool]:
    """
    expr: 'A{keyword}'（列検索→行特定） or '1{keyword}'（行検索→列特定）
    戻り値: ((row, col), skipped)
    """
    m = _SEARCH_RE.match(expr.strip())
    if not m:
        raise ValueError(f"{role}: 検索書式が不正: {expr}")
    token, keyword = m.group(1), m.group(2)

    matcher = compile_matcher(keyword, use_regex=False, ignore_case=False)
    used = sht.used_range

    if token.isalpha():
        # 列A/B/... を上から検索 → 行を特定
        col = column_index_from_string(token.upper())
        hit_r = find_in_column(sht, token.upper(), matcher)
        if hit_r is None:
            if skip_policy == "skip":
                _log(append_log, "WARN", f"{role}_ROW_NOT_FOUND", job=job_id, csv_row=csv_row, col=token, key=keyword)
                return (1, 1), True
            raise ValueError(f"{role}: 検索ヒットなし（列={token}, '{keyword}'）")
        fr, fc = hit_r + int(r_off or 0), col + int(c_off or 0)
        _log(append_log, "INFO", f"{role}_SEARCH_HIT",
             job=job_id, csv_row=csv_row, base=f"(r:{hit_r},c:{col})",
             offset=f"(r+{r_off},c+{c_off})", final=f"(r:{fr},c:{fc})")
        return (fr, fc), False
    else:
        # 行番号 1/2/... を左から検索 → 列を特定
        row = int(token)
        # 行の存在ガード
        if row < used.row or row > used.last_cell.row:
            if skip_policy == "skip":
                _log(append_log, "WARN", f"{role}_ROW_OOR", job=job_id, csv_row=csv_row, row=row)
                return (1, 1), True
            raise ValueError(f"{role}: 行が範囲外: {row}")

        hit_c = find_in_row(sht, row, matcher)
        if hit_c is None:
            if skip_policy == "skip":
                _log(append_log, "WARN", f"{role}_COL_NOT_FOUND", job=job_id, csv_row=csv_row, row=row, key=keyword)
                return (1, 1), True
            raise ValueError(f"{role}: 検索ヒットなし（行={row}, '{keyword}'）")
        fr, fc = row + int(r_off or 0), hit_c + int(c_off or 0)
        _log(append_log, "INFO", f"{role}_SEARCH_HIT",
             job=job_id, csv_row=csv_row, base=f"(r:{row},c:{hit_c})",
             offset=f"(r+{r_off},c+{c_off})", final=f"(r:{fr},c:{fc})")
        return (fr, fc), False


def _resolve_any_cell_or_range(sht: xw.Sheet, expr: str, r_off: int, c_off: int,
                               role: str, append_log: LogFn, job_id: int, csv_row: int,
                               skip_policy: str):
    """
    A1/範囲 → ((r,c),(r2,c2)|None), skipped=False
    検索式   → ((r,c), None), skipped=bool
    """
    s = (expr or "").strip()
    if _A1_RE_CELL.match(s) or _A1_RE_RANGE.match(s):
        (r1, c1), tail = _parse_ref(s)
        r1, c1 = r1 + int(r_off or 0), c1 + int(c_off or 0)
        if tail:
            r2, c2 = tail
            r2, c2 = r2 + int(r_off or 0), c2 + int(c_off or 0)
            return ((r1, c1), (r2, c2)), False
        return ((r1, c1), None), False

    if _SEARCH_RE.match(s):
        rc, skipped = _resolve_search_cell(sht, s, r_off, c_off, role, append_log, job_id, csv_row, skip_policy)
        return (rc, None), skipped

    raise ValueError(f"{role}: セル指定が不正です: {expr}")


# ==========
# メイン
# ==========
def run_transfer_from_csvs(req: TransferRequest, ctx, logger, append_log: LogFn) -> str:
    if not req.csv_paths:
        raise ValueError("転記定義CSVが指定されていません。")

    note = ""

    for csv_path in req.csv_paths:
        if not os.path.exists(csv_path):
            raise FileNotFoundError(csv_path)

    for csv_path in req.csv_paths:
        workbooks: Dict[str, xw.Book] = {}
        app: Optional[xw.App] = None

        try:
            # ==================================================
            # CSV 読み込み
            # ==================================================
            try:
                with open(csv_path, newline="", encoding="utf-8-sig") as f:
                    jobs = list(csv.DictReader(f))
            except UnicodeDecodeError:
                with open(csv_path, newline="", encoding="shift-jis") as f:
                    jobs = list(csv.DictReader(f))

            if not jobs:
                raise ValueError("転記対象がありません。")

            # ==================================================
            # Excel App を 1 回だけ起動（高速化）
            # ==================================================
            app = xw.App(visible=False, add_book=False)

            app.api.ScreenUpdating = False
            app.api.DisplayAlerts = False
            app.api.EnableEvents = False
            # app.api.Calculation = -4135  # xlCalculationManual

            _log(append_log, "INFO", "EXCEL_APP_START", mode="fast")

            # ==================================================
            # 宛先バックアップ（1回）
            # ==================================================
            for file in set(job["destination_file"] for job in jobs):
                full_path = os.path.join(ctx.base_dir, file)
                _backup_file(full_path, logger)
                _log(append_log, "INFO", "BACKUP", path=full_path)

            # ==================================================
            # ジョブ処理
            # ==================================================
            for job_id, job in enumerate(jobs, start=1):
                csv_row = job_id + 1

                source_file = job.get("source_file")
                if not source_file:
                    _log(append_log, "ERROR", "MISSING_SOURCE_FILE",
                         job=job_id, csv_row=csv_row, job_details=job)
                    raise ValueError(f"source_file が欠落しています: row={csv_row}")

                src_path = os.path.abspath(os.path.join(ctx.base_dir, source_file))
                dst_path = os.path.abspath(os.path.join(ctx.base_dir, job["destination_file"]))

                # --- open source ---
                if src_path not in workbooks:
                    workbooks[src_path] = app.books.open(src_path, read_only=False)
                    _log(append_log, "INFO", "OPEN_SRC", path=src_path)

                # --- open/create destination ---
                if dst_path not in workbooks:
                    if os.path.exists(dst_path):
                        workbooks[dst_path] = app.books.open(dst_path, read_only=False)
                        _log(append_log, "INFO", "OPEN_DST", path=dst_path)
                    else:
                        workbooks[dst_path] = app.books.add()
                        _log(append_log, "INFO", "CREATE_DST", path=dst_path)

                src_book = workbooks[src_path]
                dst_book = workbooks[dst_path]

                src_sheet = job["source_sheet"]
                dst_sheet = job["destination_sheet"]

                if src_sheet not in src_book.sheet_names:
                    raise ValueError(f"転記元シートなし: {src_sheet}")

                if dst_sheet not in dst_book.sheet_names:
                    dst_book.sheets.add(dst_sheet)

                sht_src = src_book.sheets[src_sheet]
                sht_dst = dst_book.sheets[dst_sheet]

                # --- offsets ---
                s_ro = int(job.get("source_row_offset", 0) or 0)
                s_co = int(job.get("source_col_offset", 0) or 0)
                d_ro = int(job.get("destination_row_offset", 0) or 0)
                d_co = int(job.get("destination_col_offset", 0) or 0)

                # --- resolve ---
                src_ref, _ = _resolve_any_cell_or_range(
                    sht_src, job["source_cell"], s_ro, s_co,
                    "SRC", append_log, job_id, csv_row, req.out_of_range_mode)

                dst_ref, _ = _resolve_any_cell_or_range(
                    sht_dst, job["destination_cell"], d_ro, d_co,
                    "DST", append_log, job_id, csv_row, req.out_of_range_mode)

                (sr, sc), _ = src_ref
                (dr, dc), _ = dst_ref

                # --- transfer ---
                val = sht_src.range((sr, sc)).value
                sht_dst.range((dr, dc)).value = val

            # ==================================================
            # 保存
            # ==================================================
            for path, wb in workbooks.items():
                wb.save(path)
                _log(append_log, "INFO", "SAVE_OK", path=path)

            note = csv_path

        finally:
            # ==================================================
            # 後始末（必須）
            # ==================================================
            try:
                if app:
                    app.api.Calculation = -4105  # xlCalculationAutomatic
                    app.api.ScreenUpdating = True
                    app.api.EnableEvents = True
                    app.kill()
            except Exception:
                pass

            workbooks.clear()
            gc.collect()

    return note
