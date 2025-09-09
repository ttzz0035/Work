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
    """
    - source_cell / destination_cell:
        * A1 もしくは A1:A1 の範囲指定
        * 検索式: A{文字列} / 1{文字列}（完全一致）
          ※ヒットセルを基点に source_row_offset / source_col_offset（または destination_～）を加算
    - 取得は基本1セル。source が範囲 & destination が A1（top-left）の場合は範囲サイズ分を一括貼付
    - out_of_range_mode == "skip" の場合は WARN ログを出して継続、"error" は例外で中止
    """
    if not req.csv_paths:
        raise ValueError("転記定義CSVが指定されていません。")

    note = ""
    for csv_path in req.csv_paths:
        if not os.path.exists(csv_path):
            raise FileNotFoundError(csv_path)

    for csv_path in req.csv_paths:
        workbooks: Dict[str, xw.Book] = {}
        try:
            # 設定CSV読込
            try:
                with open(csv_path, newline="", encoding="utf-8") as f:
                    jobs = list(csv.DictReader(f))
            except UnicodeDecodeError:
                with open(csv_path, newline="", encoding="shift-jis") as f:
                    jobs = list(csv.DictReader(f))
            if not jobs:
                raise ValueError("転記対象がありません。")

            # 宛先のバックアップ
            for file in set(job["destination_file"] for job in jobs):
                full_path = os.path.join(ctx.base_dir, file)
                _backup_file(full_path, logger)
                _log(append_log, "INFO", "BACKUP", path=full_path)

            # ジョブ処理
            for job_id, job in enumerate(jobs, start=1):
                csv_row = job_id + 1

                src_path = os.path.join(ctx.base_dir, job["source_file"])
                dst_path = os.path.join(ctx.base_dir, job["destination_file"])

                # open source
                if src_path not in workbooks:
                    if not os.path.exists(src_path):
                        if req.out_of_range_mode == "skip":
                            _log(append_log, "ERROR", "OPEN_SRC_MISSING", job=job_id, csv_row=csv_row, path=src_path)
                            continue
                        raise FileNotFoundError(f"転記元が存在しません: {src_path}")
                    app = xw.App(visible=False, add_book=False)
                    workbooks[src_path] = app.books.open(src_path, read_only=False)
                    _log(append_log, "INFO", "OPEN_SRC", job=job_id, csv_row=csv_row, path=src_path)

                # open/create destination
                if dst_path not in workbooks:
                    app = xw.App(visible=False, add_book=False)
                    if os.path.exists(dst_path):
                        workbooks[dst_path] = app.books.open(dst_path, read_only=False)
                        _log(append_log, "INFO", "OPEN_DST", job=job_id, csv_row=csv_row, path=dst_path)
                    else:
                        workbooks[dst_path] = app.books.add()
                        _log(append_log, "INFO", "CREATE_DST", job=job_id, csv_row=csv_row, path=dst_path)

                src_book = workbooks[src_path]
                dst_book = workbooks[dst_path]

                # シート存在確認/作成
                src_sheet = job["source_sheet"]
                dst_sheet = job["destination_sheet"]
                if src_sheet not in src_book.sheet_names:
                    if req.out_of_range_mode == "skip":
                        _log(append_log, "ERROR", "SRC_SHEET_MISSING", job=job_id, csv_row=csv_row,
                             sheet=src_sheet, file=os.path.basename(src_path))
                        continue
                    raise ValueError(f"転記元シートが存在しません: {src_sheet}")
                if dst_sheet not in dst_book.sheet_names:
                    dst_book.sheets.add(dst_sheet)
                    _log(append_log, "INFO", "CREATE_DST_SHEET", job=job_id, csv_row=csv_row, sheet=dst_sheet)

                sht_src = src_book.sheets[src_sheet]
                sht_dst = dst_book.sheets[dst_sheet]

                # オフセット
                s_ro = int(job.get("source_row_offset", 0) or 0)
                s_co = int(job.get("source_col_offset", 0) or 0)
                d_ro = int(job.get("destination_row_offset", 0) or 0)
                d_co = int(job.get("destination_col_offset", 0) or 0)

                # 参照解決
                s_expr = (job.get("source_cell") or "").strip()
                d_expr = (job.get("destination_cell") or "").strip()

                try:
                    src_ref, s_skipped = _resolve_any_cell_or_range(
                        sht_src, s_expr, s_ro, s_co, "SRC", append_log, job_id, csv_row, req.out_of_range_mode)
                    if req.out_of_range_mode == "skip" and s_skipped:
                        _log(append_log, "INFO", "SKIP_JOB_SRC", job=job_id, csv_row=csv_row)
                        continue

                    dst_ref, d_skipped = _resolve_any_cell_or_range(
                        sht_dst, d_expr, d_ro, d_co, "DST", append_log, job_id, csv_row, req.out_of_range_mode)
                    if req.out_of_range_mode == "skip" and d_skipped:
                        _log(append_log, "INFO", "SKIP_JOB_DST", job=job_id, csv_row=csv_row)
                        continue
                except Exception as e:
                    if req.out_of_range_mode == "skip":
                        _log(append_log, "ERROR", "RESOLVE_FAIL", job=job_id, csv_row=csv_row, msg=str(e))
                        continue
                    raise

                # 転記
                (sr1, sc1), s_tail = src_ref
                (dr1, dc1), d_tail = dst_ref

                if s_tail is not None and d_tail is None:
                    # source=範囲 / destination=top-left → 範囲サイズ分を一括貼付
                    (sr2, sc2) = s_tail
                    h, w = sr2 - sr1 + 1, sc2 - sc1 + 1
                    dst_rng = f"{_a1(dr1, dc1)}:{_a1(dr1 + h - 1, dc1 + w - 1)}"
                    src_rng = f"{_a1(sr1, sc1)}:{_a1(sr2, sc2)}"
                    vals = sht_src.range(src_rng).value
                    sht_dst.range(dst_rng).value = vals
                    _log(append_log, "INFO", "WRITE_RANGE_SRC_TO_DST_TL",
                         job=job_id, csv_row=csv_row,
                         src=f"{src_sheet}!{src_rng}", dst=f"{dst_sheet}!{dst_rng}", size=f"{h}x{w}")
                elif s_tail is None and d_tail is None:
                    # 単セル → 単セル
                    val = sht_src.range((sr1, sc1)).value
                    sht_dst.range((dr1, dc1)).value = val
                    _log(append_log, "INFO", "WRITE_CELL",
                         job=job_id, csv_row=csv_row,
                         src=f"{src_sheet}!{_a1(sr1, sc1)}",
                         dst=f"{dst_sheet}!{_a1(dr1, dc1)}",
                         val=str(val)[:60] if val is not None else "")
                elif s_tail is not None and d_tail is not None:
                    # 範囲 → 範囲（サイズチェック）
                    (sr2, sc2) = s_tail
                    (dr2, dc2) = d_tail
                    sh, sw = sr2 - sr1 + 1, sc2 - sc1 + 1
                    dh, dw = dr2 - dr1 + 1, dc2 - dc1 + 1
                    if (sh, sw) != (dh, dw):
                        if req.out_of_range_mode == "skip":
                            _log(append_log, "ERROR", "RANGE_SIZE_MISMATCH",
                                 job=job_id, csv_row=csv_row,
                                 src=f"{src_sheet}!{_a1(sr1, sc1)}:{_a1(sr2, sc2)}",
                                 dst=f"{dst_sheet}!{_a1(dr1, dc1)}:{_a1(dr2, dc2)}",
                                 s_size=f"{sh}x{sw}", d_size=f"{dh}x{dw}")
                            continue
                        raise ValueError("範囲サイズ不一致")
                    src_rng = f"{_a1(sr1, sc1)}:{_a1(sr2, sc2)}"
                    dst_rng = f"{_a1(dr1, dc1)}:{_a1(dr2, dc2)}"
                    vals = sht_src.range(src_rng).value
                    sht_dst.range(dst_rng).value = vals
                    _log(append_log, "INFO", "WRITE_RANGE_EXPLICIT",
                         job=job_id, csv_row=csv_row,
                         src=f"{src_sheet}!{src_rng}", dst=f"{dst_sheet}!{dst_rng}", size=f"{sh}x{sw}")
                else:
                    # src=単セル, dst=範囲 → 仕様外
                    if req.out_of_range_mode == "skip":
                        _log(append_log, "ERROR", "INVALID_COMBINATION",
                             job=job_id, csv_row=csv_row,
                             src=f"{src_sheet}!{_a1(sr1, sc1)}", dst=f"{dst_sheet}!{_a1(*d_tail[0])}:{_a1(*d_tail[1])}")
                        continue
                    raise ValueError("定義不正（src=単セル, dst=範囲）")

            # 保存
            for path, wb in workbooks.items():
                try:
                    wb.save(path)
                    _log(append_log, "INFO", "SAVE_OK", path=path)
                except Exception as e:
                    _log(append_log, "ERROR", "SAVE_FAIL", path=path, error=str(e))
            note = csv_path

        finally:
            # 後片付け
            for path, wb in workbooks.items():
                try:
                    app = wb.app
                    wb.close()
                    app.kill()
                except Exception:
                    pass
            del workbooks
            gc.collect()

    return note
