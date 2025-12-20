# excel_transfer/services/count.py
import os
import re
import xlwings as xw
from typing import Tuple, List
from models.dto import CountRequest, LogFn
from openpyxl.utils import column_index_from_string, get_column_letter

_A1 = re.compile(r"^\$?([A-Z]+)\$?(\d+)$", re.I)

def _parse_a1(a1: str) -> Tuple[int,int]:
    m = _A1.match(a1.strip())
    if not m:
        raise ValueError(f"無効なセル形式: {a1}")
    c, r = m.group(1), int(m.group(2))
    return r, column_index_from_string(c.upper())

def _a1(r: int, c: int) -> str:
    return f"{get_column_letter(c)}{r}"

def _is_empty(v) -> bool:
    if v is None:
        return True
    if isinstance(v, str) and v.strip() == "":
        return True
    return False

def _count_scan(sht: xw.Sheet, r0: int, c0: int, direction: str, tolerate_blanks: int, append_log: LogFn) -> Tuple[int, int]:
    """
    逐次スキャン。連続空白が tolerate_blanks を超えたら停止。
    戻り値: (長さ, 空白検出回数)
    """
    blanks_run = 0
    count = 0
    warnings = 0
    r, c = r0, c0
    while True:
        v = sht.range((r, c)).value
        if _is_empty(v):
            blanks_run += 1
            warnings += 1 if blanks_run == 1 else 0  # 空白始点を警告とカウント
            if blanks_run > tolerate_blanks:
                break
        else:
            blanks_run = 0
        count += 1
        if direction == "row":
            c += 1
        else:
            r += 1
    return count - 1, warnings

def _count_jump(sht: xw.Sheet, r0: int, c0: int, direction: str, tolerate_blanks: int, append_log: LogFn) -> Tuple[int, int]:
    """
    高速（ジャンプ）。tolerate_blanks==0 のときは End キー相当。
    >0 のときは、End 到達後に前進しつつ空白許容分を吸収（簡易）。
    """
    if direction == "row":
        end = sht.api.Cells(r0, c0).End(xw.constants.XlDirection.xlToRight)
        last_c = end.Column
        if tolerate_blanks > 0:
            c = last_c + 1
            blanks_run = 0
            while True:
                v = sht.range((r0, c)).value
                if _is_empty(v):
                    blanks_run += 1
                    if blanks_run > tolerate_blanks:
                        break
                else:
                    blanks_run = 0
                    last_c = c
                c += 1
        length = last_c - c0 + 1
        warnings = 0  # 空白検出件数は scan より粗く扱う
        return max(0, length), warnings
    else:
        end = sht.api.Cells(r0, c0).End(xw.constants.XlDirection.xlDown)
        last_r = end.Row
        if tolerate_blanks > 0:
            r = last_r + 1
            blanks_run = 0
            while True:
                v = sht.range((r, c0)).value
                if _is_empty(v):
                    blanks_run += 1
                    if blanks_run > tolerate_blanks:
                        break
                else:
                    blanks_run = 0
                    last_r = r
                r += 1
        length = last_r - r0 + 1
        warnings = 0
        return max(0, length), warnings

def run_count(req: CountRequest, ctx, logger, append_log: LogFn) -> str:
    append_log("=== Count開始 ===")
    results = []
    for path in req.files:
        if not os.path.exists(path):
            append_log(f"[ERR] ファイルなし: {path}")
            continue

        app = None
        book = None
        try:
            app = xw.App(visible=False, add_book=False)
            book = app.books.open(path, read_only=True)

            sht = book.sheets[req.sheet] if req.sheet else book.sheets[0]
            r0, c0 = _parse_a1(req.start_cell)

            if req.mode == "scan":
                length, warns = _count_scan(sht, r0, c0, req.direction, req.tolerate_blanks, append_log)
            else:
                length, warns = _count_jump(sht, r0, c0, req.direction, req.tolerate_blanks, append_log)

            results.append((os.path.basename(path), sht.name, req.direction, _a1(r0, c0), length, warns))
            if warns > 0:
                append_log(f"[WARN] {os.path.basename(path)}:{sht.name} {req.start_cell} で空白を検出（{warns}件）")
        except Exception as e:
            append_log(f"[ERR] Count失敗: {path} ({e})")
        finally:
            try:
                if book:
                    book.close()  # ← save引数なし
            except Exception:
                pass
            try:
                if app:
                    app.kill()
            except Exception:
                pass

    # サマリ出力
    for fn, sh, d, start, length, warns in results:
        append_log(f"[OK] {fn}:{sh} {start}→{d} 長さ={length}（空白警告={warns}）")

    return "count_done"
