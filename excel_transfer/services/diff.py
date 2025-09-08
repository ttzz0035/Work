# excel_transfer/services/diff.py
import os
import xlwings as xw
from typing import Tuple, Dict, Any, List
from models.dto import DiffRequest, LogFn

def _sheet_or_first(book: xw.Book, name: str):
    if name:
        return book.sheets[name]
    return book.sheets[0]

def _read_sheet_to_dict(sht: xw.Sheet, key_cols: List[str], compare_formula: bool) -> Dict[Tuple, Dict[str, Any]]:
    """
    1行目をヘッダとして、dictに読み込む簡易実装。
    key_cols がなければ行番号をキーにする。
    """
    used = sht.used_range
    vals = used.value
    if vals is None:
        return {}
    if not isinstance(vals, list):
        vals = [[vals]]

    # ヘッダ抽出
    headers = vals[0] if len(vals) > 0 else []
    if not isinstance(headers, list):
        headers = [headers]
    col_index = {str(h): i for i, h in enumerate(headers)}
    rows = vals[1:]

    out = {}
    for idx, row in enumerate(rows, start=2):
        if not isinstance(row, list):
            row = [row]
        # キー生成
        if key_cols:
            key = tuple(row[col_index.get(k, -1)] for k in key_cols)
        else:
            key = (idx,)  # 行番号
        # 値
        record = {}
        for h, i in col_index.items():
            try:
                cell = sht.range((idx, i + 1))
                v = cell.formula if compare_formula else cell.value
            except Exception:
                v = None
            record[h] = v
        out[key] = record
    return out

def run_diff(req: DiffRequest, ctx, logger, append_log: LogFn) -> str:
    append_log("=== Diff開始 ===")
    if not os.path.exists(req.file_a) or not os.path.exists(req.file_b):
        raise ValueError("比較ファイルが存在しません。")

    app_a = app_b = None
    book_a = book_b = None
    try:
        app_a = xw.App(visible=False, add_book=False)
        app_b = xw.App(visible=False, add_book=False)
        book_a = app_a.books.open(req.file_a, read_only=True)
        book_b = app_b.books.open(req.file_b, read_only=True)

        # 先頭シートで比較（必要なら拡張）
        sht_a = book_a.sheets[0]
        sht_b = book_b.sheets[0]
        append_log(f"比較: {sht_a.name}")

        a_dict = _read_sheet_to_dict(sht_a, req.key_cols, req.compare_formula)
        b_dict = _read_sheet_to_dict(sht_b, req.key_cols, req.compare_formula)

        a_keys = set(a_dict.keys())
        b_keys = set(b_dict.keys())

        only_a = sorted(a_keys - b_keys)
        only_b = sorted(b_keys - a_keys)
        both = sorted(a_keys & b_keys)

        for k in only_a:
            append_log(f"[DEL] {k}")
        for k in only_b:
            append_log(f"[ADD] {k}")

        for k in both:
            ra = a_dict[k]; rb = b_dict[k]
            # 差分検出（簡易）
            diffs = []
            for col in sorted(set(ra.keys()) | set(rb.keys())):
                va, vb = ra.get(col), rb.get(col)
                if va != vb:
                    diffs.append((col, va, vb))
            if diffs:
                append_log(f"[MOD] {k} -> {len(diffs)}差分")
                if req.include_context:
                    for col, va, vb in diffs[:50]:
                        append_log(f"  - {col}: A={va} | B={vb}")

        if req.compare_shapes:
            append_log("[INFO] 形状（図・画像）の比較は未対応です（将来拡張予定）。")

        return "diff_done"

    finally:
        try:
            if book_a:
                book_a.close()  # save渡さない
        except Exception:
            pass
        try:
            if app_a:
                app_a.kill()
        except Exception:
            pass

        try:
            if book_b:
                book_b.close()
        except Exception:
            pass
        try:
            if app_b:
                app_b.kill()
        except Exception:
            pass
