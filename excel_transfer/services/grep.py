# excel_transfer/services/grep.py
import os
import re
import xlwings as xw
from typing import Tuple, List
from models.dto import GrepRequest, LogFn

EXCEL_EXTS = (".xlsx", ".xlsm", ".xlsb", ".xls")

def _find_excel_files(root: str) -> List[str]:
    hits = []
    for dp, _, fns in os.walk(root):
        for fn in fns:
            if fn.lower().endswith(EXCEL_EXTS):
                hits.append(os.path.join(dp, fn))
    return hits

def _match(text: str, pat: str, use_regex: bool, ignore_case: bool) -> bool:
    if text is None:
        return False
    s = str(text)
    if use_regex:
        flags = re.IGNORECASE if ignore_case else 0
        return re.search(pat, s, flags) is not None
    else:
        if ignore_case:
            return pat.lower() in s.lower()
        return pat in s

def run_grep(req: GrepRequest, ctx, logger, append_log: LogFn) -> Tuple[str, int]:
    append_log("=== Grep開始 ===")
    if not os.path.isdir(req.root_dir):
        raise ValueError(f"ディレクトリが存在しません: {req.root_dir}")

    files = _find_excel_files(req.root_dir)
    total = 0
    for path in files:
        app = None
        book = None
        try:
            app = xw.App(visible=False, add_book=False)
            book = app.books.open(path, read_only=True)
            for sht in book.sheets:
                vr = sht.used_range
                vals = vr.value
                # used_range が None のケースに備える
                if vals is None:
                    continue
                # 1セルの場合はスカラ、行列の場合は2D配列
                if not isinstance(vals, list):
                    vals = [[vals]]
                for r, row in enumerate(vals, start=vr.row):
                    # row がスカラの可能性（1列）にも対応
                    if not isinstance(row, list):
                        row = [row]
                    for c, v in enumerate(row, start=vr.column):
                        if _match(v, req.keyword, req.use_regex, req.ignore_case):
                            total += 1
                            append_log(f"[HIT] {os.path.basename(path)}[{sht.name}!R{r}C{c}] {str(v)[:60]}")
        except Exception as e:
            append_log(f"[WARN] Grep失敗: {path} ({e})")
        finally:
            try:
                if book:
                    book.close()  # ← save引数は渡さない
            except Exception:
                pass
            try:
                if app:
                    app.kill()
            except Exception:
                pass

    return (req.root_dir, total)
