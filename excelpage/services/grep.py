import os
import xlwings as xw
from typing import Tuple, List
from models.dto import GrepRequest, LogFn
from utils.search_utils import compile_matcher

EXCEL_EXTS = (".xlsx", ".xlsm", ".xlsb", ".xls")

def _find_excel_files(root: str) -> List[str]:
    hits = []
    for dp, _, fns in os.walk(root):
        for fn in fns:
            if fn.lower().endswith(EXCEL_EXTS):
                hits.append(os.path.join(dp, fn))
    return hits

def run_grep(req: GrepRequest, ctx, logger, append_log: LogFn) -> Tuple[str, int]:
    append_log("=== Grep開始 ===")
    if not os.path.isdir(req.root_dir):
        raise ValueError(f"ディレクトリが存在しません: {req.root_dir}")

    files = _find_excel_files(req.root_dir)
    total = 0
    matcher = compile_matcher(req.keyword, req.use_regex, req.ignore_case)

    for path in files:
        app = None
        book = None
        try:
            app = xw.App(visible=False, add_book=False)
            book = app.books.open(path, read_only=True)
            for sht in book.sheets:
                vr = sht.used_range
                vals = vr.value
                if vals is None:
                    continue
                if not isinstance(vals, list):
                    vals = [[vals]]
                for r, row in enumerate(vals, start=vr.row):
                    if not isinstance(row, list):
                        row = [row]
                    for c, v in enumerate(row, start=vr.column):
                        if matcher(v):
                            total += 1
                            append_log(f"[HIT] {os.path.basename(path)}[{sht.name}!R{r}C{c}] {str(v)[:60]}")
        except Exception as e:
            append_log(f"[WARN] Grep失敗: {path} ({e})")
        finally:
            try:
                if book:
                    book.close()
            except Exception:
                pass
            try:
                if app:
                    app.kill()
            except Exception:
                pass

    return (req.root_dir, total)
