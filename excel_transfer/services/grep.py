# excel_transfer/services/grep.py
import os, csv, re
from pathlib import Path
from typing import Tuple, List
from models.dto import GrepRequest, LogFn
from utils.excel import open_app, safe_kill, list_excel_files, used_range_2d_values, normalize_2d

def run_grep(req: GrepRequest, ctx, logger, append_log: LogFn) -> Tuple[str, int]:
    if not req.root_dir or not os.path.isdir(req.root_dir):
        raise ValueError("検索ルートフォルダが不正です。")
    if not req.keyword:
        raise ValueError("キーワードを入力してください。")

    append_log("=== Grep開始 ===")
    files = list_excel_files(Path(req.root_dir))
    hits: List[tuple] = []
    app = open_app()
    try:
        pattern = None
        if req.use_regex:
            flags = re.IGNORECASE if req.ignore_case else 0
            pattern = re.compile(req.keyword, flags)
        kw = req.keyword.lower() if req.ignore_case else req.keyword

        for f in files:
            try:
                wb = app.books.open(str(f))
            except Exception as e:
                if logger: logger.warning(f"[WARN] オープン失敗: {f} ({e})")
                continue
            try:
                for sh in wb.sheets:
                    vals = normalize_2d(used_range_2d_values(sh, as_formula=False))
                    for r_idx, row in enumerate(vals, start=1):
                        for c_idx, v in enumerate(row, start=1):
                            s = "" if v is None else str(v)
                            if req.use_regex:
                                if pattern.search(s):
                                    hits.append((str(f), sh.name, r_idx, c_idx, s))
                            else:
                                tgt = s.lower() if req.ignore_case else s
                                if kw in tgt:
                                    hits.append((str(f), sh.name, r_idx, c_idx, s))
            finally:
                wb.close(save=False)
                append_log(f"Grep完了: {f.name}")
    finally:
        safe_kill(app)

    out = os.path.join(ctx.output_dir, "grep_results.csv")
    os.makedirs(ctx.output_dir, exist_ok=True)
    with open(out, "w", newline="", encoding="utf-8-sig") as fw:
        w = csv.writer(fw)
        w.writerow(["file","sheet","row","col","value"])
        w.writerows(hits)
    return out, len(hits)
