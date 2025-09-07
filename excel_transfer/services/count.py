# excel_transfer/services/count.py
import os, csv, re
from typing import Tuple
from models.dto import CountRequest, LogFn
from utils.excel import open_app, safe_kill

_A1_RE = re.compile(r"^([A-Z]+)(\d+)$", re.I)

def _parse_a1(a1: str) -> Tuple[int,int]:
    m = _A1_RE.match(a1.strip())
    if not m: raise ValueError(f"無効なセル: {a1}")
    col_letters, row = m.groups()
    col = 0
    for ch in col_letters.upper():
        col = col*26 + (ord(ch)-64)
    return int(row), int(col)

def _is_empty(v) -> bool:
    if v is None: return True
    if isinstance(v, str) and v.strip() == "": return True
    return False

def run_count(req: CountRequest, ctx, logger, append_log: LogFn) -> str:
    if not req.files: raise ValueError("ファイルが指定されていません。")
    if req.direction not in ("row","col"): raise ValueError("direction は row/col のみ")

    out_csv = os.path.join(ctx.output_dir, "contiguous_count.csv")
    rows = []
    app = open_app()
    try:
        for f in req.files:
            if not os.path.isfile(f):
                append_log(f"[WARN] ファイル無し: {f}")
                continue
            try:
                wb = app.books.open(f)
            except Exception as e:
                append_log(f"[WARN] オープン失敗: {f} ({e})")
                continue
            try:
                sh = wb.sheets[req.sheet] if req.sheet and req.sheet in [s.name for s in wb.sheets] else wb.sheets[0]
                r0, c0 = _parse_a1(req.start_cell)
                v0 = sh.range(r0, c0).value
                if _is_empty(v0):
                    rows.append([f, sh.name, req.start_cell, req.direction, 0, 0, f"{req.start_cell} is empty"])
                    append_log(f"[WARN] {os.path.basename(f)}:{sh.name} {req.start_cell} が空です")
                    continue

                if req.mode == "jump":
                    start_rng = sh.range(r0, c0)
                    if req.direction == "row":
                        end_rng = start_rng.end("right")
                        contiguous = end_rng.column - c0 + 1
                        blank_run = 0
                        next_col = end_rng.column + 1
                        used_last_col = sh.used_range.last_cell.column
                        if next_col <= used_last_col and _is_empty(sh.range(r0, next_col).value):
                            to_rng = sh.range(r0, next_col).end("right")
                            blank_run = max(0, to_rng.column - next_col)
                        rows.append([f, sh.name, req.start_cell, req.direction, contiguous, blank_run, ""])
                    else:
                        end_rng = start_rng.end("down")
                        contiguous = end_rng.row - r0 + 1
                        blank_run = 0
                        next_row = end_rng.row + 1
                        used_last_row = sh.used_range.last_cell.row
                        if next_row <= used_last_row and _is_empty(sh.range(next_row, c0).value):
                            to_rng = sh.range(next_row, c0).end("down")
                            blank_run = max(0, to_rng.row - next_row)
                        rows.append([f, sh.name, req.start_cell, req.direction, contiguous, blank_run, ""])
                else:
                    # 逐次スキャン（許容空白つき）
                    count = 1
                    blank_run_len = 0
                    r, c = r0, c0
                    while True:
                        if req.direction == "row":
                            c += 1
                        else:
                            r += 1
                        val = sh.range(r, c).value
                        if _is_empty(val):
                            # 空白連続長
                            local = 1
                            rr, cc = r, c
                            while True:
                                if req.direction == "row": cc += 1
                                else: rr += 1
                                v2 = sh.range(rr, cc).value
                                if _is_empty(v2): local += 1; continue
                                break
                            blank_run_len = local
                            if local <= req.tolerate_blanks:
                                count += local + 1
                                r, c = (rr, cc)
                                continue
                            else:
                                warn_addr = sh.range(r, c).get_address(False, False)
                                rows.append([f, sh.name, req.start_cell, req.direction, count, blank_run_len,
                                             f"empty run {blank_run_len} at {warn_addr} (> tolerate {req.tolerate_blanks})"])
                                break
                        else:
                            count += 1

                    if len(rows)==0 or rows[-1][0]!=f or rows[-1][2]!=req.start_cell or rows[-1][3]!=req.direction:
                        rows.append([f, sh.name, req.start_cell, req.direction, count, blank_run_len, ""])

                append_log(f"計測: {os.path.basename(f)}:{sh.name} {req.start_cell} -> OK")
            finally:
                wb.close(save=False)
    finally:
        safe_kill(app)

    with open(out_csv, "w", newline="", encoding="utf-8-sig") as fw:
        w = csv.writer(fw)
        w.writerow(["file","sheet","start_cell","direction","contiguous_count","blank_run_len","warning"])
        w.writerows(rows)
    return out_csv
