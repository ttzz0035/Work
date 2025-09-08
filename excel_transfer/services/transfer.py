# excel_transfer/services/transfer.py
import os, csv, shutil, datetime, re, gc
import xlwings as xw
from typing import Dict, Tuple, List, Optional, Set
from openpyxl.utils import column_index_from_string, get_column_letter
from models.dto import TransferRequest, LogFn

# A1セル / A1:A1 範囲（$/小文字許容）
_A1_RE_CELL  = re.compile(r"^\$?([A-Z]+)\$?(\d+)$", re.I)
_A1_RE_RANGE = re.compile(r"^\$?([A-Z]+)\$?(\d+)\s*:\s*\$?([A-Z]+)\$?(\d+)$", re.I)


# ==========
# ユーティリティ
# ==========
def _backup_file(file_path, logger):
    if os.path.exists(file_path):
        ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        backup_path = f"{os.path.splitext(file_path)[0]}_backup_{ts}.xlsx"
        shutil.copy2(file_path, backup_path)
        if logger:
            logger.info(f"バックアップ作成: {backup_path}")
        return backup_path
    return ""


def _a1(row: int, col: int) -> str:
    return f"{get_column_letter(col)}{row}"


def _parse_cell(a1: str) -> Tuple[int, int]:
    m = _A1_RE_CELL.match(a1.strip())
    if not m:
        raise ValueError(f"無効なセル形式: {a1}（例: A1 / $B$3）")
    col, row = m.group(1), m.group(2)
    return int(row), column_index_from_string(col.upper())


def _parse_ref(a1_or_range: str) -> Tuple[Tuple[int,int], Optional[Tuple[int,int]]]:
    """
    単セル: ((r1,c1), None)
    範囲  : ((r1,c1), (r2,c2))  ※r1<=r2, c1<=c2 に正規化
    """
    s = a1_or_range.strip()
    m2 = _A1_RE_RANGE.match(s)
    if m2:
        r1, c1 = int(m2.group(2)), column_index_from_string(m2.group(1).upper())
        r2, c2 = int(m2.group(4)), column_index_from_string(m2.group(3).upper())
        r1, r2 = (r1, r2) if r1 <= r2 else (r2, r1)
        c1, c2 = (c1, c2) if c1 <= c2 else (c2, c1)
        return (r1, c1), (r2, c2)
    r, c = _parse_cell(s)
    return (r, c), None


def _apply_offset_cell(rc: Tuple[int,int], r_off: int, c_off: int) -> Tuple[int,int]:
    r, c = rc
    return r + int(r_off or 0), c + int(c_off or 0)


def _apply_offset_ref(ref: Tuple[Tuple[int,int], Optional[Tuple[int,int]]], r_off: int, c_off: int):
    (r1,c1), tail = ref
    r1n, c1n = _apply_offset_cell((r1,c1), r_off, c_off)
    if tail is None:
        return (r1n, c1n), None
    (r2,c2) = tail
    r2n, c2n = _apply_offset_cell((r2,c2), r_off, c_off)
    # 正規化
    r1n, r2n = (r1n, r2n) if r1n <= r2n else (r2n, r1n)
    c1n, c2n = (c1n, c2n) if c1n <= c2n else (c2n, c1n)
    return (r1n, c1n), (r2n, c2n)


def _to_a1_range(ref: Tuple[Tuple[int,int], Optional[Tuple[int,int]]]) -> str:
    (r1,c1), tail = ref
    if tail is None:
        return _a1(r1, c1)
    (r2,c2) = tail
    return f"{_a1(r1,c1)}:{_a1(r2,c2)}"


def _check_oor(ref: Tuple[Tuple[int,int], Optional[Tuple[int,int]]]) -> bool:
    """行<1 or 列<1 を含むか"""
    (r1,c1), tail = ref
    if r1 < 1 or c1 < 1: return True
    if tail:
        r2,c2 = tail
        if r2 < 1 or c2 < 1: return True
    return False


def _log(append_log: LogFn, level: str, code: str, **fields):
    parts: List[str] = [f"[{level.upper()}][{code}]"]
    for k, v in fields.items():
        parts.append(f"{k}={v}")
    append_log(" ".join(parts))


def _fmt_rc(ref: Tuple[Tuple[int,int], Optional[Tuple[int,int]]]) -> str:
    """A1化せず、常に (r,c) 表現で返す（範囲外の負数でも安全）"""
    (r1,c1), tail = ref
    if tail is None:
        return f"(r:{r1},c:{c1})"
    (r2,c2) = tail
    return f"(r:{r1},c:{c1})~(r:{r2},c:{c2})"


def _calc_with_policy(
    a1_or_range: str, r_off: int, c_off: int, policy: str, role: str, job: dict,
    append_log: LogFn, job_id: int, csv_row: int
) -> Tuple[Tuple[Tuple[int,int], Optional[Tuple[int,int]]], bool]:
    """
    返り値: (ref, skipped)
      - ref: ((r1,c1), None|((r2,c2)))（offset適用済み）
      - skipped: True の場合、このジョブをスキップ
    """
    base = _parse_ref(a1_or_range)
    ref = _apply_offset_ref(base, r_off, c_off)
    if _check_oor(ref):
        # ★ A1化しないで安全に出力
        _log(append_log, "WARN", "OOR",
             job=job_id, csv_row=csv_row, role=role, base=a1_or_range,
             off=f"(r:{r_off},c:{c_off})", new=_fmt_rc(ref),
             src=f"{job.get('source_file')}[{job.get('source_sheet')}!{job.get('source_cell')}]",
             dst=f"{job.get('destination_file')}[{job.get('destination_sheet')}!{job.get('destination_cell')}]")
        if policy == "skip":
            _log(append_log, "WARN", "OOR_SKIP", job=job_id, csv_row=csv_row)
            return ref, True
        elif policy == "error":
            # ★ ここでも A1 文字列は作らずに数値で説明
            raise ValueError(
                f"無効なセル（オフセットで範囲外）: {a1_or_range} "
                f"[row_offset={r_off}, col_offset={c_off}] → {_fmt_rc(ref)}"
            )
        else:
            # clamp（現UI未露出）
            (r1,c1), tail = ref
            r1, c1 = max(1,r1), max(1,c1)
            if tail:
                r2,c2 = tail
                r2, c2 = max(1,r2), max(1,c2)
                ref = ( (r1,c1), (r2,c2) )
            else:
                ref = ( (r1,c1), None )
            _log(append_log, "WARN", "OOR_CLAMP", job=job_id, csv_row=csv_row, to=_to_a1_range(ref))
    return ref, False


# ==========
# 集約・矩形化ヘルパ
# ==========
def _cells_from_ref(ref: Tuple[Tuple[int,int], Optional[Tuple[int,int]]]) -> List[Tuple[int,int]]:
    (r1,c1), tail = ref
    if tail is None:
        return [(r1,c1)]
    (r2,c2) = tail
    out = []
    for r in range(r1, r2+1):
        for c in range(c1, c2+1):
            out.append((r,c))
    return out


def _rect_of_cells(cells: Set[Tuple[int,int]]) -> Tuple[Tuple[int,int], Tuple[int,int], bool]:
    """
    与えられたセル集合が完全に矩形充填しているかを判定。
    戻り値: ((r1,c1),(r2,c2), is_full)
      - is_full=True の場合、cells は (r1..r2)×(c1..c2) を完全に含む
    """
    rows = [r for r,_ in cells]
    cols = [c for _,c in cells]
    r1, r2 = min(rows), max(rows)
    c1, c2 = min(cols), max(cols)
    expected = (r2 - r1 + 1) * (c2 - c1 + 1)
    is_full = (len(cells) == expected)
    return (r1,c1), (r2,c2), is_full


# ==========
# メイン処理
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
        try:
            # 設定CSV読み込み（上から順に処理）
            try:
                with open(csv_path, newline="", encoding="utf-8") as f:
                    jobs = list(csv.DictReader(f))
            except UnicodeDecodeError:
                with open(csv_path, newline="", encoding="shift-jis") as f:
                    jobs = list(csv.DictReader(f))
            if not jobs:
                raise ValueError("転記対象がありません。")

            # 先に宛先ブックのバックアップ
            involved = set(job["destination_file"] for job in jobs)
            for file in involved:
                full_path = os.path.join(ctx.base_dir, file)
                _backup_file(full_path, logger)
                _log(append_log, "INFO", "BACKUP", path=full_path)

            # ---------- 準備：ブック/シートオープン、参照解析 ----------
            parsed_items = []  # (job_id, csv_row, src_path, dst_path, src_sheet, dst_sheet, src_ref, dst_ref, job)
            for job_id, job in enumerate(jobs, start=1):
                csv_row = job_id + 1  # header=1

                src_path = os.path.join(ctx.base_dir, job["source_file"])
                dst_path = os.path.join(ctx.base_dir, job["destination_file"])

                # open source
                if src_path not in workbooks:
                    if not os.path.exists(src_path):
                        msg = f"転記元が存在しません: {src_path}"
                        if req.out_of_range_mode == "skip":
                            _log(append_log, "ERROR", "OPEN_SRC_MISSING", job=job_id, csv_row=csv_row, path=src_path)
                            continue
                        else:
                            raise FileNotFoundError(msg)
                    app = xw.App(visible=False, add_book=False)
                    workbooks[src_path] = app.books.open(src_path, read_only=False)
                    _log(append_log, "INFO", "OPEN_SRC", job=job_id, csv_row=csv_row, path=src_path)

                # open/create dest
                if dst_path not in workbooks:
                    app = xw.App(visible=False, add_book=False)
                    if os.path.exists(dst_path):
                        workbooks[dst_path] = app.books.open(dst_path, read_only=False)
                        _log(append_log, "INFO", "OPEN_DST", job=job_id, csv_row=csv_row, path=dst_path)
                    else:
                        workbooks[dst_path] = app.books.add()
                        _log(append_log, "INFO", "CREATE_DST", job=job_id, csv_row=csv_row, path=dst_path)

                # sheet check
                if job["source_sheet"] not in workbooks[src_path].sheet_names:
                    code = "SRC_SHEET_MISSING"
                    if req.out_of_range_mode == "skip":
                        _log(append_log, "ERROR", code, job=job_id, csv_row=csv_row,
                             sheet=job["source_sheet"], file=job["source_file"])
                        continue
                    else:
                        raise ValueError(f"転記元シートが存在しません: {job['source_sheet']} in {job['source_file']}")

                if job["destination_sheet"] not in workbooks[dst_path].sheet_names:
                    workbooks[dst_path].sheets.add(job["destination_sheet"])
                    _log(append_log, "INFO", "CREATE_DST_SHEET", job=job_id, csv_row=csv_row,
                         sheet=job["destination_sheet"], file=job["destination_file"])

                # offsets
                s_ro = int(job.get("source_row_offset", 0) or 0)
                s_co = int(job.get("source_col_offset", 0) or 0)
                d_ro = int(job.get("destination_row_offset", 0) or 0)
                d_co = int(job.get("destination_col_offset", 0) or 0)

                # refs
                src_ref, skip_src = _calc_with_policy(
                    job["source_cell"], s_ro, s_co, req.out_of_range_mode, "src", job, append_log, job_id, csv_row)
                if req.out_of_range_mode == "skip" and skip_src:
                    _log(append_log, "INFO", "SKIP_JOB_SRC", job=job_id, csv_row=csv_row)
                    continue

                dst_ref, skip_dst = _calc_with_policy(
                    job["destination_cell"], d_ro, d_co, req.out_of_range_mode, "dst", job, append_log, job_id, csv_row)
                if req.out_of_range_mode == "skip" and skip_dst:
                    _log(append_log, "INFO", "SKIP_JOB_DST", job=job_id, csv_row=csv_row)
                    continue

                parsed_items.append((
                    job_id, csv_row, src_path, dst_path,
                    job["source_sheet"], job["destination_sheet"],
                    src_ref, dst_ref, job
                ))

            # ---------- 明示範囲はその場で一括転記 ----------
            rest_items = []
            for (job_id, csv_row, src_path, dst_path, src_sh, dst_sh, src_ref, dst_ref, job) in parsed_items:
                if src_ref[1] is not None and dst_ref[1] is not None:
                    (sr1,sc1),(sr2,sc2) = src_ref
                    (dr1,dc1),(dr2,dc2) = dst_ref
                    src_h, src_w = sr2-sr1+1, sc2-sc1+1
                    dst_h, dst_w = dr2-dr1+1, dc2-dc1+1
                    if src_h != dst_h or src_w != dst_w:
                        raise ValueError(f"範囲サイズ不一致: src={_to_a1_range(src_ref)} dst={_to_a1_range(dst_ref)}")
                    src_rng = _to_a1_range(src_ref)
                    dst_rng = _to_a1_range(dst_ref)
                    vals = workbooks[src_path].sheets[src_sh].range(src_rng).value
                    workbooks[dst_path].sheets[dst_sh].range(dst_rng).value = vals
                    _log(append_log, "INFO", "WRITE_RANGE_EXPLICIT",
                         job=job_id, csv_row=csv_row,
                         src=f"{os.path.basename(src_path)}[{src_sh}!{src_rng}]",
                         dst=f"{os.path.basename(dst_path)}[{dst_sh}!{dst_rng}]")
                else:
                    rest_items.append((job_id, csv_row, src_path, dst_path, src_sh, dst_sh, src_ref, dst_ref, job))

            # ---------- 単セル群を宛先ファイル/シートでグループ化 ----------
            from collections import defaultdict
            groups = defaultdict(list)  # key=(dst_path,dst_sh) -> list of item

            for item in rest_items:
                job_id, csv_row, src_path, dst_path, src_sh, dst_sh, src_ref, dst_ref, job = item
                # 片方だけ範囲は定義不正（今回の仕様では不可）
                if (src_ref[1] is None) != (dst_ref[1] is None):
                    raise ValueError(
                        f"定義不正（単セルと範囲の混在）: src={_to_a1_range(src_ref)} dst={_to_a1_range(dst_ref)} "
                        f"@ job={job_id} csv_row={csv_row}"
                    )
                groups[(dst_path, dst_sh)].append(item)

            # ---------- グループごとに「矩形化→形状一致なら一括」「不一致なら行単位」 ----------
            for (dst_path, dst_sh), items in groups.items():
                # すべて単セル？
                all_cell = all(it[6][1] is None and it[7][1] is None for it in items)
                if not all_cell:
                    continue

                # src/dst のセル集合
                src_cells: Set[Tuple[int,int]] = set()
                dst_cells: Set[Tuple[int,int]] = set()
                same_src_origin = True
                base_src_path, base_src_sh = items[0][2], items[0][4]

                for (job_id, csv_row, src_path, _, src_sh, _, src_ref, dst_ref, _) in items:
                    s = _cells_from_ref(src_ref)  # 単セル→1件
                    d = _cells_from_ref(dst_ref)
                    src_cells.update(s)
                    dst_cells.update(d)
                    if src_path != base_src_path or src_sh != base_src_sh:
                        same_src_origin = False

                (sr1,sc1),(sr2,sc2), src_full = _rect_of_cells(src_cells)
                (dr1,dc1),(dr2,dc2), dst_full = _rect_of_cells(dst_cells)
                src_h, src_w = sr2-sr1+1, sc2-sc1+1
                dst_h, dst_w = dr2-dr1+1, dc2-dc1+1

                if src_full and dst_full and (src_h == dst_h) and (src_w == dst_w) and same_src_origin:
                    # --- 矩形形状一致：一括範囲コピー（COM 2回）
                    src_rng = f"{_a1(sr1,sc1)}:{_a1(sr2,sc2)}"
                    dst_rng = f"{_a1(dr1,dc1)}:{_a1(dr2,dc2)}"
                    vals = workbooks[base_src_path].sheets[base_src_sh].range(src_rng).value
                    workbooks[dst_path].sheets[dst_sh].range(dst_rng).value = vals
                    _log(append_log, "INFO", "WRITE_RANGE_AUTO_RECT",
                         src=f"{os.path.basename(base_src_path)}[{base_src_sh}!{src_rng}]",
                         dst=f"{os.path.basename(dst_path)}[{dst_sh}!{dst_rng}]",
                         size=f"{src_h}x{src_w}", cells=len(items))
                    continue

                # --- フォールバック：行単位の連続区間でまとめ書き
                row_buckets: Dict[int, List[Tuple[int,int,Tuple,int,Tuple]]] = defaultdict(list)
                for (job_id, csv_row, src_path, _, src_sh, _, src_ref, dst_ref, job) in items:
                    (dr, dc), _ = dst_ref
                    row_buckets[dr].append((dc, job_id, (src_path, src_sh), csv_row, src_ref))

                for row, lst in row_buckets.items():
                    lst.sort(key=lambda t: (t[2], t[0]))  # by (src_key, dst_col)
                    i = 0
                    while i < len(lst):
                        dst_cols: List[int] = []
                        src_refs: List[Tuple[Tuple[int,int], Optional[Tuple[int,int]]]] = []
                        csv_rows: List[int] = []
                        src_key0 = lst[i][2]
                        j = i
                        while j < len(lst) and lst[j][2] == src_key0:
                            dst_cols.append(lst[j][0])
                            csv_rows.append(lst[j][3])
                            src_refs.append(lst[j][4])
                            j += 1

                        min_col, max_col = min(dst_cols), max(dst_cols)
                        width = max_col - min_col + 1
                        values = [[""] * width]

                        spath, ssh = src_key0
                        for col, sref in zip(dst_cols, src_refs):
                            src_a1 = _to_a1_range(sref)  # 単セル
                            v = workbooks[spath].sheets[ssh].range(src_a1).value
                            v = "" if v is None else v
                            values[0][col - min_col] = v
                            _log(append_log, "INFO", "READ_CELL",
                                 src=f"{os.path.basename(spath)}[{ssh}!{src_a1}]")

                        dst_rng = f"{_a1(row, min_col)}:{_a1(row, max_col)}"
                        workbooks[dst_path].sheets[dst_sh].range(dst_rng).value = values
                        _log(append_log, "INFO", "WRITE_RANGE_FALLBACK_ROW",
                             dst=f"{os.path.basename(dst_path)}[{dst_sh}!{dst_rng}]",
                             count=len(dst_cols))
                        i = j

            # 保存（skip時でも保存まで到達）
            for path, wb in workbooks.items():
                try:
                    wb.save(path)
                    _log(append_log, "INFO", "SAVE_OK", path=path)
                except Exception as e:
                    _log(append_log, "ERROR", "SAVE_FAIL", path=path, error=str(e))

            note = csv_path

        finally:
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
