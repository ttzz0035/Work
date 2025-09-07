# excel_transfer/services/transfer.py
import os, csv, shutil, datetime, re, gc
import xlwings as xw
from typing import Dict, Tuple, List
from openpyxl.utils import column_index_from_string, get_column_letter
from models.dto import TransferRequest, LogFn

# A1形式（$/小文字も許容）
_A1_RE = re.compile(r"^\$?([A-Z]+)\$?(\d+)$", re.I)


# ==============
# ユーティリティ
# ==============
def _backup_file(file_path, logger):
    if os.path.exists(file_path):
        ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        backup_path = f"{os.path.splitext(file_path)[0]}_backup_{ts}.xlsx"
        shutil.copy2(file_path, backup_path)
        if logger:
            logger.info(f"バックアップ作成: {backup_path}")
        return backup_path
    return ""


def _parse_a1_with_offset(cell: str, row_offset: int, col_offset: int) -> Tuple[int, int, int, int]:
    """
    A1＋オフセット → (new_row, new_col, base_row, base_col)
    """
    m = _A1_RE.match(cell.strip())
    if not m:
        raise ValueError(f"無効なセル形式: {cell}（例: A1 / $B$3）")
    col_letters, row_str = m.groups()
    base_col = column_index_from_string(col_letters.upper())
    base_row = int(row_str)
    new_col = base_col + int(col_offset or 0)
    new_row = base_row + int(row_offset or 0)
    return new_row, new_col, base_row, base_col


def _a1(row: int, col: int) -> str:
    return f"{get_column_letter(col)}{row}"


def _log(append_log: LogFn, level: str, code: str, **fields):
    """
    構造化ログ: [LEVEL][CODE] key=value ...
    例: [WARN][OOR_SKIP] job=12 csv_row=8 role=dst base=A1 off=(r:3,c:-2) new=(r:4,c:-1)
    """
    parts: List[str] = [f"[{level.upper()}][{code}]"]
    for k, v in fields.items():
        parts.append(f"{k}={v}")
    append_log(" ".join(parts))


def _calc_cell_with_policy(
    cell: str,
    r_off: int,
    c_off: int,
    policy: str,                # "skip" | "error" | （将来: "clamp"）
    role: str,                  # "src" | "dst"
    job: dict,
    append_log: LogFn,
    job_id: int,
    csv_row: int,
) -> str:
    """
    範囲外ポリシーに従って A1 を決定。
    - skip : このジョブをスキップ（警告ログ）→ 空文字を返す
    - error: 例外（処理中止）
    - clamp: 未公開（必要なら将来ON）
    """
    new_row, new_col, base_row, base_col = _parse_a1_with_offset(cell, r_off, c_off)

    if new_row < 1 or new_col < 1:
        _log(
            append_log, "WARN", "OOR",
            job=job_id, csv_row=csv_row, role=role,
            base=cell, off=f"(r:{r_off},c:{c_off})",
            new=f"(r:{new_row},c:{new_col})",
            src=f"{job.get('source_file')}[{job.get('source_sheet')}!{job.get('source_cell')}]",
            dst=f"{job.get('destination_file')}[{job.get('destination_sheet')}!{job.get('destination_cell')}]",
        )
        if policy == "skip":
            _log(append_log, "WARN", "OOR_SKIP", job=job_id, csv_row=csv_row)
            return ""  # スキップ指示
        elif policy == "error":
            # humanメッセージも併記
            raise ValueError(
                f"無効なセル（オフセットで範囲外）: {cell} "
                f"[row_offset={r_off}, col_offset={c_off}] → 行{new_row}, 列{new_col}"
            )
        else:
            # clamp（今はUI非公開）
            new_row = max(1, new_row)
            new_col = max(1, new_col)
            fixed = _a1(new_row, new_col)
            _log(append_log, "WARN", "OOR_CLAMP", job=job_id, csv_row=csv_row, to=fixed)
            return fixed

    return _a1(new_row, new_col)


# ==============
# メインロジック
# ==============
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
            # 定義CSV読み込み
            try:
                with open(csv_path, newline="", encoding="utf-8") as f:
                    jobs = list(csv.DictReader(f))
            except UnicodeDecodeError:
                with open(csv_path, newline="", encoding="shift-jis") as f:
                    jobs = list(csv.DictReader(f))
            if not jobs:
                raise ValueError("転記対象がありません。")

            # バックアップ
            involved = set(job["destination_file"] for job in jobs)
            for file in involved:
                full_path = os.path.join(ctx.base_dir, file)
                _backup_file(full_path, logger)
                _log(append_log, "INFO", "BACKUP", path=full_path)

            # 補助（座標⇔A1）
            from collections import defaultdict

            def cell_to_coord(cell):
                m = re.match(r"([A-Z]+)(\d+)", cell, re.I)
                if m:
                    col_letter, row = m.groups()
                    from openpyxl.utils import column_index_from_string
                    return int(row), column_index_from_string(col_letter)
                return None

            def coord_to_cell(row, col):
                from openpyxl.utils import get_column_letter
                return f"{get_column_letter(col)}{row}"

            # 行単位でまとめる（dst_file, dst_sheet, row）
            grouped = defaultdict(list)

            # job連番・CSV行番号（ヘッダ=1行目 → データ開始=2行目）
            for job_id, job in enumerate(jobs, start=1):
                csv_row = job_id + 1  # headerを1として+1
                src_path = os.path.join(ctx.base_dir, job["source_file"])
                dst_path = os.path.join(ctx.base_dir, job["destination_file"])

                # ブックOpen（元/先）
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

                if dst_path not in workbooks:
                    app = xw.App(visible=False, add_book=False)
                    if os.path.exists(dst_path):
                        workbooks[dst_path] = app.books.open(dst_path, read_only=False)
                        _log(append_log, "INFO", "OPEN_DST", job=job_id, csv_row=csv_row, path=dst_path)
                    else:
                        workbooks[dst_path] = app.books.add()
                        _log(append_log, "INFO", "CREATE_DST", job=job_id, csv_row=csv_row, path=dst_path)

                # シート存在確認・作成
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

                # オフセット
                src_row_offset = int(job.get("source_row_offset", 0) or 0)
                src_col_offset = int(job.get("source_col_offset", 0) or 0)
                dst_row_offset = int(job.get("destination_row_offset", 0) or 0)
                dst_col_offset = int(job.get("destination_col_offset", 0) or 0)

                # A1 + offset 計算（ポリシー適用）
                src_cell = _calc_cell_with_policy(
                    job["source_cell"], src_row_offset, src_col_offset,
                    req.out_of_range_mode, "src", job, append_log, job_id, csv_row
                )
                if req.out_of_range_mode == "skip" and not src_cell:
                    _log(append_log, "INFO", "SKIP_JOB_SRC", job=job_id, csv_row=csv_row)
                    continue

                dst_cell = _calc_cell_with_policy(
                    job["destination_cell"], dst_row_offset, dst_col_offset,
                    req.out_of_range_mode, "dst", job, append_log, job_id, csv_row
                )
                if req.out_of_range_mode == "skip" and not dst_cell:
                    _log(append_log, "INFO", "SKIP_JOB_DST", job=job_id, csv_row=csv_row)
                    continue

                dst_coord = cell_to_coord(dst_cell)
                if dst_coord:
                    grouped[(dst_path, job["destination_sheet"], dst_coord[0])].append(
                        (dst_coord[1], job_id, csv_row, job, src_cell, dst_cell)
                    )
                    _log(append_log, "INFO", "GROUP", job=job_id, csv_row=csv_row,
                         dst_file=job["destination_file"], dst_sheet=job["destination_sheet"], row=dst_coord[0])

            # まとめて書き込み（行ベクトル）
            for (dst_path, dst_sheet, row), cell_jobs in grouped.items():
                cell_jobs.sort()  # col昇順
                values = []
                for col, job_id, csv_row, job, src_cell, dst_cell in cell_jobs:
                    src_path = os.path.join(ctx.base_dir, job["source_file"])
                    value = workbooks[src_path].sheets[job["source_sheet"]].range(src_cell).value or ""
                    values.append(value)
                    _log(append_log, "INFO", "READ_CELL", job=job_id, csv_row=csv_row,
                         src=f"{job['source_file']}[{job['source_sheet']}!{src_cell}]")

                start_col = cell_jobs[0][0]
                end_col = cell_jobs[-1][0]
                rng = f"{_a1(row, start_col)}:{_a1(row, end_col)}"
                workbooks[dst_path].sheets[dst_sheet].range(rng).value = [values]
                _log(append_log, "INFO", "WRITE_RANGE",
                     dst_file=os.path.basename(dst_path), dst_sheet=dst_sheet, range=rng, count=len(values))

                # 個別ログ
                for (col, job_id, csv_row, job, src_cell, dst_cell), v in zip(cell_jobs, values):
                    _log(append_log, "INFO", "WRITE_CELL", job=job_id, csv_row=csv_row,
                         src=f"{job['source_file']}[{job['source_sheet']}!{src_cell}]",
                         dst=f"{job['destination_file']}[{job['destination_sheet']}!{dst_cell}]")

            # 保存
            for path, wb in workbooks.items():
                try:
                    wb.save(path)
                    _log(append_log, "INFO", "SAVE_OK", path=path)
                except Exception as e:
                    _log(append_log, "ERROR", "SAVE_FAIL", path=path, error=str(e))

            note = csv_path

        finally:
            # 後始末
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
