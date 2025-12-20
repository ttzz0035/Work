import json
from pathlib import Path
from typing import Callable

from logger import get_logger
from services.transfer import run_transfer_from_csvs
from models.dto import TransferRequest

logger = get_logger("JobRunner")


def run_job(job_path: str, append_log: Callable[[str], None] | None = None):
    """
    job.json を読み込み、転記を実行する
    """
    job_path = Path(job_path)

    logger.info("Load job: %s", job_path)

    with job_path.open("r", encoding="utf-8") as f:
        job = json.load(f)

    csv_path = job["rule_csv"]

    req = TransferRequest(
        csv_paths=[csv_path],
        out_of_range_mode=job.get("on_error", "raise"),
    )

    ctx = {
        "source_file": job["source_file"],
        "source_sheet": job["source_sheet"],
        "destination_file": job["destination_file"],
        "destination_sheet": job["destination_sheet"],
    }

    logger.info(
        "Run transfer: %s!%s -> %s!%s",
        ctx["source_file"],
        ctx["source_sheet"],
        ctx["destination_file"],
        ctx["destination_sheet"],
    )

    note = run_transfer_from_csvs(
        req=req,
        ctx=ctx,
        logger=logger,
        append_log=append_log,
    )

    logger.info("Job finished: %s", note)
    return note
