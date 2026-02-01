# SPDX-License-Identifier: MIT
import logging
import threading
from datetime import datetime
from typing import Callable

from task_base import TaskBase
from task_impl import TaskImpl


# =========================================================
# Logger（root → UI queue）
# =========================================================
def _get_worker_logger():
    lg = logging.getLogger("WORKER")
    lg.setLevel(logging.INFO)
    lg.handlers.clear()
    lg.propagate = True
    return lg


# =========================================================
# Worker 本体
# =========================================================
def _run_worker_impl(
    runtime: dict,
    ui_call: Callable[[Callable], None],
    append_logs: Callable[[], None],
    update_status: Callable[[], None],
    stop_run: Callable[[], None],
):
    logger = _get_worker_logger()

    runtime["running"] = True
    runtime["started_at"] = datetime.now()
    runtime["ticks"] = 0

    logger.info(
        f"開始: job={runtime.get('item_id')} "
        f"period={runtime.get('start')}～{runtime.get('end')}"
    )
    ui_call(append_logs)
    ui_call(update_status)

    # ★ 抽象型として扱う（状態を持つ Task）
    task: TaskBase = TaskImpl(
        runtime=runtime,
        logger=logger,
        ui_call=ui_call,
        append_logs=append_logs,
        update_status=update_status,
    )

    try:
        task.run()
    except Exception as e:
        logger.exception(f"実行エラー: {e}")
        ui_call(append_logs)
    finally:
        runtime["running"] = False
        logger.info(f"終了: ticks={runtime['ticks']}")
        ui_call(append_logs)
        ui_call(update_status)
        ui_call(stop_run)



# =========================================================
# UI エントリ
# =========================================================
def run_worker(
    runtime: dict,
    append_logs: Callable[[], None],
    update_status: Callable[[], None],
    stop_run: Callable[[], None],
):
    logger = _get_worker_logger()

    if runtime.get("running"):
        logger.warning("既に worker が実行中です")
        return

    ui_call = lambda fn: fn()

    th = threading.Thread(
        target=_run_worker_impl,
        name="WorkerThread",
        args=(runtime, ui_call, append_logs, update_status, stop_run),
        daemon=True,
    )
    th.start()

    logger.info("worker thread started")
