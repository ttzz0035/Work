from __future__ import annotations

import json
import threading
import time
from typing import Dict, Any

from PySide6.QtCore import QThread
from logger import get_logger


logger = get_logger("MacroPlayer")


class MacroPlayThread(QThread):
    """
    マクロ再生スレッド（停止対応・ExcelWorker 完全対応）
    """

    def __init__(self, excel_worker, macro_path: str, sleep_sec: float = 0.05):
        super().__init__()
        self._excel = excel_worker
        self._macro_path = macro_path
        self._sleep_sec = sleep_sec
        self._stop_event = threading.Event()

    # ===============================
    # public
    # ===============================
    def stop(self):
        logger.info("[Macro] stop requested")
        self._stop_event.set()

    # ===============================
    # thread entry
    # ===============================
    def run(self):
        logger.info("[Macro] run start path=%s", self._macro_path)

        try:
            with open(self._macro_path, "r", encoding="utf-8") as f:
                macro = json.load(f)

            steps = macro.get("steps", [])
            logger.info("[Macro] steps=%s", len(steps))

            for idx, step in enumerate(steps):
                if self._stop_event.is_set():
                    logger.warning("[Macro] stopped at step=%s", idx)
                    break

                op = step.get("op")
                args = step.get("args", {})

                logger.info("[Macro] STEP %s op=%s args=%s", idx, op, args)
                self._dispatch(op, args)

                time.sleep(self._sleep_sec)

        except Exception as e:
            logger.error("[Macro] run failed: %s", e, exc_info=True)

        logger.info("[Macro] run end")

    # ===============================
    # dispatcher
    # ===============================
    def _dispatch(self, op: str, args: Dict[str, Any]):
        """
        ExcelWorker の request_* API に完全マッピング
        """

        ew = self._excel

        # ---- book / sheet ----
        if op == "open":
            ew.request_open(args["path"])

        elif op == "close":
            ew.request_close(args["path"])

        elif op == "list_sheets":
            ew.request_list_sheets(args["path"])

        elif op == "activate_book":
            ew.request_activate_book(args["path"], args.get("front", False))

        elif op == "activate_sheet":
            ew.request_activate_sheet(
                args["path"], args["sheet"], args.get("front", False)
            )

        # ---- selection ----
        elif op == "select_cell":
            ew.request_select_cell(args["cell"])

        elif op == "select_range":
            ew.request_select_range(args["anchor"], args["active"])

        elif op == "select_all":
            ew.request_select_all()

        # ---- value ----
        elif op == "set_cell_value":
            ew.request_set_cell_value(args["cell"], args["value"])

        # ---- move ----
        elif op == "move_cell":
            ew.request_move_cell(args["direction"], int(args.get("step", 1)))

        elif op == "select_move":
            ew.request_select_move(args["direction"])

        elif op == "move_edge":
            ew.request_move_edge(args["direction"])

        elif op == "select_edge":
            ew.request_select_edge(args["direction"])

        # ---- clipboard ----
        elif op == "copy":
            ew.request_copy()

        elif op == "cut":
            ew.request_cut()

        elif op == "paste":
            ew.request_paste()

        # ---- undo / redo ----
        elif op == "undo":
            ew.request_undo()

        elif op == "redo":
            ew.request_redo()

        # ---- fill ----
        elif op == "fill_down":
            ew.request_fill_down()

        elif op == "fill_right":
            ew.request_fill_right()

        # ---- lifecycle ----
        elif op == "quit":
            ew.request_quit()

        else:
            logger.error("[Macro] unsupported op=%s args=%s", op, args)
