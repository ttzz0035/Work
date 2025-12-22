# excel_worker.py
from __future__ import annotations

import os
import queue
from typing import Dict, Optional

import pythoncom
import win32com.client

from PySide6.QtCore import QThread, Signal

from Logger import Logger


logger = Logger(
    name="ExcelWorker",
    log_file_path="logs/app.log",
    level="DEBUG",
)


class ExcelWorker(QThread):
    """
    Excel COM ワーカー（遅延起動）
    """

    sheets_ready = Signal(str, list)
    active_cell_changed = Signal(str)

    book_closed = Signal(str)
    book_close_failed = Signal(str, str)
    quit_finished = Signal()

    _XL_DIR = {
        "up": -4121,
        "down": -4120,
        "left": -4159,
        "right": -4161,
    }

    _SK_ARROW = {
        "up": "{UP}",
        "down": "{DOWN}",
        "left": "{LEFT}",
        "right": "{RIGHT}",
    }

    def __init__(self):
        super().__init__()
        self._cmd_q: queue.Queue = queue.Queue()
        self._books: Dict[str, object] = {}
        self._app = None

        self._running: bool = True
        self._active_path: str = ""
        self._ctx: Dict[str, str] = {"address": "", "sheet": "", "workbook": ""}

        logger.info("[ExcelWorker] initialized")

    # ===============================
    # public API（UI スレッド）
    # ===============================
    def request_open(self, path: str):
        self._cmd_q.put(("open", os.path.abspath(path)))

    def request_close(self, path: str):
        self._cmd_q.put(("close", os.path.abspath(path)))

    def request_list_sheets(self, path: str):
        self._cmd_q.put(("list_sheets", os.path.abspath(path)))

    def request_activate_book(self, path: str, front: bool = False):
        self._cmd_q.put(("activate_book", os.path.abspath(path), front))

    def request_activate_sheet(self, path: str, sheet: str, front: bool = False):
        self._cmd_q.put(("activate_sheet", os.path.abspath(path), sheet, front))

    def request_select_cell(self, cell: str):
        self._cmd_q.put(("select_cell", cell))

    def request_set_cell_value(self, cell: str, value):
        self._cmd_q.put(("set_cell_value", cell, value))

    def request_move_cell(self, direction: str, step: int = 1):
        self._cmd_q.put(("move_cell", direction, step))

    def request_select_move(self, direction: str):
        self._cmd_q.put(("select_move", direction))

    def request_move_edge(self, direction: str):
        self._cmd_q.put(("move_edge", direction))

    def request_select_edge(self, direction: str):
        self._cmd_q.put(("select_edge", direction))

    def request_copy(self):
        self._cmd_q.put(("copy",))

    def request_cut(self):
        self._cmd_q.put(("cut",))

    def request_paste(self):
        self._cmd_q.put(("paste",))

    def request_undo(self):
        self._cmd_q.put(("undo",))

    def request_redo(self):
        self._cmd_q.put(("redo",))

    def request_select_all(self):
        self._cmd_q.put(("select_all",))

    def request_fill_down(self):
        self._cmd_q.put(("fill_down",))

    def request_fill_right(self):
        self._cmd_q.put(("fill_right",))

    def request_quit(self):
        self._cmd_q.put(("quit",))

    def request_stop(self):
        logger.info("[ExcelWorker] request_stop")
        self._running = False
        self._cmd_q.put(("quit",))

    def shutdown(self, confirm_save: bool = True):
        logger.info(f"[ExcelWorker] shutdown(confirm_save={confirm_save})")
        self.request_stop()

    # ===============================
    # worker thread
    # ===============================
    def run(self):
        pythoncom.CoInitialize()
        logger.info("[ExcelWorker] thread started (COM initialized)")

        try:
            while self._running:
                op, *args = self._cmd_q.get()

                try:
                    getattr(self, f"_{op}")(*args)
                except AttributeError:
                    logger.warning(f"[ExcelWorker] unknown op={op} args={args}")
                except Exception as e:
                    logger.error(f"[ExcelWorker] op failed op={op} err={e}")

        finally:
            self._shutdown()
            pythoncom.CoUninitialize()
            logger.info("[ExcelWorker] thread stopped (COM uninitialized)")
            self.quit_finished.emit()

    # ===============================
    # internal helpers
    # ===============================
    def _ensure_app(self):
        if self._app:
            return
        logger.info("[ExcelWorker] DispatchEx Excel.Application")
        self._app = win32com.client.DispatchEx("Excel.Application")
        self._app.Visible = True
        self._app.DisplayAlerts = False

    def _active_book(self):
        if not self._app:
            return None
        try:
            return self._app.ActiveWorkbook
        except Exception:
            return None

    def _snapshot_ctx_from_com(self):
        try:
            return {
                "address": str(self._app.ActiveCell.Address),
                "sheet": str(self._app.ActiveSheet.Name),
                "workbook": str(self._app.ActiveWorkbook.Name),
            }
        except Exception:
            return {"address": "", "sheet": "", "workbook": ""}

    def _update_context_cache(self):
        self._ctx = self._snapshot_ctx_from_com()

    # ===============================
    # Excel ops
    # ===============================
    def _move_cell(self, direction: str, step: int):
        wb = self._active_book()
        if not wb:
            logger.warning("[CUT] EXCEL_MOVE skipped (no book)")
            return

        before = self._snapshot_ctx_from_com()
        logger.warning(
            f"[CUT] EXCEL_MOVE ENTER dir={direction} step={step} before={before}"
        )

        key = self._SK_ARROW[direction]
        for i in range(step):
            logger.warning(f"[CUT] EXCEL_MOVE SENDKEY i={i} key={key}")
            self._app.SendKeys(key)

        self._update_context_cache()
        after = dict(self._ctx)

        logger.warning(f"[CUT] EXCEL_MOVE EXIT after={after}")

    # ===============================
    # shutdown
    # ===============================
    def _shutdown(self):
        logger.info("[ExcelWorker] _shutdown begin")

        for p, wb in list(self._books.items()):
            try:
                logger.info(f"[ExcelWorker] closing book path={p}")
                wb.Close(SaveChanges=False)
                self.book_closed.emit(p)
            except Exception as e:
                logger.error(f"[ExcelWorker] close book failed path={p} err={e}")
                self.book_close_failed.emit(p, str(e))

        self._books.clear()
        self._active_path = ""

        if self._app:
            try:
                logger.info("[ExcelWorker] app quit")
                self._app.Quit()
            except Exception as e:
                logger.error(f"[ExcelWorker] app quit failed: {e}")

        self._app = None
        self._ctx = {"address": "", "sheet": "", "workbook": ""}

        logger.info("[ExcelWorker] _shutdown done")
