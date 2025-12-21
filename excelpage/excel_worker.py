# excel_worker.py
import os
import queue
import pythoncom
import win32com.client
from PySide6.QtCore import QThread, Signal
from logger import get_logger

logger = get_logger("ExcelWorker")


class ExcelWorker(QThread):
    """
    Excel COM ワーカー（遅延起動）
    """

    sheets_ready = Signal(str, list)
    active_cell_changed = Signal(str)

    def __init__(self):
        super().__init__()
        self._cmd_q: queue.Queue = queue.Queue()
        self._books: dict[str, object] = {}
        self._app = None
        self._running = True

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

    def request_select_range(self, anchor: str, active: str):
        self._cmd_q.put(("select_range", anchor, active))

    def request_set_cell_value(self, cell: str, value):
        self._cmd_q.put(("set_cell_value", cell, value))

    def request_move_cell(self, direction: str, step: int = 1):
        self._cmd_q.put(("move_cell", direction, step))

    def request_copy(self):
        self._cmd_q.put(("copy",))

    def request_cut(self):
        self._cmd_q.put(("cut",))

    def request_paste(self):
        self._cmd_q.put(("paste",))

    def request_quit(self):
        self._cmd_q.put(("quit",))

    # ===============================
    # worker thread
    # ===============================
    def run(self):
        pythoncom.CoInitialize()
        logger.info("[ExcelWorker] thread started")

        try:
            while self._running:
                cmd = self._cmd_q.get()
                op = cmd[0]
                args = cmd[1:]

                if op == "open":
                    self._open(args[0])

                elif op == "close":
                    self._close(args[0])

                elif op == "list_sheets":
                    self._list_sheets(args[0])

                elif op == "activate_book":
                    self._activate_book(args[0], args[1])

                elif op == "activate_sheet":
                    self._activate_sheet(args[0], args[1], args[2])

                elif op == "select_cell":
                    self._select_cell(args[0])

                elif op == "select_range":
                    self._select_range(args[0], args[1])

                elif op == "set_cell_value":
                    self._set_cell_value(args[0], args[1])

                elif op == "move_cell":
                    self._move_cell(args[0], args[1])

                elif op == "copy":
                    self._copy()

                elif op == "cut":
                    self._cut()

                elif op == "paste":
                    self._paste()

                elif op == "quit":
                    break

        finally:
            self._shutdown()
            pythoncom.CoUninitialize()
            logger.info("[ExcelWorker] thread stopped")

    # ===============================
    # internal helpers
    # ===============================
    def _ensure_app(self):
        if self._app:
            return
        self._app = win32com.client.Dispatch("Excel.Application")
        self._app.Visible = True
        self._app.DisplayAlerts = False

    def _active_book(self):
        return self._app.ActiveWorkbook if self._app else None

    def _emit_active_cell(self):
        try:
            cell = self._app.ActiveCell.Address
            self.active_cell_changed.emit(cell)
        except Exception:
            pass

    def _open(self, path: str):
        if path in self._books or not os.path.exists(path):
            return
        self._ensure_app()
        wb = self._app.Workbooks.Open(path)
        self._books[path] = wb
        self._emit_active_cell()

    def _close(self, path: str):
        wb = self._books.pop(path, None)
        if wb:
            wb.Close(SaveChanges=False)

    def _list_sheets(self, path: str):
        wb = self._books.get(path)
        if wb:
            self.sheets_ready.emit(path, [ws.Name for ws in wb.Worksheets])

    def _activate_book(self, path: str, front: bool):
        wb = self._books.get(path)
        if wb:
            wb.Activate()
            self._emit_active_cell()

    def _activate_sheet(self, path: str, sheet: str, front: bool):
        wb = self._books.get(path)
        if wb:
            wb.Worksheets(sheet).Activate()
            self._emit_active_cell()

    def _select_cell(self, cell: str):
        wb = self._active_book()
        if wb:
            wb.Application.Range(cell).Select()
            self._emit_active_cell()

    def _select_range(self, anchor: str, active: str):
        wb = self._active_book()
        if not wb:
            return
        wb.Application.Range(anchor).Select()
        wb.Application.Selection.Extend(wb.Application.ActiveCell)
        self._emit_active_cell()

    def _set_cell_value(self, cell: str, value):
        wb = self._active_book()
        if wb:
            if cell == "*":
                wb.Application.Selection.Value = value
            else:
                wb.Application.Range(cell).Value = value
            self._emit_active_cell()

    def _move_cell(self, direction: str, step: int):
        wb = self._active_book()
        if not wb:
            return
        dx, dy = {
            "up": (-step, 0),
            "down": (step, 0),
            "left": (0, -step),
            "right": (0, step),
        }[direction]
        wb.Application.ActiveCell.Offset(dx, dy).Select()
        self._emit_active_cell()

    def _copy(self):
        wb = self._active_book()
        if wb:
            wb.Application.Selection.Copy()

    def _cut(self):
        wb = self._active_book()
        if wb:
            wb.Application.Selection.Cut()

    def _paste(self):
        wb = self._active_book()
        if wb:
            wb.Application.ActiveSheet.Paste()

    def _shutdown(self):
        for wb in self._books.values():
            try:
                wb.Close(SaveChanges=False)
            except Exception:
                pass
        self._books.clear()
        if self._app:
            self._app.Quit()
        self._app = None
