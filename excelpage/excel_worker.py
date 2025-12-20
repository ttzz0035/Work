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
    - 起動時に Excel を立ち上げない
    - 必要になった瞬間だけ起動
    - UI / TreeView への依存なし
    """

    sheets_ready = Signal(str, list)

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

    def request_quit(self):
        self._cmd_q.put(("quit",))

    def has_workbook(self, path: str) -> bool:
        return os.path.abspath(path) in self._books

    def get_workbook(self, path: str):
        return self._books.get(os.path.abspath(path))

    def request_select_cell(self, cell: str):
        self._enqueue(("select_cell", cell))

    def request_set_cell_value(self, cell: str, value):
        self._enqueue(("set_cell_value", cell, value))

    def request_move_cell(self, direction: str, step: int = 1):
        self._enqueue(("move_cell", direction, step))

    def request_copy(self):
        self._enqueue(("copy",))

    def request_paste(self):
        self._enqueue(("paste",))

    # ===============================
    # worker thread
    # ===============================
    def run(self):
        pythoncom.CoInitialize()
        logger.info("ExcelWorker thread started")

        try:
            while self._running:
                cmd = self._cmd_q.get()
                op = cmd[0]

                if op == "open":
                    self._open(cmd[1])

                elif op == "close":
                    self._close(cmd[1])

                elif op == "list_sheets":
                    self._list_sheets(cmd[1])

                elif op == "activate_book":
                    self._activate_book(cmd[1], cmd[2])

                elif op == "activate_sheet":
                    self._activate_sheet(cmd[1], cmd[2], cmd[3])

                elif op == "quit":
                    break

                elif cmd == "select_cell":
                    cell = args[0]
                    self._book.app.api.Range(cell).Select()

                elif cmd == "set_cell_value":
                    cell, val = args
                    self._book.app.api.Range(cell).Value = val

                elif cmd == "move_cell":
                    direction, step = args
                    dx, dy = {
                        "up": (-1, 0),
                        "down": (1, 0),
                        "left": (0, -1),
                        "right": (0, 1),
                    }[direction]
                    self._book.app.api.ActiveCell.Offset(dx * step, dy * step).Select()

                elif cmd == "copy":
                    self._book.app.api.Selection.Copy()

                elif cmd == "paste":
                    self._book.app.api.ActiveSheet.Paste()

        finally:
            self._shutdown()
            pythoncom.CoUninitialize()
            logger.info("ExcelWorker thread stopped")

    # ===============================
    # internal
    # ===============================
    def _ensure_app(self):
        if self._app is not None:
            return

        self._app = win32com.client.Dispatch("Excel.Application")
        self._app.Visible = True
        self._app.DisplayAlerts = False
        logger.info("Excel COM started (lazy)")

    def _open(self, path: str):
        if path in self._books:
            return

        if not os.path.exists(path):
            logger.error("Excel not found: %s", path)
            return

        self._ensure_app()

        try:
            wb = self._app.Workbooks.Open(path)
            self._books[path] = wb
            logger.info("Excel opened: %s", path)
        except Exception as e:
            logger.error("Excel open failed: %s (%s)", path, e)

    def _close(self, path: str):
        wb = self._books.pop(path, None)
        if not wb:
            return
        try:
            wb.Close(SaveChanges=False)
            logger.info("Excel closed: %s", path)
        except Exception as e:
            logger.error("Excel close failed: %s (%s)", path, e)

    def _list_sheets(self, path: str):
        wb = self._books.get(path)
        if not wb:
            return
        try:
            sheets = [ws.Name for ws in wb.Worksheets]
            self.sheets_ready.emit(path, sheets)
        except Exception as e:
            logger.error("List sheets failed: %s (%s)", path, e)

    def _activate_book(self, path: str, front: bool):
        wb = self._books.get(path)
        if not wb:
            return
        try:
            wb.Activate()
            if front:
                self._app.Visible = True
                self._app.WindowState = -4143  # xlNormal
        except Exception as e:
            logger.error("Activate book failed: %s (%s)", path, e)

    def _activate_sheet(self, path: str, sheet: str, front: bool):
        wb = self._books.get(path)
        if not wb:
            return
        try:
            wb.Worksheets(sheet).Activate()
            if front:
                self._app.Visible = True
                self._app.WindowState = -4143
        except Exception as e:
            logger.error(
                "Activate sheet failed: %s / %s (%s)", path, sheet, e
            )

    def _shutdown(self):
        for path, wb in list(self._books.items()):
            try:
                wb.Close(SaveChanges=False)
            except Exception:
                pass
        self._books.clear()

        if self._app:
            try:
                self._app.Quit()
            except Exception:
                pass
        self._app = None
