# excel_worker.py
from __future__ import annotations

import os
import queue
from typing import Dict, Any, Optional

import pythoncom
import win32com.client
from PySide6.QtCore import QThread, Signal

from logger import get_logger

logger = get_logger("ExcelWorker")


class ExcelWorker(QThread):
    """
    Excel COM ワーカー（遅延起動）

    TreeView からの意味操作を受け取り、COM 操作へ変換する。
    - UIスレッドからは request_* でキュー投入のみ
    - worker thread が順番に実行

    ★対応:
      open/close/list_sheets/activate/select/set/move/copy/cut/paste
      select_move(Shift+Arrow) / move_edge(Ctrl+Arrow) / select_edge(Ctrl+Shift+Arrow)
      undo/redo/select_all/fill_down/fill_right
      get_active_context（UIポーリング用）
      shutdown/stop/request_stop（終了安全化）
    """

    sheets_ready = Signal(str, list)
    active_cell_changed = Signal(str)

    # Excel XlDirection
    _XL_DIR = {
        "up": -4121,     # xlUp
        "down": -4120,   # xlDown
        "left": -4159,   # xlToLeft
        "right": -4161,  # xlToRight
    }

    def __init__(self):
        super().__init__()
        self._cmd_q: queue.Queue = queue.Queue()
        self._books: Dict[str, object] = {}
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

    # --- 互換: TreeView から呼ばれ得る停止API ---
    def request_stop(self):
        logger.info("[ExcelWorker] request_stop")
        self._running = False
        self._cmd_q.put(("quit",))

    def stop(self):
        logger.info("[ExcelWorker] stop()")
        try:
            self.request_stop()
        except Exception as e:
            logger.error("[ExcelWorker] stop request failed: %s", e, exc_info=True)
        try:
            self.wait(3000)
        except Exception as e:
            logger.error("[ExcelWorker] wait failed: %s", e, exc_info=True)

    def shutdown(self, confirm_save: bool = True):
        """
        安全終了（TreeView の closeEvent から呼ばれる想定）
        confirm_save=True の場合、開いているブックは保存確認を出さず閉じる（DisplayAlerts=False なので基本出ない）
        """
        logger.info("[ExcelWorker] shutdown(confirm_save=%s)", confirm_save)
        # confirm_save は現状「保存せず閉じる」運用に合わせる（事故らない）
        # 将来必要なら SaveChanges=confirm_save を切り替え可能
        self.request_stop()

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

                try:
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

                    elif op == "select_move":
                        self._select_move(args[0])

                    elif op == "move_edge":
                        self._move_edge(args[0])

                    elif op == "select_edge":
                        self._select_edge(args[0])

                    elif op == "copy":
                        self._copy()

                    elif op == "cut":
                        self._cut()

                    elif op == "paste":
                        self._paste()

                    elif op == "undo":
                        self._undo()

                    elif op == "redo":
                        self._redo()

                    elif op == "select_all":
                        self._select_all()

                    elif op == "fill_down":
                        self._fill_down()

                    elif op == "fill_right":
                        self._fill_right()

                    elif op == "quit":
                        logger.info("[ExcelWorker] quit received")
                        break

                    else:
                        logger.info("[ExcelWorker] unknown op=%s args=%s", op, args)

                except Exception as e:
                    logger.error("[ExcelWorker] op failed op=%s err=%s", op, e, exc_info=True)

        finally:
            try:
                self._shutdown()
            except Exception as e:
                logger.error("[ExcelWorker] shutdown failed: %s", e, exc_info=True)

            pythoncom.CoUninitialize()
            logger.info("[ExcelWorker] thread stopped")

    # ===============================
    # internal helpers
    # ===============================
    def _ensure_app(self):
        if self._app:
            return
        logger.info("[ExcelWorker] Dispatch Excel.Application")
        self._app = win32com.client.Dispatch("Excel.Application")
        self._app.Visible = True
        self._app.DisplayAlerts = False

    def _active_book(self):
        if not self._app:
            return None
        try:
            return self._app.ActiveWorkbook
        except Exception:
            return None

    def _emit_active_cell(self):
        try:
            if not self._app:
                return
            cell = self._app.ActiveCell.Address
            self.active_cell_changed.emit(cell)
        except Exception:
            pass

    def get_active_context(self) -> Dict[str, str]:
        """
        UI からポーリングされる想定（Inspector）
        """
        try:
            if not self._app:
                return {}
            addr = ""
            sheet = ""
            book = ""
            try:
                if self._app.ActiveCell is not None:
                    addr = str(self._app.ActiveCell.Address)
            except Exception:
                addr = ""
            try:
                if self._app.ActiveSheet is not None:
                    sheet = str(self._app.ActiveSheet.Name)
            except Exception:
                sheet = ""
            try:
                if self._app.ActiveWorkbook is not None:
                    book = str(self._app.ActiveWorkbook.Name)
            except Exception:
                book = ""
            return {"address": addr, "sheet": sheet, "workbook": book}
        except Exception:
            return {}

    # ===============================
    # Excel ops
    # ===============================
    def _open(self, path: str):
        try:
            if path in self._books:
                logger.info("[ExcelWorker] open ignored (already opened) path=%s", path)
                return
            if not os.path.exists(path):
                logger.info("[ExcelWorker] open ignored (not exists) path=%s", path)
                return
            self._ensure_app()
            logger.info("[ExcelWorker] open path=%s", path)
            wb = self._app.Workbooks.Open(path)
            self._books[path] = wb
            self._emit_active_cell()
        except Exception as e:
            logger.error("[ExcelWorker] open failed: %s", e, exc_info=True)

    def _close(self, path: str):
        try:
            wb = self._books.pop(path, None)
            if wb:
                logger.info("[ExcelWorker] close path=%s", path)
                wb.Close(SaveChanges=False)
        except Exception as e:
            logger.error("[ExcelWorker] close failed: %s", e, exc_info=True)

    def _list_sheets(self, path: str):
        try:
            wb = self._books.get(path)
            if wb:
                sheets = [ws.Name for ws in wb.Worksheets]
                logger.info("[ExcelWorker] list_sheets path=%s count=%s", path, len(sheets))
                self.sheets_ready.emit(path, sheets)
        except Exception as e:
            logger.error("[ExcelWorker] list_sheets failed: %s", e, exc_info=True)

    def _activate_book(self, path: str, front: bool):
        try:
            wb = self._books.get(path)
            if wb:
                logger.info("[ExcelWorker] activate_book path=%s", path)
                wb.Activate()
                self._emit_active_cell()
        except Exception as e:
            logger.error("[ExcelWorker] activate_book failed: %s", e, exc_info=True)

    def _activate_sheet(self, path: str, sheet: str, front: bool):
        try:
            wb = self._books.get(path)
            if wb:
                logger.info("[ExcelWorker] activate_sheet path=%s sheet=%s", path, sheet)
                wb.Worksheets(sheet).Activate()
                self._emit_active_cell()
        except Exception as e:
            logger.error("[ExcelWorker] activate_sheet failed: %s", e, exc_info=True)

    def _select_cell(self, cell: str):
        try:
            wb = self._active_book()
            if wb:
                logger.info("[ExcelWorker] select_cell %s", cell)
                wb.Application.Range(cell).Select()
                self._emit_active_cell()
        except Exception as e:
            logger.error("[ExcelWorker] select_cell failed: %s", e, exc_info=True)

    def _select_range(self, anchor: str, active: str):
        try:
            wb = self._active_book()
            if not wb:
                return
            logger.info("[ExcelWorker] select_range anchor=%s active=%s", anchor, active)
            app = wb.Application
            app.Range(anchor).Select()
            # Extend は ActiveCell を基準に伸びるので、実質 anchor->現在のActive を伸ばす
            # TreeView/Inspector 側が anchor を管理する想定
            app.Selection.Extend(app.ActiveCell)
            self._emit_active_cell()
        except Exception as e:
            logger.error("[ExcelWorker] select_range failed: %s", e, exc_info=True)

    def _set_cell_value(self, cell: str, value):
        try:
            wb = self._active_book()
            if wb:
                logger.info("[ExcelWorker] set_cell_value cell=%s value=%s", cell, value)
                app = wb.Application
                if cell == "*":
                    app.Selection.Value = value
                else:
                    app.Range(cell).Value = value
                self._emit_active_cell()
        except Exception as e:
            logger.error("[ExcelWorker] set_cell_value failed: %s", e, exc_info=True)

    def _move_cell(self, direction: str, step: int):
        try:
            wb = self._active_book()
            if not wb:
                return
            app = wb.Application
            dx, dy = {
                "up": (-step, 0),
                "down": (step, 0),
                "left": (0, -step),
                "right": (0, step),
            }[direction]
            logger.info("[ExcelWorker] move_cell dir=%s step=%s", direction, step)
            app.ActiveCell.Offset(dx, dy).Select()
            self._emit_active_cell()
        except Exception as e:
            logger.error("[ExcelWorker] move_cell failed: %s", e, exc_info=True)

    def _select_move(self, direction: str):
        """
        Shift + Arrow 相当（選択範囲を1セル拡張）
        """
        try:
            wb = self._active_book()
            if not wb:
                return
            app = wb.Application
            dx, dy = {
                "up": (-1, 0),
                "down": (1, 0),
                "left": (0, -1),
                "right": (0, 1),
            }[direction]
            logger.info("[ExcelWorker] select_move dir=%s", direction)
            app.Selection.Extend(app.ActiveCell.Offset(dx, dy))
            self._emit_active_cell()
        except Exception as e:
            logger.error("[ExcelWorker] select_move failed: %s", e, exc_info=True)

    def _move_edge(self, direction: str):
        """
        Ctrl + Arrow 相当（端まで移動）
        """
        try:
            wb = self._active_book()
            if not wb:
                return
            app = wb.Application
            logger.info("[ExcelWorker] move_edge dir=%s", direction)
            app.ActiveCell.End(self._XL_DIR[direction]).Select()
            self._emit_active_cell()
        except Exception as e:
            logger.error("[ExcelWorker] move_edge failed: %s", e, exc_info=True)

    def _select_edge(self, direction: str):
        """
        Ctrl + Shift + Arrow 相当（端まで選択）
        """
        try:
            wb = self._active_book()
            if not wb:
                return
            app = wb.Application
            logger.info("[ExcelWorker] select_edge dir=%s", direction)
            app.Selection.Extend(app.ActiveCell.End(self._XL_DIR[direction]))
            self._emit_active_cell()
        except Exception as e:
            logger.error("[ExcelWorker] select_edge failed: %s", e, exc_info=True)

    def _copy(self):
        try:
            wb = self._active_book()
            if wb:
                logger.info("[ExcelWorker] copy")
                wb.Application.Selection.Copy()
        except Exception as e:
            logger.error("[ExcelWorker] copy failed: %s", e, exc_info=True)

    def _cut(self):
        try:
            wb = self._active_book()
            if wb:
                logger.info("[ExcelWorker] cut")
                wb.Application.Selection.Cut()
        except Exception as e:
            logger.error("[ExcelWorker] cut failed: %s", e, exc_info=True)

    def _paste(self):
        try:
            wb = self._active_book()
            if wb:
                logger.info("[ExcelWorker] paste")
                wb.Application.ActiveSheet.Paste()
        except Exception as e:
            logger.error("[ExcelWorker] paste failed: %s", e, exc_info=True)

    def _undo(self):
        try:
            wb = self._active_book()
            if wb:
                logger.info("[ExcelWorker] undo")
                wb.Application.Undo()
                self._emit_active_cell()
        except Exception as e:
            logger.error("[ExcelWorker] undo failed: %s", e, exc_info=True)

    def _redo(self):
        """
        Excel COM は Redo API が安定しないことがあるので SendKeys を使用。
        """
        try:
            wb = self._active_book()
            if wb:
                logger.info("[ExcelWorker] redo")
                # Ctrl+Y
                wb.Application.SendKeys("^y")
                self._emit_active_cell()
        except Exception as e:
            logger.error("[ExcelWorker] redo failed: %s", e, exc_info=True)

    def _select_all(self):
        try:
            wb = self._active_book()
            if wb:
                logger.info("[ExcelWorker] select_all")
                wb.Application.Cells.Select()
                self._emit_active_cell()
        except Exception as e:
            logger.error("[ExcelWorker] select_all failed: %s", e, exc_info=True)

    def _fill_down(self):
        try:
            wb = self._active_book()
            if wb:
                logger.info("[ExcelWorker] fill_down")
                wb.Application.Selection.FillDown()
                self._emit_active_cell()
        except Exception as e:
            logger.error("[ExcelWorker] fill_down failed: %s", e, exc_info=True)

    def _fill_right(self):
        try:
            wb = self._active_book()
            if wb:
                logger.info("[ExcelWorker] fill_right")
                wb.Application.Selection.FillRight()
                self._emit_active_cell()
        except Exception as e:
            logger.error("[ExcelWorker] fill_right failed: %s", e, exc_info=True)

    # ===============================
    # shutdown
    # ===============================
    def _shutdown(self):
        logger.info("[ExcelWorker] _shutdown begin")
        # ブックを閉じる
        for p, wb in list(self._books.items()):
            try:
                logger.info("[ExcelWorker] closing book path=%s", p)
                wb.Close(SaveChanges=False)
            except Exception as e:
                logger.error("[ExcelWorker] close book failed path=%s err=%s", p, e, exc_info=True)

        self._books.clear()

        # Excel app 終了
        try:
            if self._app:
                logger.info("[ExcelWorker] app quit")
                self._app.Quit()
        except Exception as e:
            logger.error("[ExcelWorker] app quit failed: %s", e, exc_info=True)

        self._app = None
        logger.info("[ExcelWorker] _shutdown done")
