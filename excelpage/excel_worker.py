# excel_worker.py
from __future__ import annotations

import os
import queue
from typing import Dict, Optional

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
    - COM は worker thread だけが触る（重要）
    - get_active_context は COM に触らずキャッシュを返す（重要）

    ★対応:
      open/close/list_sheets/activate/select/set/move/copy/cut/paste
      select_move(Shift+Arrow) / move_edge(Ctrl+Arrow) / select_edge(Ctrl+Shift+Arrow)
      undo/redo/select_all/fill_down/fill_right
      get_active_context（UIポーリング用：キャッシュ返却）
      shutdown/stop/request_stop（終了安全化）
    """

    sheets_ready = Signal(str, list)
    active_cell_changed = Signal(str)

    # TreeView 側で「閉じ終わり」を待つための通知
    book_closed = Signal(str)
    book_close_failed = Signal(str, str)
    quit_finished = Signal()

    # Excel XlDirection
    _XL_DIR = {
        "up": -4121,     # xlUp
        "down": -4120,   # xlDown
        "left": -4159,   # xlToLeft
        "right": -4161,  # xlToRight
    }

    # SendKeys 用
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

        # UIポーリング用のコンテキストキャッシュ（COMは触らない）
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
        confirm_save=True の場合でも現状は SaveChanges=False 運用（事故防止）
        """
        logger.info("[ExcelWorker] shutdown(confirm_save=%s)", confirm_save)
        self.request_stop()

    # ===============================
    # UI polling API（COMに触らない）
    # ===============================
    def get_active_context(self) -> Dict[str, str]:
        """
        Inspector が UIスレッドから呼ぶ想定。
        COMに触ると死ぬので、キャッシュを返すだけ。
        """
        try:
            return dict(self._ctx)
        except Exception:
            return {}

    # ===============================
    # worker thread
    # ===============================
    def run(self):
        pythoncom.CoInitialize()
        logger.info("[ExcelWorker] thread started (COM initialized)")

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
            logger.info("[ExcelWorker] thread stopped (COM uninitialized)")
            try:
                self.quit_finished.emit()
            except Exception:
                pass

    # ===============================
    # internal helpers
    # ===============================
    def _ensure_app(self):
        if self._app:
            return
        logger.info("[ExcelWorker] DispatchEx Excel.Application")
        # DispatchEx の方が他Excelインスタンスと分離できて安定しやすい
        self._app = win32com.client.DispatchEx("Excel.Application")
        self._app.Visible = True
        self._app.DisplayAlerts = False

    def _try_fix_active_book(self) -> Optional[object]:
        """
        ActiveWorkbook が取れない/None の時に復旧を試みる。
        """
        if not self._app:
            return None

        try:
            wb = self._app.ActiveWorkbook
            if wb is not None:
                return wb
        except Exception:
            wb = None

        if self._active_path and self._active_path in self._books:
            try:
                self._books[self._active_path].Activate()
                return self._books[self._active_path]
            except Exception:
                return None

        try:
            if self._books:
                any_path = next(iter(self._books.keys()))
                self._books[any_path].Activate()
                self._active_path = any_path
                return self._books[any_path]
        except Exception:
            return None

        return None

    def _active_book(self) -> Optional[object]:
        if not self._app:
            return None
        return self._try_fix_active_book()

    def _snapshot_ctx_from_com(self) -> Dict[str, str]:
        """
        worker thread 内でのみ呼ぶ（COMに触る）
        """
        if not self._app:
            return {"address": "", "sheet": "", "workbook": ""}

        addr = ""
        sheet = ""
        book = ""

        try:
            ac = self._app.ActiveCell
            if ac is not None:
                addr = str(ac.Address)
        except Exception:
            addr = ""

        try:
            sh = self._app.ActiveSheet
            if sh is not None:
                sheet = str(sh.Name)
        except Exception:
            sheet = ""

        try:
            wb = self._app.ActiveWorkbook
            if wb is not None:
                book = str(wb.Name)
        except Exception:
            book = ""

        return {"address": addr, "sheet": sheet, "workbook": book}

    def _update_context_cache(self):
        """
        worker thread 内でのみ COM に触って ctx を更新する
        """
        try:
            self._ctx = self._snapshot_ctx_from_com()
            addr = self._ctx.get("address", "")
            if addr:
                try:
                    self.active_cell_changed.emit(addr)
                except Exception:
                    pass
        except Exception as e:
            logger.error("[ExcelWorker] update_context_cache failed: %s", e, exc_info=True)

    def _log_before_after(self, op: str, before: Dict[str, str], after: Dict[str, str], extra: str = ""):
        logger.info(
            "[ExcelWorker] %s %s before=%s!%s(%s) after=%s!%s(%s)%s",
            op,
            extra,
            before.get("sheet", ""),
            before.get("address", ""),
            before.get("workbook", ""),
            after.get("sheet", ""),
            after.get("address", ""),
            after.get("workbook", ""),
            "",
        )

    # ===============================
    # Excel ops
    # ===============================
    def _open(self, path: str):
        try:
            if path in self._books:
                logger.info("[ExcelWorker] open ignored (already opened) path=%s", path)
                self._active_path = path
                try:
                    self._books[path].Activate()
                except Exception:
                    pass
                self._update_context_cache()
                return

            if not os.path.exists(path):
                logger.info("[ExcelWorker] open ignored (not exists) path=%s", path)
                return

            self._ensure_app()
            logger.info("[ExcelWorker] open path=%s", path)

            wb = self._app.Workbooks.Open(path)
            self._books[path] = wb
            self._active_path = path

            try:
                wb.Activate()
            except Exception:
                pass

            self._update_context_cache()

        except Exception as e:
            logger.error("[ExcelWorker] open failed: %s", e, exc_info=True)

    def _close(self, path: str):
        try:
            wb = self._books.pop(path, None)
            if wb:
                logger.info("[ExcelWorker] close path=%s", path)
                wb.Close(SaveChanges=False)

            if self._active_path == path:
                self._active_path = ""

            self._update_context_cache()

            try:
                self.book_closed.emit(path)
            except Exception:
                pass

        except Exception as e:
            logger.error("[ExcelWorker] close failed: %s", e, exc_info=True)
            try:
                self.book_close_failed.emit(path, str(e))
            except Exception:
                pass

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
                self._active_path = path
                self._update_context_cache()
        except Exception as e:
            logger.error("[ExcelWorker] activate_book failed: %s", e, exc_info=True)

    def _activate_sheet(self, path: str, sheet: str, front: bool):
        try:
            wb = self._books.get(path)
            if wb:
                logger.info("[ExcelWorker] activate_sheet path=%s sheet=%s", path, sheet)
                wb.Activate()
                self._active_path = path
                wb.Worksheets(sheet).Activate()
                self._update_context_cache()
        except Exception as e:
            logger.error("[ExcelWorker] activate_sheet failed: %s", e, exc_info=True)

    def _select_cell(self, cell: str):
        try:
            wb = self._active_book()
            if not wb:
                logger.info("[ExcelWorker] select_cell ignored (no active book) cell=%s", cell)
                return

            before = self._snapshot_ctx_from_com()

            app = self._app
            logger.info("[ExcelWorker] select_cell %s", cell)
            app.Range(cell).Select()

            self._update_context_cache()
            after = dict(self._ctx)
            self._log_before_after("select_cell", before, after, extra=f"cell={cell}")

        except Exception as e:
            logger.error("[ExcelWorker] select_cell failed: %s", e, exc_info=True)

    def _set_cell_value(self, cell: str, value):
        try:
            wb = self._active_book()
            if not wb:
                logger.info("[ExcelWorker] set_cell_value ignored (no active book)")
                return

            before = self._snapshot_ctx_from_com()

            app = self._app
            logger.info("[ExcelWorker] set_cell_value cell=%s value=%s", cell, value)

            if cell == "*":
                app.Selection.Value = value
            else:
                app.Range(cell).Value = value

            self._update_context_cache()
            after = dict(self._ctx)
            self._log_before_after("set_cell_value", before, after, extra=f"cell={cell}")

        except Exception as e:
            logger.error("[ExcelWorker] set_cell_value failed: %s", e, exc_info=True)

    def _move_cell(self, direction: str, step: int):
        try:
            wb = self._active_book()
            if not wb:
                logger.info("[ExcelWorker] move_cell ignored (no active book)")
                return

            app = self._app
            ac = app.ActiveCell  # ★ 常に ActiveCell 基準

            before = str(ac.Address)

            row_off, col_off = {
                "up": (-step, 0),
                "down": (step, 0),
                "left": (0, -step),
                "right": (0, step),
            }[direction]

            target = ac.Offset(row_off, col_off)  # ★ 位置引数のみ
            after = str(target.Address)

            logger.info(
                "[ExcelWorker] move_cell %s -> %s dir=%s",
                before, after, direction
            )

            target.Select()
            self._update_context_cache()

        except Exception as e:
            logger.error(
                "[ExcelWorker] move_cell failed dir=%s step=%s err=%s",
                direction, step, e, exc_info=True
            )

    # -------------------------------------------------
    # ★修正ポイント 2: Shift+Arrow は SendKeys（Excelの選択状態と完全一致）
    #   Extend は環境差で死ぬ/Selection起点にするとズレるため
    # -------------------------------------------------
    def _select_move(self, direction: str):
        """
        Shift + Arrow 相当（選択範囲を1セル拡張）
        """
        try:
            wb = self._active_book()
            if not wb:
                logger.info("[ExcelWorker] select_move ignored (no active book)")
                return

            before = self._snapshot_ctx_from_com()

            app = self._app
            key = self._SK_ARROW[direction]

            # Shift+Arrow
            logger.info("[ExcelWorker] select_move dir=%s sendkeys=+%s", direction, key)
            app.SendKeys("+" + key)

            self._update_context_cache()
            after = dict(self._ctx)
            self._log_before_after("select_move", before, after, extra=f"dir={direction}")

        except Exception as e:
            logger.error(
                "[ExcelWorker] select_move failed dir=%s err=%s",
                direction, e, exc_info=True
            )

    # -------------------------------------------------
    # Ctrl+Arrow は COM End でOK（安定）
    # -------------------------------------------------
    def _move_edge(self, direction: str):
        """
        Ctrl + Arrow 相当（端まで移動）
        """
        try:
            wb = self._active_book()
            if not wb:
                logger.info("[ExcelWorker] move_edge ignored (no active book)")
                return

            before = self._snapshot_ctx_from_com()

            app = self._app
            logger.info("[ExcelWorker] move_edge dir=%s", direction)
            app.ActiveCell.End(self._XL_DIR[direction]).Select()

            self._update_context_cache()
            after = dict(self._ctx)
            self._log_before_after("move_edge", before, after, extra=f"dir={direction}")

        except Exception as e:
            logger.error("[ExcelWorker] move_edge failed: %s", e, exc_info=True)

    # -------------------------------------------------
    # ★修正ポイント 3: Ctrl+Shift+Arrow も SendKeys（選択の起点ズレを起こさない）
    # -------------------------------------------------
    def _select_edge(self, direction: str):
        """
        Ctrl + Shift + Arrow 相当（端まで選択）
        """
        try:
            wb = self._active_book()
            if not wb:
                logger.info("[ExcelWorker] select_edge ignored (no active book)")
                return

            before = self._snapshot_ctx_from_com()

            app = self._app
            key = self._SK_ARROW[direction]

            # Ctrl+Shift+Arrow
            logger.info("[ExcelWorker] select_edge dir=%s sendkeys=^+%s", direction, key)
            app.SendKeys("^+" + key)

            self._update_context_cache()
            after = dict(self._ctx)
            self._log_before_after("select_edge", before, after, extra=f"dir={direction}")

        except Exception as e:
            logger.error("[ExcelWorker] select_edge failed: %s", e, exc_info=True)

    def _copy(self):
        try:
            wb = self._active_book()
            if not wb:
                logger.info("[ExcelWorker] copy ignored (no active book)")
                return
            logger.info("[ExcelWorker] copy")
            self._app.Selection.Copy()
        except Exception as e:
            logger.error("[ExcelWorker] copy failed: %s", e, exc_info=True)

    def _cut(self):
        try:
            wb = self._active_book()
            if not wb:
                logger.info("[ExcelWorker] cut ignored (no active book)")
                return
            logger.info("[ExcelWorker] cut")
            self._app.Selection.Cut()
        except Exception as e:
            logger.error("[ExcelWorker] cut failed: %s", e, exc_info=True)

    def _paste(self):
        try:
            wb = self._active_book()
            if not wb:
                logger.info("[ExcelWorker] paste ignored (no active book)")
                return

            before = self._snapshot_ctx_from_com()

            logger.info("[ExcelWorker] paste")
            self._app.ActiveSheet.Paste()

            self._update_context_cache()
            after = dict(self._ctx)
            self._log_before_after("paste", before, after)

        except Exception as e:
            logger.error("[ExcelWorker] paste failed: %s", e, exc_info=True)

    def _undo(self):
        try:
            wb = self._active_book()
            if not wb:
                logger.info("[ExcelWorker] undo ignored (no active book)")
                return

            before = self._snapshot_ctx_from_com()

            logger.info("[ExcelWorker] undo")
            self._app.Undo()

            self._update_context_cache()
            after = dict(self._ctx)
            self._log_before_after("undo", before, after)

        except Exception as e:
            logger.error("[ExcelWorker] undo failed: %s", e, exc_info=True)

    def _redo(self):
        """
        Excel COM は Redo API が安定しないことがあるので SendKeys を使用。
        """
        try:
            wb = self._active_book()
            if not wb:
                logger.info("[ExcelWorker] redo ignored (no active book)")
                return

            before = self._snapshot_ctx_from_com()

            logger.info("[ExcelWorker] redo (SendKeys ^y)")
            self._app.SendKeys("^y")

            self._update_context_cache()
            after = dict(self._ctx)
            self._log_before_after("redo", before, after)

        except Exception as e:
            logger.error("[ExcelWorker] redo failed: %s", e, exc_info=True)

    def _select_all(self):
        try:
            wb = self._active_book()
            if not wb:
                logger.info("[ExcelWorker] select_all ignored (no active book)")
                return

            before = self._snapshot_ctx_from_com()

            logger.info("[ExcelWorker] select_all")
            self._app.Cells.Select()

            self._update_context_cache()
            after = dict(self._ctx)
            self._log_before_after("select_all", before, after)

        except Exception as e:
            logger.error("[ExcelWorker] select_all failed: %s", e, exc_info=True)

    def _fill_down(self):
        try:
            wb = self._active_book()
            if not wb:
                logger.info("[ExcelWorker] fill_down ignored (no active book)")
                return

            before = self._snapshot_ctx_from_com()

            logger.info("[ExcelWorker] fill_down")
            self._app.Selection.FillDown()

            self._update_context_cache()
            after = dict(self._ctx)
            self._log_before_after("fill_down", before, after)

        except Exception as e:
            logger.error("[ExcelWorker] fill_down failed: %s", e, exc_info=True)

    def _fill_right(self):
        try:
            wb = self._active_book()
            if not wb:
                logger.info("[ExcelWorker] fill_right ignored (no active book)")
                return

            before = self._snapshot_ctx_from_com()

            logger.info("[ExcelWorker] fill_right")
            self._app.Selection.FillRight()

            self._update_context_cache()
            after = dict(self._ctx)
            self._log_before_after("fill_right", before, after)

        except Exception as e:
            logger.error("[ExcelWorker] fill_right failed: %s", e, exc_info=True)

    # ===============================
    # shutdown
    # ===============================
    def _shutdown(self):
        logger.info("[ExcelWorker] _shutdown begin")

        for p, wb in list(self._books.items()):
            try:
                logger.info("[ExcelWorker] closing book path=%s", p)
                wb.Close(SaveChanges=False)
                try:
                    self.book_closed.emit(p)
                except Exception:
                    pass
            except Exception as e:
                logger.error("[ExcelWorker] close book failed path=%s err=%s", p, e, exc_info=True)
                try:
                    self.book_close_failed.emit(p, str(e))
                except Exception:
                    pass

        self._books.clear()
        self._active_path = ""

        try:
            if self._app:
                logger.info("[ExcelWorker] app quit")
                self._app.Quit()
        except Exception as e:
            logger.error("[ExcelWorker] app quit failed: %s", e, exc_info=True)

        self._app = None
        self._ctx = {"address": "", "sheet": "", "workbook": ""}

        logger.info("[ExcelWorker] _shutdown done")
