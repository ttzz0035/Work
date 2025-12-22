# excel_worker.py
from __future__ import annotations

import os
import queue
import traceback
from typing import Dict, Optional, Any, Tuple

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

    - UI スレッドからは request_* でキュー投入のみ
    - worker thread が順番に実行
    - COM は worker thread だけが触る（重要）
    - get_active_context は COM に触らずキャッシュを返す（重要）
    """

    sheets_ready = Signal(str, list)
    active_cell_changed = Signal(str)

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
        self._ctx: Dict[str, str] = {"address": "", "sheet": "", "workbook": ""}

        logger.info("[ExcelWorker] initialized")

    # ===============================
    # public API（UI スレッド）
    # ===============================
    # --- 互換: 旧UIが open/close を投げても動く ---
    def request_open(self, path: str):
        self._cmd_q.put(("open", os.path.abspath(path)))

    def request_close(self, path: str):
        self._cmd_q.put(("close", os.path.abspath(path)))

    # --- 正式: UI(意味)としてはこちらを使ってOK ---
    def request_open_book(self, path: str):
        self._cmd_q.put(("open_book", os.path.abspath(path)))

    def request_close_book(self, path: str):
        self._cmd_q.put(("close_book", os.path.abspath(path)))

    def request_list_sheets(self, path: str):
        self._cmd_q.put(("list_sheets", os.path.abspath(path)))

    def request_activate_book(self, path: str, front: bool = False):
        self._cmd_q.put(("activate_book", os.path.abspath(path), bool(front)))

    def request_activate_sheet(self, path: str, sheet: str, front: bool = False):
        self._cmd_q.put(("activate_sheet", os.path.abspath(path), str(sheet), bool(front)))

    def request_select_cell(self, cell: str):
        self._cmd_q.put(("select_cell", str(cell)))

    def request_set_cell_value(self, cell: str, value: Any):
        self._cmd_q.put(("set_cell_value", str(cell), value))

    def request_move_cell(self, direction: str, step: int = 1):
        self._cmd_q.put(("move_cell", str(direction), int(step)))

    def request_select_move(self, direction: str):
        self._cmd_q.put(("select_move", str(direction)))

    def request_move_edge(self, direction: str):
        self._cmd_q.put(("move_edge", str(direction)))

    def request_select_edge(self, direction: str):
        self._cmd_q.put(("select_edge", str(direction)))

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

    # --- 停止API ---
    def request_stop(self):
        logger.info("[ExcelWorker] request_stop")
        self._running = False
        self._cmd_q.put(("quit",))

    def shutdown(self, confirm_save: bool = True):
        logger.info(f"[ExcelWorker] shutdown(confirm_save={confirm_save})")
        self.request_stop()

    # ===============================
    # UI polling API（COMに触らない）
    # ===============================
    def get_active_context(self) -> Dict[str, str]:
        """
        Inspector / TreeView から呼ばれる想定。
        COMには触らず、キャッシュのみ返す（安全）。
        """
        try:
            logger.debug(f"[CUT] get_active_context return ctx={self._ctx}")
            return dict(self._ctx)
        except Exception as e:
            logger.error(f"[CUT] get_active_context failed err={e}")
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
                    # getattrディスパッチ（命名規約: _{op}）
                    fn = getattr(self, f"_{op}", None)
                    if fn is None:
                        logger.warning(f"[ExcelWorker] unknown op={op} args={args}")
                        continue
                    fn(*args)

                except Exception as e:
                    tb = traceback.format_exc()
                    logger.error(f"[ExcelWorker] op failed op={op} err={e}\n{tb}")

        finally:
            try:
                self._shutdown()
            except Exception as e:
                tb = traceback.format_exc()
                logger.error(f"[ExcelWorker] shutdown failed err={e}\n{tb}")

            pythoncom.CoUninitialize()
            logger.info("[ExcelWorker] thread stopped (COM uninitialized)")

            try:
                self.quit_finished.emit()
            except Exception as e:
                logger.error(f"[ExcelWorker] quit_finished emit failed err={e}")

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

    def _try_fix_active_book(self) -> Optional[object]:
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
        try:
            self._ctx = self._snapshot_ctx_from_com()
            addr = self._ctx.get("address", "")
            if addr:
                try:
                    self.active_cell_changed.emit(addr)
                except Exception:
                    pass
        except Exception as e:
            tb = traceback.format_exc()
            logger.error(f"[ExcelWorker] update_context_cache failed err={e}\n{tb}")

    def _log_before_after(self, op: str, before: Dict[str, str], after: Dict[str, str], extra: str = ""):
        logger.info(
            f"[ExcelWorker] {op} {extra} "
            f"before={before.get('sheet','')}!{before.get('address','')}({before.get('workbook','')}) "
            f"after={after.get('sheet','')}!{after.get('address','')}({after.get('workbook','')})"
        )

    # ===============================
    # op aliases（根本修正: 意味操作をWorkerが理解する）
    # ===============================
    def _open_book(self, path: str):
        # 正式名
        self._open(path)

    def _close_book(self, path: str):
        # 正式名
        self._close(path)

    # 互換（旧名）
    def _open(self, path: str):
        try:
            ap = os.path.abspath(path)
            if ap in self._books:
                logger.info(f"[ExcelWorker] open ignored (already opened) path={ap}")
                self._active_path = ap
                try:
                    self._books[ap].Activate()
                except Exception:
                    pass
                self._update_context_cache()
                return

            if not os.path.exists(ap):
                logger.warning(f"[ExcelWorker] open ignored (not exists) path={ap}")
                return

            self._ensure_app()
            logger.info(f"[ExcelWorker] open path={ap}")

            wb = self._app.Workbooks.Open(ap)
            self._books[ap] = wb
            self._active_path = ap

            try:
                wb.Activate()
            except Exception:
                pass

            self._update_context_cache()

        except Exception as e:
            tb = traceback.format_exc()
            logger.error(f"[ExcelWorker] open failed path={path} err={e}\n{tb}")

    def _close(self, path: str):
        try:
            ap = os.path.abspath(path)
            wb = self._books.pop(ap, None)
            if wb:
                logger.info(f"[ExcelWorker] close path={ap}")
                wb.Close(SaveChanges=False)

            if self._active_path == ap:
                self._active_path = ""

            self._update_context_cache()

            try:
                self.book_closed.emit(ap)
            except Exception:
                pass

        except Exception as e:
            tb = traceback.format_exc()
            logger.error(f"[ExcelWorker] close failed path={path} err={e}\n{tb}")
            try:
                self.book_close_failed.emit(os.path.abspath(path), str(e))
            except Exception:
                pass

    def _list_sheets(self, path: str):
        try:
            ap = os.path.abspath(path)
            wb = self._books.get(ap)
            if not wb:
                logger.warning(f"[ExcelWorker] list_sheets ignored (not opened) path={ap}")
                return

            sheets = [ws.Name for ws in wb.Worksheets]
            logger.info(f"[ExcelWorker] list_sheets path={ap} count={len(sheets)}")
            try:
                self.sheets_ready.emit(ap, sheets)
            except Exception:
                pass

        except Exception as e:
            tb = traceback.format_exc()
            logger.error(f"[ExcelWorker] list_sheets failed path={path} err={e}\n{tb}")

    def _activate_book(self, path: str, front: bool = False):
        try:
            ap = os.path.abspath(path)
            wb = self._books.get(ap)
            if not wb:
                logger.warning(f"[ExcelWorker] activate_book ignored (not opened) path={ap}")
                return

            logger.info(f"[ExcelWorker] activate_book path={ap} front={front}")
            wb.Activate()
            self._active_path = ap
            self._update_context_cache()

        except Exception as e:
            tb = traceback.format_exc()
            logger.error(f"[ExcelWorker] activate_book failed path={path} err={e}\n{tb}")

    def _activate_sheet(self, path: str, sheet: str, front: bool = False):
        try:
            ap = os.path.abspath(path)
            wb = self._books.get(ap)
            if not wb:
                logger.warning(f"[ExcelWorker] activate_sheet ignored (not opened) path={ap}")
                return

            logger.info(f"[ExcelWorker] activate_sheet path={ap} sheet={sheet} front={front}")
            wb.Activate()
            self._active_path = ap
            wb.Worksheets(sheet).Activate()
            self._update_context_cache()

        except Exception as e:
            tb = traceback.format_exc()
            logger.error(f"[ExcelWorker] activate_sheet failed path={path} sheet={sheet} err={e}\n{tb}")

    def _select_cell(self, cell: str):
        try:
            wb = self._active_book()
            if not wb:
                logger.warning(f"[ExcelWorker] select_cell ignored (no active book) cell={cell}")
                return

            before = self._snapshot_ctx_from_com()

            logger.info(f"[ExcelWorker] select_cell cell={cell}")
            self._app.Range(cell).Select()

            self._update_context_cache()
            after = dict(self._ctx)
            self._log_before_after("select_cell", before, after, extra=f"cell={cell}")

        except Exception as e:
            tb = traceback.format_exc()
            logger.error(f"[ExcelWorker] select_cell failed cell={cell} err={e}\n{tb}")

    def _set_cell_value(self, cell: str, value: Any):
        try:
            wb = self._active_book()
            if not wb:
                logger.warning("[ExcelWorker] set_cell_value ignored (no active book)")
                return

            before = self._snapshot_ctx_from_com()

            logger.info(f"[ExcelWorker] set_cell_value cell={cell} value={value}")

            if cell == "*":
                self._app.Selection.Value = value
            else:
                self._app.Range(cell).Value = value

            self._update_context_cache()
            after = dict(self._ctx)
            self._log_before_after("set_cell_value", before, after, extra=f"cell={cell}")

        except Exception as e:
            tb = traceback.format_exc()
            logger.error(f"[ExcelWorker] set_cell_value failed cell={cell} err={e}\n{tb}")

    def _move_cell(self, direction: str, step: int):
        try:
            wb = self._active_book()
            if not wb:
                logger.warning("[CUT] EXCEL_MOVE skipped (no book)")
                return

            before = self._snapshot_ctx_from_com()
            logger.warning(f"[CUT] EXCEL_MOVE ENTER dir={direction} step={step} before={before}")

            key = self._SK_ARROW.get(direction)
            if not key:
                logger.warning(f"[CUT] EXCEL_MOVE invalid direction={direction}")
                return

            for i in range(int(step)):
                logger.warning(f"[CUT] EXCEL_MOVE SENDKEY i={i} key={key}")
                self._app.SendKeys(key)

            self._update_context_cache()
            after = dict(self._ctx)
            logger.warning(f"[CUT] EXCEL_MOVE EXIT after={after}")

        except Exception as e:
            tb = traceback.format_exc()
            logger.error(f"[ExcelWorker] move_cell failed dir={direction} step={step} err={e}\n{tb}")

    def _select_move(self, direction: str):
        """
        Shift + Arrow
        """
        try:
            wb = self._active_book()
            if not wb:
                logger.warning("[ExcelWorker] select_move ignored (no active book)")
                return

            before = self._snapshot_ctx_from_com()

            key = self._SK_ARROW.get(direction)
            if not key:
                logger.warning(f"[ExcelWorker] select_move invalid direction={direction}")
                return

            logger.info(f"[ExcelWorker] select_move dir={direction} sendkeys=+{key}")
            self._app.SendKeys("+" + key)

            self._update_context_cache()
            after = dict(self._ctx)
            self._log_before_after("select_move", before, after, extra=f"dir={direction}")

        except Exception as e:
            tb = traceback.format_exc()
            logger.error(f"[ExcelWorker] select_move failed dir={direction} err={e}\n{tb}")

    def _move_edge(self, direction: str):
        """
        Ctrl + Arrow
        """
        try:
            wb = self._active_book()
            if not wb:
                logger.warning("[ExcelWorker] move_edge ignored (no active book)")
                return

            before = self._snapshot_ctx_from_com()

            xl_dir = self._XL_DIR.get(direction)
            if xl_dir is None:
                logger.warning(f"[ExcelWorker] move_edge invalid direction={direction}")
                return

            logger.info(f"[ExcelWorker] move_edge dir={direction}")
            self._app.ActiveCell.End(xl_dir).Select()

            self._update_context_cache()
            after = dict(self._ctx)
            self._log_before_after("move_edge", before, after, extra=f"dir={direction}")

        except Exception as e:
            tb = traceback.format_exc()
            logger.error(f"[ExcelWorker] move_edge failed dir={direction} err={e}\n{tb}")

    def _select_edge(self, direction: str):
        """
        Ctrl + Shift + Arrow
        """
        try:
            wb = self._active_book()
            if not wb:
                logger.warning("[ExcelWorker] select_edge ignored (no active book)")
                return

            before = self._snapshot_ctx_from_com()

            key = self._SK_ARROW.get(direction)
            if not key:
                logger.warning(f"[ExcelWorker] select_edge invalid direction={direction}")
                return

            logger.info(f"[ExcelWorker] select_edge dir={direction} sendkeys=^+{key}")
            self._app.SendKeys("^+" + key)

            self._update_context_cache()
            after = dict(self._ctx)
            self._log_before_after("select_edge", before, after, extra=f"dir={direction}")

        except Exception as e:
            tb = traceback.format_exc()
            logger.error(f"[ExcelWorker] select_edge failed dir={direction} err={e}\n{tb}")

    def _copy(self):
        try:
            wb = self._active_book()
            if not wb:
                logger.warning("[ExcelWorker] copy ignored (no active book)")
                return
            logger.info("[ExcelWorker] copy")
            self._app.Selection.Copy()
        except Exception as e:
            tb = traceback.format_exc()
            logger.error(f"[ExcelWorker] copy failed err={e}\n{tb}")

    def _cut(self):
        try:
            wb = self._active_book()
            if not wb:
                logger.warning("[ExcelWorker] cut ignored (no active book)")
                return
            logger.info("[ExcelWorker] cut")
            self._app.Selection.Cut()
        except Exception as e:
            tb = traceback.format_exc()
            logger.error(f"[ExcelWorker] cut failed err={e}\n{tb}")

    def _paste(self):
        try:
            wb = self._active_book()
            if not wb:
                logger.warning("[ExcelWorker] paste ignored (no active book)")
                return

            before = self._snapshot_ctx_from_com()

            logger.info("[ExcelWorker] paste")
            self._app.ActiveSheet.Paste()

            self._update_context_cache()
            after = dict(self._ctx)
            self._log_before_after("paste", before, after)

        except Exception as e:
            tb = traceback.format_exc()
            logger.error(f"[ExcelWorker] paste failed err={e}\n{tb}")

    def _undo(self):
        try:
            wb = self._active_book()
            if not wb:
                logger.warning("[ExcelWorker] undo ignored (no active book)")
                return

            before = self._snapshot_ctx_from_com()

            logger.info("[ExcelWorker] undo")
            self._app.Undo()

            self._update_context_cache()
            after = dict(self._ctx)
            self._log_before_after("undo", before, after)

        except Exception as e:
            tb = traceback.format_exc()
            logger.error(f"[ExcelWorker] undo failed err={e}\n{tb}")

    def _redo(self):
        try:
            wb = self._active_book()
            if not wb:
                logger.warning("[ExcelWorker] redo ignored (no active book)")
                return

            before = self._snapshot_ctx_from_com()

            logger.info("[ExcelWorker] redo (SendKeys ^y)")
            self._app.SendKeys("^y")

            self._update_context_cache()
            after = dict(self._ctx)
            self._log_before_after("redo", before, after)

        except Exception as e:
            tb = traceback.format_exc()
            logger.error(f"[ExcelWorker] redo failed err={e}\n{tb}")

    def _select_all(self):
        try:
            wb = self._active_book()
            if not wb:
                logger.warning("[ExcelWorker] select_all ignored (no active book)")
                return

            before = self._snapshot_ctx_from_com()

            logger.info("[ExcelWorker] select_all")
            self._app.Cells.Select()

            self._update_context_cache()
            after = dict(self._ctx)
            self._log_before_after("select_all", before, after)

        except Exception as e:
            tb = traceback.format_exc()
            logger.error(f"[ExcelWorker] select_all failed err={e}\n{tb}")

    def _fill_down(self):
        try:
            wb = self._active_book()
            if not wb:
                logger.warning("[ExcelWorker] fill_down ignored (no active book)")
                return

            before = self._snapshot_ctx_from_com()

            logger.info("[ExcelWorker] fill_down")
            self._app.Selection.FillDown()

            self._update_context_cache()
            after = dict(self._ctx)
            self._log_before_after("fill_down", before, after)

        except Exception as e:
            tb = traceback.format_exc()
            logger.error(f"[ExcelWorker] fill_down failed err={e}\n{tb}")

    def _fill_right(self):
        try:
            wb = self._active_book()
            if not wb:
                logger.warning("[ExcelWorker] fill_right ignored (no active book)")
                return

            before = self._snapshot_ctx_from_com()

            logger.info("[ExcelWorker] fill_right")
            self._app.Selection.FillRight()

            self._update_context_cache()
            after = dict(self._ctx)
            self._log_before_after("fill_right", before, after)

        except Exception as e:
            tb = traceback.format_exc()
            logger.error(f"[ExcelWorker] fill_right failed err={e}\n{tb}")

    def _quit(self):
        logger.info("[ExcelWorker] quit received")
        self._running = False

    # ===============================
    # shutdown
    # ===============================
    def _shutdown(self):
        logger.info("[ExcelWorker] _shutdown begin")

        for p, wb in list(self._books.items()):
            try:
                logger.info(f"[ExcelWorker] closing book path={p}")
                wb.Close(SaveChanges=False)
                try:
                    self.book_closed.emit(p)
                except Exception:
                    pass
            except Exception as e:
                tb = traceback.format_exc()
                logger.error(f"[ExcelWorker] close book failed path={p} err={e}\n{tb}")
                try:
                    self.book_close_failed.emit(p, str(e))
                except Exception:
                    pass

        self._books.clear()
        self._active_path = ""

        if self._app:
            try:
                logger.info("[ExcelWorker] app quit")
                self._app.Quit()
            except Exception as e:
                tb = traceback.format_exc()
                logger.error(f"[ExcelWorker] app quit failed err={e}\n{tb}")

        self._app = None
        self._ctx = {"address": "", "sheet": "", "workbook": ""}

        logger.info("[ExcelWorker] _shutdown done")
