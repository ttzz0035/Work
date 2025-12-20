# ui/tree_view.py
from __future__ import annotations

import os
import json
from typing import List, Tuple, Dict, Optional, Any

from PySide6.QtWidgets import (
    QTreeView,
    QMenu,
    QFileDialog,
    QMessageBox,
    QAbstractItemView,
    QApplication,
    QLabel,
    QDialog,
    QInputDialog,
)
from PySide6.QtGui import QStandardItemModel, QStandardItem
from PySide6.QtCore import (
    Qt,
    QModelIndex,
    QPoint,
    QItemSelectionModel,
    QThread,
    Signal,
)

from models.node_tag import NodeTag
from models.dto import DiffRequest
from excel_worker import ExcelWorker
from logger import get_logger

from ui.diff_dialog import DiffOptionDialog
from ui.inspector_panel import InspectorPanel
from services.macro_recorder import get_macro_recorder


logger = get_logger("TreeView")

EXCEL_EXTS = (".xlsx", ".xlsm", ".xls", ".xlsb")
ROLE_TAG = Qt.ItemDataRole.UserRole + 1


# =====================================================
# Utils
# =====================================================
def _is_openable_excel_path(path: str) -> bool:
    name = os.path.basename(path)
    if not name or name.startswith("~$"):
        return False
    return name.lower().endswith(EXCEL_EXTS)


def _abspath(p: str) -> str:
    return os.path.abspath(p) if p else ""


# =====================================================
# Diff Thread
# =====================================================
class DiffThread(QThread):
    finished_ok = Signal(dict)
    finished_ng = Signal(str)

    def __init__(self, req: DiffRequest, append_log):
        super().__init__()
        self._req = req
        self._append_log = append_log

    def run(self):
        try:
            from services.diff import run_diff

            ctx: Dict[str, Any] = {}
            run_diff(self._req, ctx, logger, self._append_log)
            self.finished_ok.emit(
                {
                    "items": ctx.get("diff_items", []),
                    "meta": ctx.get("diff_meta", {}),
                }
            )
        except Exception as e:
            self.finished_ng.emit(str(e))


# =====================================================
# Tree View
# =====================================================
class LauncherTreeView(QTreeView):
    """
    Tree:
      group(folder)
        └─ file (Excel)
             └─ sheet

    - Excel 切替・Sheet 切替
    - キー操作（Alt↑↓ / Ctrl↑↓ / Enter）
    - Macro 録画（意味コマンド）
    """

    # -------------------------------------------------
    # Init
    # -------------------------------------------------
    def __init__(self, parent=None):
        super().__init__(parent)

        # --- model ---
        self._model = QStandardItemModel()
        self._model.setHorizontalHeaderLabels(["Workspace"])
        self.setModel(self._model)

        self.setSelectionMode(QTreeView.ExtendedSelection)
        self.setSelectionBehavior(QTreeView.SelectRows)
        self.setEditTriggers(QTreeView.NoEditTriggers)
        self.setUniformRowHeights(True)
        self.setAnimated(True)

        # --- DnD ---
        self.setDragEnabled(True)
        self.setAcceptDrops(True)
        self.setDropIndicatorShown(True)
        self.setDragDropMode(QAbstractItemView.InternalMove)

        # --- context menu ---
        self.setContextMenuPolicy(Qt.CustomContextMenu)
        self.customContextMenuRequested.connect(self._on_context_menu)

        # --- Excel worker ---
        self._excel = ExcelWorker()
        self._excel.sheets_ready.connect(self._on_sheets_ready)
        self._excel.start()

        self.selectionModel().selectionChanged.connect(self._on_selection_changed)

        # --- diff ---
        self._diff_thread: Optional[DiffThread] = None
        self._diff_dialog = None

        # --- macro ---
        self._macro = get_macro_recorder()

        # --- inspector ---
        self._inspector = InspectorPanel(self)
        self._inspector.hide()

        # --- busy overlay ---
        self._busy_reason: Optional[str] = None
        self._busy_label = QLabel(self)
        self._busy_label.setVisible(False)
        self._busy_label.setStyleSheet(
            "QLabel { background: rgba(0,0,0,160); color:white; "
            "padding:6px 10px; border-radius:6px; font-weight:bold; }"
        )
        self._busy_label.setAttribute(Qt.WA_TransparentForMouseEvents)

        logger.info("LauncherTreeView initialized")

    # =================================================
    # Engine (唯一の実行入口)
    # =================================================
    def _engine_exec(self, op: str, **kwargs):
        logger.info("[ENGINE] %s %s", op, kwargs)

        # --- record ---
        try:
            self._macro.record(op, **kwargs)
        except Exception as e:
            logger.error("macro record failed: %s", e)

        # --- execute ---
        if op == "open_book":
            self._excel.request_open(kwargs["path"])

        elif op == "close_book":
            self._excel.request_close(kwargs["path"])

        elif op == "activate_book":
            self._excel.request_activate_book(kwargs["path"], front=True)

        elif op == "activate_sheet":
            self._excel.request_activate_sheet(
                kwargs["path"], kwargs["sheet"], front=True
            )

        elif op == "list_sheets":
            self._excel.request_list_sheets(kwargs["path"])

        elif op == "select_cell":
            self._excel.request_select_cell(kwargs["cell"])

        elif op == "set_cell_value":
            self._excel.request_set_cell_value(
                kwargs["cell"], kwargs.get("value", "")
            )

        else:
            logger.error("Unknown engine op: %s", op)

    # =================================================
    # Key handling
    # =================================================
    def keyPressEvent(self, event):
        key = event.key()
        mod = event.modifiers()

        if key in (Qt.Key_Return, Qt.Key_Enter):
            self._execute_current_selection()
            event.accept()
            return

        if (mod & Qt.ControlModifier) and key == Qt.Key_R:
            if self._macro.is_recording():
                self.macro_stop()
            else:
                self.macro_start_dialog()
            event.accept()
            return

        if (mod & Qt.ControlModifier) and key == Qt.Key_S:
            self.macro_save_dialog()
            event.accept()
            return

        super().keyPressEvent(event)

    # =================================================
    # Context Menu
    # =================================================
    def _on_context_menu(self, pos: QPoint):
        idx = self.indexAt(pos)
        if idx.isValid():
            self._set_single_selection(idx)

        menu = QMenu(self)
        menu.addAction("Add Files...", self.add_files_dialog)
        menu.addAction("Add Folder...", self.add_folder_dialog)

        menu.addSeparator()
        menu.addAction("Inspector (Record Mode)", self._open_inspector)

        menu.addSeparator()
        if not self._macro.is_recording():
            menu.addAction("● Macro Start", self.macro_start_dialog)
        else:
            menu.addAction("■ Macro Stop", self.macro_stop)
        menu.addAction("Macro Save...", self.macro_save_dialog)
        menu.addAction("Macro Clear", self.macro_clear)

        sheets = self._get_selected_sheet_tags()
        if len(sheets) == 2:
            menu.addSeparator()
            menu.addAction("Diff", self._run_diff_two_sheets)

        menu.exec(self.viewport().mapToGlobal(pos))

    # =================================================
    # Macro UI
    # =================================================
    def macro_start_dialog(self):
        name, ok = QInputDialog.getText(self, "Macro", "Macro name:", text="macro")
        if ok:
            self._macro.start(name=name or "macro")
            QMessageBox.information(self, "Macro", "録画開始")

    def macro_stop(self):
        self._macro.stop()
        QMessageBox.information(
            self, "Macro", f"録画停止（{self._macro.steps_count()} steps）"
        )

    def macro_clear(self):
        self._macro.clear()
        QMessageBox.information(self, "Macro", "コマンドをクリアしました")

    def macro_save_dialog(self):
        path, _ = QFileDialog.getSaveFileName(
            self, "Save Macro", "", "Macro JSON (*.json)"
        )
        if path:
            self._macro.save_json(path)
            QMessageBox.information(self, "Macro", f"保存しました:\n{path}")

    # =================================================
    # Selection handling
    # =================================================
    def _execute_current_selection(self):
        idx = self.currentIndex()
        if not idx.isValid():
            return
        tag = idx.data(ROLE_TAG)
        if not isinstance(tag, NodeTag):
            return

        if tag.kind == "file":
            self._engine_exec("activate_book", path=tag.path)
            item = self._model.itemFromIndex(idx)
            if item and not self._has_sheet_children(item):
                self._engine_exec("list_sheets", path=tag.path)

        elif tag.kind == "sheet":
            self._engine_exec(
                "activate_sheet", path=tag.path, sheet=tag.sheet
            )

    def _on_selection_changed(self, *_):
        self._execute_current_selection()

    def _on_sheets_ready(self, path: str, sheets):
        item = self._find_file_item(path)
        if not item:
            return

        # clear
        for r in reversed(range(item.rowCount())):
            ch = item.child(r)
            tag = ch.data(ROLE_TAG)
            if isinstance(tag, NodeTag) and tag.kind == "sheet":
                item.removeRow(r)

        for s in sheets:
            item.appendRow(self._create_item(s, NodeTag("sheet", path, s)))

        self.expand(item.index())

    # =================================================
    # Diff
    # =================================================
    def _get_selected_sheet_tags(self) -> List[NodeTag]:
        out = []
        for idx in self._selected_primary_indexes():
            tag = idx.data(ROLE_TAG)
            if isinstance(tag, NodeTag) and tag.kind == "sheet":
                out.append(tag)
        return out

    def _run_diff_two_sheets(self):
        sheets = self._get_selected_sheet_tags()
        if len(sheets) != 2:
            QMessageBox.information(self, "Diff", "シートを2つ選択してください")
            return

        a, b = sheets
        opt_dlg = DiffOptionDialog(self, [], False, True, False)
        if opt_dlg.exec() != QDialog.Accepted:
            return

        opt = opt_dlg.get_options()
        req = DiffRequest(
            file_a=a.path,
            sheet_a=a.sheet,
            file_b=b.path,
            sheet_b=b.sheet,
            key_cols=opt["key_cols"],
            compare_formula=opt["compare_formula"],
            include_context=opt["include_context"],
            compare_shapes=opt["compare_shapes"],
        )

        self._diff_thread = DiffThread(req, self._append_command_log)
        self._diff_thread.finished_ok.connect(self._on_diff_ok)
        self._diff_thread.finished_ng.connect(self._on_diff_ng)
        self._diff_thread.start()

    def _on_diff_ok(self, payload: dict):
        if self._diff_dialog is None:
            from ui.diff_dialog import DiffDialog
            self._diff_dialog = DiffDialog(self)
        self._diff_dialog.set_results(
            payload.get("items", []), payload.get("meta", {})
        )
        self._diff_dialog.exec()

    def _on_diff_ng(self, msg: str):
        QMessageBox.critical(self, "Diff", msg)

    # =================================================
    # Add / Remove
    # =================================================
    def add_files_dialog(self):
        files, _ = QFileDialog.getOpenFileNames(
            self, "Add Excel Files", "", "Excel (*.xlsx *.xlsm *.xls *.xlsb)"
        )
        for f in files:
            self._add_file(f)

    def add_folder_dialog(self):
        folder = QFileDialog.getExistingDirectory(self, "Add Folder")
        if folder:
            self._add_folder(folder)

    def _add_file(self, path: str, parent: Optional[QStandardItem] = None):
        if not _is_openable_excel_path(path):
            return
        ap = _abspath(path)
        it = self._create_item(os.path.basename(ap), NodeTag("file", ap))
        (parent or self._model).appendRow(it)
        self._engine_exec("open_book", path=ap)

    def _add_folder(self, folder: str):
        root = self._create_item(os.path.basename(folder), NodeTag("folder", folder))
        self._model.appendRow(root)
        for name in sorted(os.listdir(folder)):
            full = os.path.join(folder, name)
            if _is_openable_excel_path(full):
                self._add_file(full, root)
        self.expand(root.index())

    # =================================================
    # Helpers
    # =================================================
    def _append_command_log(self, line: str):
        safe = str(line).encode("utf-8", errors="replace").decode("utf-8")
        logger.info("[CMD] %s", safe)

    def _create_item(self, text: str, tag: NodeTag) -> QStandardItem:
        it = QStandardItem(text)
        it.setEditable(False)
        it.setData(tag, ROLE_TAG)
        return it

    def _has_sheet_children(self, item: QStandardItem) -> bool:
        for r in range(item.rowCount()):
            tag = item.child(r).data(ROLE_TAG)
            if isinstance(tag, NodeTag) and tag.kind == "sheet":
                return True
        return False

    def _find_file_item(self, path: str) -> Optional[QStandardItem]:
        target = _abspath(path)
        root = self._model.invisibleRootItem()
        stack = [root.child(r) for r in range(root.rowCount())]
        while stack:
            it = stack.pop()
            tag = it.data(ROLE_TAG)
            if isinstance(tag, NodeTag) and tag.kind == "file":
                if _abspath(tag.path) == target:
                    return it
            for r in range(it.rowCount()):
                stack.append(it.child(r))
        return None

    def _selected_primary_indexes(self) -> List[QModelIndex]:
        sm = self.selectionModel()
        return sm.selectedRows(0) if sm else []

    def _set_single_selection(self, idx: QModelIndex):
        sm = self.selectionModel()
        if sm:
            sm.clearSelection()
            sm.setCurrentIndex(
                idx,
                QItemSelectionModel.ClearAndSelect | QItemSelectionModel.Rows,
            )

    def _open_inspector(self):
        self._inspector.show()
        self._inspector.raise_()
        self._inspector.activateWindow()
