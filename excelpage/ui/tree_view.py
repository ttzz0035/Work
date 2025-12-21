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

from ui.hover_action_delegate import HoverActionDelegate


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
            logger.exception("DiffThread failed: %s", e)
            self.finished_ng.emit(str(e))


# =====================================================
# Tree View
# =====================================================
class LauncherTreeView(QTreeView):
    """
    Tree:
      group(folder)  ※仮想グループ（OSパスとは無関係）
        └─ file (Excel)
             └─ sheet

    - Excel 切替・Sheet 切替
    - キー操作（Enter）
    - Macro 録画（意味コマンド）
    - グループ追加/フォルダ読込：グループ名入力（空ならフォルダ名）
    - DnD/リネームはUIのみ（OS移動しない）
    - ホバー時に編集/削除ボタン表示（delegate）
    """

    # -------------------------------------------------
    # Init
    # -------------------------------------------------
    def __init__(self, parent=None):
        super().__init__(parent)

        # --- model ---
        self._model = QStandardItemModel()
        self._model.setHorizontalHeaderLabels([""])
        self.setModel(self._model)
        self.header().hide()

        self.setSelectionMode(QTreeView.ExtendedSelection)
        self.setSelectionBehavior(QTreeView.SelectRows)

        # 編集：F2/選択クリック（sheetだけ禁止にするのは item 側で制御）
        self.setEditTriggers(
            QTreeView.EditKeyPressed |
            QTreeView.SelectedClicked
        )

        self.setUniformRowHeights(True)
        self.setAnimated(True)

        # --- DnD ---
        self.setDragEnabled(True)
        self.setAcceptDrops(True)
        self.setDropIndicatorShown(True)
        self.setDragDropMode(QAbstractItemView.InternalMove)

        # --- hover tracking ---
        self.setMouseTracking(True)

        # --- delegate (hover action buttons) ---
        self._hover_delegate = HoverActionDelegate(ROLE_TAG, self)
        self.setItemDelegate(self._hover_delegate)
        self.entered.connect(self._on_item_entered)
        self._hover_delegate.edit_requested.connect(self._on_hover_edit_clicked)
        self._hover_delegate.delete_requested.connect(self._on_hover_delete_clicked)

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
        self._inspector = InspectorPanel()
        self._inspector.set_tree(self)
        self._inspector.hide()

        self._excel.active_cell_changed.connect(self._inspector.set_current_cell)

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
    # Hover handling
    # =================================================
    def _on_item_entered(self, idx: QModelIndex):
        try:
            if idx and idx.isValid():
                self._hover_delegate.set_hover_index(idx)
            else:
                self._hover_delegate.clear_hover_index()
            self.viewport().update()
        except Exception as e:
            logger.error("[Hover] entered handler failed: %s", e, exc_info=True)

    def leaveEvent(self, event):
        try:
            self._hover_delegate.clear_hover_index()
            self.viewport().update()
        except Exception as e:
            logger.error("[Hover] leaveEvent failed: %s", e, exc_info=True)
        super().leaveEvent(event)

    def _on_hover_edit_clicked(self, idx: QModelIndex):
        try:
            if not idx or not idx.isValid():
                return
            tag = idx.data(ROLE_TAG)
            if not isinstance(tag, NodeTag):
                return

            # sheet は編集禁止
            if tag.kind == "sheet":
                logger.info("[Hover] edit ignored (sheet)")
                return

            logger.info("[Hover] edit start kind=%s text=%s", tag.kind, str(idx.data()))
            self._set_single_selection(idx)
            self.edit(idx)

        except Exception as e:
            logger.error("[Hover] edit clicked failed: %s", e, exc_info=True)

    def _on_hover_delete_clicked(self, idx: QModelIndex):
        try:
            if not idx or not idx.isValid():
                return

            tag = idx.data(ROLE_TAG)
            if not isinstance(tag, NodeTag):
                return

            # sheet は削除禁止（安全）
            if tag.kind == "sheet":
                logger.info("[Hover] delete ignored (sheet)")
                return

            item = self._model.itemFromIndex(idx)
            if item is None:
                return

            name = item.text()
            if tag.kind == "folder":
                msg = f"グループ「{name}」を削除しますか？\n（配下のノードも削除されます）"
            elif tag.kind == "file":
                msg = f"ファイル「{name}」をツリーから削除しますか？\n（Excelファイル自体は削除しません）"
            else:
                msg = f"「{name}」を削除しますか？"

            ret = QMessageBox.question(
                self,
                "Remove",
                msg,
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.No,
            )
            if ret != QMessageBox.Yes:
                logger.info("[Hover] delete canceled by user name=%s", name)
                return

            logger.info("[Hover] delete confirmed kind=%s name=%s", tag.kind, name)
            self._remove_item_by_index(idx)

        except Exception as e:
            logger.error("[Hover] delete clicked failed: %s", e, exc_info=True)

    def _remove_item_by_index(self, idx: QModelIndex):
        try:
            if not idx or not idx.isValid():
                return
            item = self._model.itemFromIndex(idx)
            if item is None:
                return

            parent = item.parent()
            if parent is None:
                root = self._model.invisibleRootItem()
                root.removeRow(item.row())
            else:
                parent.removeRow(item.row())

            logger.info("[Remove] removed row=%s", idx.row())

        except Exception as e:
            logger.error("[Remove] remove failed: %s", e, exc_info=True)

    # =================================================
    # Project Export / Import
    # =================================================
    def export_project(self) -> Dict[str, Any]:
        logger.info("[Project] export start")
        root = self._model.invisibleRootItem()
        groups: List[Dict[str, Any]] = []
        for r in range(root.rowCount()):
            groups.append(self._export_item_recursive(root.child(r)))
        out = {
            "version": 1,
            "groups": groups,
        }
        logger.info("[Project] export completed groups=%s", len(groups))
        return out

    def _export_item_recursive(self, item: QStandardItem) -> Dict[str, Any]:
        tag = item.data(ROLE_TAG)
        tag_dict = None
        if isinstance(tag, NodeTag):
            try:
                tag_dict = tag.to_dict()
            except Exception as e:
                logger.error("[Project] tag.to_dict failed: %s", e, exc_info=True)
                tag_dict = {
                    "kind": getattr(tag, "kind", ""),
                    "path": getattr(tag, "path", ""),
                    "sheet": getattr(tag, "sheet", ""),
                }

        node = {
            "text": item.text(),
            "tag": tag_dict,
            "children": [],
        }

        for r in range(item.rowCount()):
            node["children"].append(self._export_item_recursive(item.child(r)))

        return node

    def import_project(self, data: Dict[str, Any]) -> None:
        logger.info("[Project] import start")
        if not isinstance(data, dict):
            raise ValueError("project data is not dict")

        self._model.clear()
        self._model.setHorizontalHeaderLabels([""])

        groups = data.get("groups", [])
        if not isinstance(groups, list):
            raise ValueError("project.groups is not list")

        for g in groups:
            it = self._import_item_recursive(g)
            self._model.appendRow(it)
            self.expand(it.index())

        logger.info("[Project] import completed groups=%s", len(groups))

    def _import_item_recursive(self, node: Dict[str, Any]) -> QStandardItem:
        if not isinstance(node, dict):
            raise ValueError("project node is not dict")

        text = str(node.get("text", ""))
        tag_data = node.get("tag", None)

        tag: NodeTag
        if isinstance(tag_data, dict):
            try:
                tag = NodeTag.from_dict(tag_data)
            except Exception as e:
                logger.error("[Project] NodeTag.from_dict failed: %s", e, exc_info=True)
                kind = str(tag_data.get("kind", "folder"))
                path = str(tag_data.get("path", ""))
                sheet = str(tag_data.get("sheet", ""))
                tag = NodeTag(kind, path, sheet)
        else:
            tag = NodeTag("folder", "")

        it = self._create_item(text, tag)

        children = node.get("children", [])
        if isinstance(children, list):
            for ch in children:
                it.appendRow(self._import_item_recursive(ch))

        if isinstance(tag, NodeTag) and tag.kind == "file":
            try:
                self._engine_exec("open_book", path=tag.path)
            except Exception as e:
                logger.error("[Project] open_book failed: %s", e, exc_info=True)

        return it

    # =================================================
    # Engine (唯一の実行入口)
    # =================================================
    def _engine_exec(self, op: str, source: Optional[str] = None, **kwargs):
        """
        source:
          - "inspector" : ★ マクロ記録対象
          - "macro"     : マクロ再生（記録しない）
          - None        : 通常UI操作（記録しない）
        """
        logger.info("[ENGINE] %s source=%s %s", op, source, kwargs)

        # =================================================
        # ★ Macro record : Inspector ONLY
        # =================================================
        if source == "inspector":
            try:
                if self._macro.is_recording():
                    self._macro.record(op, **kwargs)
            except Exception as e:
                logger.error("macro record failed: %s", e)

        # =================================================
        # Execute
        # =================================================
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
                kwargs.get("cell", "*"),
                kwargs.get("value", ""),
            )

        elif op == "move_cell":
            self._excel.request_move_cell(
                kwargs["direction"], kwargs.get("step", 1)
            )

        elif op == "move_edge":
            self._excel.request_move_edge(kwargs["direction"])

        elif op == "select_move":
            self._excel.request_select_move(kwargs["direction"])

        elif op == "select_edge":
            self._excel.request_select_edge(kwargs["direction"])

        elif op == "copy":
            self._excel.request_copy()

        elif op == "paste":
            self._excel.request_paste()

        elif op == "cut":
            self._excel.request_cut()

        elif op == "undo":
            self._excel.request_undo()

        elif op == "redo":
            self._excel.request_redo()

        elif op == "select_all":
            self._excel.request_select_all()

        elif op == "fill_down":
            self._excel.request_fill_down()

        elif op == "fill_right":
            self._excel.request_fill_right()

        elif op == "get_active_context":
            return self._excel.get_active_context()

        else:
            logger.debug("Non-exec op: %s", op)

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

        # Delete キーで削除（sheet は無視）
        if key == Qt.Key_Delete and not (mod & (Qt.ControlModifier | Qt.ShiftModifier | Qt.AltModifier)):
            idx = self.currentIndex()
            if idx and idx.isValid():
                tag = idx.data(ROLE_TAG)
                if isinstance(tag, NodeTag) and tag.kind != "sheet":
                    logger.info("[Key] Delete pressed -> remove")
                    self._on_hover_delete_clicked(idx)
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

        # ★ グループ追加
        menu.addAction("Add Group...", self.add_group_dialog)

        menu.addSeparator()
        menu.addAction("Add Files...", self.add_files_dialog)
        menu.addAction("Add Folder...", self.add_folder_dialog)

        menu.addSeparator()
        menu.addAction("Inspector (Record Mode)", self._open_inspector)

        sheets = self._get_selected_sheet_tags()
        if len(sheets) == 2:
            menu.addSeparator()
            menu.addAction("Diff", self._run_diff_two_sheets)

        menu.exec(self.viewport().mapToGlobal(pos))

    # =================================================
    # Group (Virtual)
    # =================================================
    def add_group_dialog(self):
        name, ok = QInputDialog.getText(
            self,
            "Add Group",
            "Group name:",
            text="New Group",
        )
        if not ok:
            logger.info("[GROUP] add canceled")
            return

        group_name = (name or "").strip()
        if not group_name:
            logger.info("[GROUP] add empty -> ignored")
            return

        it = self._create_item(group_name, NodeTag("folder", ""))
        self._model.appendRow(it)
        self.expand(it.index())

        logger.info("[GROUP] added: %s", group_name)
        try:
            self._macro.record("add_group", name=group_name)
        except Exception:
            pass

    # =================================================
    # Macro UI
    # =================================================
    def macro_start_dialog(self):
        name, ok = QInputDialog.getText(self, "Macro", "Macro name:", text="macro")
        if ok:
            self._macro.start(name=name or "macro")
            QMessageBox.information(self, "Macro", "録画開始")
            logger.info("[MACRO] start name=%s", name or "macro")

    def macro_stop(self):
        self._macro.stop()
        QMessageBox.information(
            self, "Macro", f"録画停止（{self._macro.steps_count()} steps）"
        )
        logger.info("[MACRO] stop steps=%s", self._macro.steps_count())

    def macro_clear(self):
        self._macro.clear()
        QMessageBox.information(self, "Macro", "コマンドをクリアしました")
        logger.info("[MACRO] cleared")

    def macro_save_dialog(self):
        path, _ = QFileDialog.getSaveFileName(
            self, "Save Macro", "", "Macro JSON (*.json)"
        )
        if path:
            try:
                self._macro.save_json(path)
                QMessageBox.information(self, "Macro", f"保存しました:\n{path}")
                logger.info("[MACRO] saved: %s", path)
            except Exception as e:
                logger.exception("Macro save failed: %s", e)
                QMessageBox.critical(self, "Macro", f"保存に失敗しました:\n{e}")

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
            self._engine_exec("activate_sheet", path=tag.path, sheet=tag.sheet)

        else:
            # folder/group 選択は Excel 操作しない
            pass

    def _on_selection_changed(self, *_):
        self._execute_current_selection()

    def _on_sheets_ready(self, path: str, sheets):
        item = self._find_file_item(path)
        if not item:
            return

        # clear sheet children
        for r in reversed(range(item.rowCount())):
            ch = item.child(r)
            tag = ch.data(ROLE_TAG)
            if isinstance(tag, NodeTag) and tag.kind == "sheet":
                item.removeRow(r)

        for s in sheets:
            item.appendRow(self._create_item(s, NodeTag("sheet", path, s)))

        self.expand(item.index())
        logger.info("[SHEETS] ready path=%s count=%s", path, len(sheets))

    # =================================================
    # Diff
    # =================================================
    def _get_selected_sheet_tags(self) -> List[NodeTag]:
        out: List[NodeTag] = []
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

        logger.info("[DIFF] start a=%s/%s b=%s/%s", a.path, a.sheet, b.path, b.sheet)

    def _on_diff_ok(self, payload: dict):
        if self._diff_dialog is None:
            from ui.diff_dialog import DiffDialog
            self._diff_dialog = DiffDialog(self)
        self._diff_dialog.set_results(payload.get("items", []), payload.get("meta", {}))
        self._diff_dialog.exec()
        logger.info("[DIFF] ok items=%s", len(payload.get("items", []) or []))

    def _on_diff_ng(self, msg: str):
        logger.error("[DIFF] ng: %s", msg)
        QMessageBox.critical(self, "Diff", msg)

    # =================================================
    # Add (Files / Folder)
    # =================================================
    def add_files_dialog(self):
        files, _ = QFileDialog.getOpenFileNames(
            self, "Add Excel Files", "", "Excel (*.xlsx *.xlsm *.xls *.xlsb)"
        )
        for f in files:
            self._add_file(f)

    def add_folder_dialog(self):
        folder = QFileDialog.getExistingDirectory(self, "Add Folder")
        if not folder:
            return

        default_name = os.path.basename(folder)
        name, ok = QInputDialog.getText(
            self,
            "Group Name",
            "Group name:",
            text=default_name,
        )

        # 入力なければフォルダ名
        group_name = (name or "").strip() if ok else ""
        if not group_name:
            group_name = default_name

        logger.info("[FOLDER] selected=%s group_name=%s", folder, group_name)
        self._add_folder(folder, group_name)

    def _add_file(self, path: str, parent: Optional[QStandardItem] = None):
        if not _is_openable_excel_path(path):
            return

        ap = _abspath(path)
        it = self._create_item(os.path.basename(ap), NodeTag("file", ap))

        (parent or self._model).appendRow(it)

        logger.info("[FILE] add path=%s parent=%s", ap, "yes" if parent else "root")
        self._engine_exec("open_book", path=ap)

    def _add_folder(self, folder: str, group_name: str):
        # ★ 仮想グループ（OSパスは保持しない方針）
        root = self._create_item(group_name, NodeTag("folder", ""))
        self._model.appendRow(root)

        for name in sorted(os.listdir(folder)):
            full = os.path.join(folder, name)
            if _is_openable_excel_path(full):
                self._add_file(full, root)

        self.expand(root.index())
        logger.info("[GROUP] folder imported folder=%s -> group=%s", folder, group_name)

        # Macroに残す（危険操作じゃない）
        try:
            self._macro.record("import_folder", folder=folder, group=group_name)
        except Exception:
            pass

    # =================================================
    # Helpers
    # =================================================
    def _append_command_log(self, line: str):
        safe = str(line).encode("utf-8", errors="replace").decode("utf-8")
        logger.info("[CMD] %s", safe)

    def _create_item(self, text: str, tag: NodeTag) -> QStandardItem:
        it = QStandardItem(text)
        it.setData(tag, ROLE_TAG)

        # ★ 編集可否：sheet は不可、それ以外は可
        if isinstance(tag, NodeTag) and tag.kind == "sheet":
            it.setEditable(False)
        else:
            it.setEditable(True)

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
        logger.info("[INSPECTOR] opened")

    # =================================================
    # Shutdown Excel (public)
    # =================================================
    def shutdown_excel_on_exit(self):
        logger.info("[TreeView] shutdown_excel_on_exit begin")

        try:
            if self._inspector:
                self._inspector.close()
        except Exception as e:
            logger.error("[TreeView] inspector close failed: %s", e, exc_info=True)

        try:
            if hasattr(self, "_excel") and self._excel:
                if hasattr(self._excel, "shutdown"):
                    logger.info("[TreeView] ExcelWorker.shutdown(confirm_save=True)")
                    self._excel.shutdown(confirm_save=True)
                elif hasattr(self._excel, "stop"):
                    logger.info("[TreeView] ExcelWorker.stop()")
                    self._excel.stop()
                elif hasattr(self._excel, "request_stop"):
                    logger.info("[TreeView] ExcelWorker.request_stop()")
                    self._excel.request_stop()
                elif hasattr(self._excel, "quit"):
                    logger.info("[TreeView] ExcelWorker.quit()")
                    self._excel.quit()

                try:
                    if self._excel.isRunning():
                        logger.info("[TreeView] ExcelWorker is running -> quit/wait")
                        self._excel.quit()
                        self._excel.wait(3000)
                except Exception as e2:
                    logger.error("[TreeView] ExcelWorker wait failed: %s", e2, exc_info=True)

        except Exception as e:
            logger.exception("[TreeView] shutdown excel failed: %s", e)

        logger.info("[TreeView] shutdown_excel_on_exit done")

    # =================================================
    # Close Event
    # =================================================
    def closeEvent(self, event):
        logger.info("[TreeView] closeEvent begin")
        try:
            self.shutdown_excel_on_exit()
        except Exception as e:
            logger.error("[TreeView] closeEvent shutdown failed: %s", e, exc_info=True)
        logger.info("[TreeView] closeEvent -> super")
        super().closeEvent(event)
