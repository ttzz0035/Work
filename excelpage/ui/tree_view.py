from __future__ import annotations

import os
import json
import time
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
    QProgressDialog,
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

        # ★ 追加：終了時 close 完了を待つ
        if hasattr(self._excel, "book_closed"):
            self._excel.book_closed.connect(self._on_exit_book_closed)
        if hasattr(self._excel, "book_close_failed"):
            self._excel.book_close_failed.connect(self._on_exit_book_close_failed)
        if hasattr(self._excel, "quit_finished"):
            self._excel.quit_finished.connect(self._on_exit_quit_finished)

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

        # --- macro play ---
        self._macro_play_thread = None

        # --- exit state ---
        self._is_shutting_down: bool = False
        self._exit_total: int = 0
        self._exit_done: int = 0
        self._exit_closed: set[str] = set()
        self._exit_failed: set[str] = set()
        self._exit_quit_done: bool = False
        self._exit_progress: Optional[QProgressDialog] = None

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

            # ★ file を消すなら Excel も閉じる（エンジン残留対策）
            if isinstance(tag, NodeTag) and tag.kind == "file":
                try:
                    logger.info("[Hover] request close_book before remove path=%s", tag.path)
                    self._engine_exec("close_book", path=tag.path)
                except Exception as e:
                    logger.error("[Hover] close_book failed: %s", e, exc_info=True)

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
                    ctx = {}
                    try:
                        ctx = self._excel.get_active_context() or {}
                    except Exception as e:
                        logger.error("[MACRO] get_active_context failed: %s", e)

                    record_kwargs = dict(kwargs)
                    if "workbook" not in record_kwargs:
                        record_kwargs["workbook"] = ctx.get("workbook", "")
                    if "sheet" not in record_kwargs:
                        record_kwargs["sheet"] = ctx.get("sheet", "")

                    self._macro.record(op, **record_kwargs)

            except Exception as e:
                logger.error("macro record failed: %s", e, exc_info=True)

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

        menu.addAction("Add Group...", self.add_group_dialog)

        menu.addSeparator()
        menu.addAction("Add Files...", self.add_files_dialog)
        menu.addAction("Add Folder...", self.add_folder_dialog)

        menu.addSeparator()
        menu.addAction("Inspector (Record Mode)", self._open_inspector)
        menu.addSeparator()
        menu.addAction("▶ Run Macro...", self.macro_play_dialog)
        menu.addAction("⏹ Stop Macro", self.macro_stop_play)

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
            pass

    def _on_selection_changed(self, *_):
        self._execute_current_selection()

    def _on_sheets_ready(self, path: str, sheets):
        item = self._find_file_item(path)
        if not item:
            return

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
        root = self._create_item(group_name, NodeTag("folder", ""))
        self._model.appendRow(root)

        for name in sorted(os.listdir(folder)):
            full = os.path.join(folder, name)
            if _is_openable_excel_path(full):
                self._add_file(full, root)

        self.expand(root.index())
        logger.info("[GROUP] folder imported folder=%s -> group=%s", folder, group_name)

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
    # Exit progress callbacks
    # =================================================
    def _on_exit_book_closed(self, path: str):
        if not self._is_shutting_down:
            return
        ap = _abspath(path)
        if ap in self._exit_closed or ap in self._exit_failed:
            return
        self._exit_closed.add(ap)
        self._exit_done += 1
        logger.info("[EXIT] book_closed %s (%s/%s)", ap, self._exit_done, self._exit_total)
        self._update_exit_progress()

    def _on_exit_book_close_failed(self, path: str, msg: str):
        if not self._is_shutting_down:
            return
        ap = _abspath(path)
        if ap in self._exit_closed or ap in self._exit_failed:
            return
        self._exit_failed.add(ap)
        self._exit_done += 1
        logger.error("[EXIT] book_close_failed %s err=%s (%s/%s)", ap, msg, self._exit_done, self._exit_total)
        self._update_exit_progress()

    def _on_exit_quit_finished(self):
        if not self._is_shutting_down:
            return
        self._exit_quit_done = True
        logger.info("[EXIT] quit_finished received")

    def _update_exit_progress(self):
        if not self._exit_progress:
            return
        try:
            self._exit_progress.setValue(self._exit_done)
            self._exit_progress.setLabelText(f"Closing Excel books... {self._exit_done}/{self._exit_total}")
        except Exception:
            pass

    # =================================================
    # Shutdown Excel (public)
    # =================================================
    def _collect_tree_book_paths(self) -> List[str]:
        out: List[str] = []
        root = self._model.invisibleRootItem()
        stack = [root.child(r) for r in range(root.rowCount())]
        while stack:
            it = stack.pop()
            tag = it.data(ROLE_TAG)
            if isinstance(tag, NodeTag) and tag.kind == "file":
                out.append(_abspath(tag.path))
            for r in range(it.rowCount()):
                stack.append(it.child(r))

        seen = set()
        uniq: List[str] = []
        for p in out:
            if p and p not in seen:
                seen.add(p)
                uniq.append(p)
        return uniq

    def shutdown_excel_on_exit(self):
        """
        ✅ 要求どおり：
        終了時は「ツリー上のブック」を先に閉じて、その後 ExcelWorker を shutdown してエンジンを残さない。
        """
        if self._is_shutting_down:
            logger.info("[EXIT] shutdown already in progress -> skip")
            return

        self._is_shutting_down = True
        logger.info("[EXIT] shutdown_excel_on_exit begin")

        try:
            if self._inspector:
                self._inspector.close()
        except Exception as e:
            logger.error("[EXIT] inspector close failed: %s", e, exc_info=True)

        paths = self._collect_tree_book_paths()
        self._exit_total = len(paths)
        self._exit_done = 0
        self._exit_closed = set()
        self._exit_failed = set()
        self._exit_quit_done = False

        if self._exit_total > 0:
            self._exit_progress = QProgressDialog(
                "Closing Excel books...",
                None,
                0,
                self._exit_total,
                self,
            )
            self._exit_progress.setWindowTitle("Closing")
            self._exit_progress.setWindowModality(Qt.ApplicationModal)
            self._exit_progress.setCancelButton(None)
            self._exit_progress.setMinimumDuration(0)
            self._exit_progress.show()
            self._update_exit_progress()

        # 1) ツリー上のブックを全部 close 要求
        for p in paths:
            try:
                logger.info("[EXIT] request_close path=%s", p)
                self._excel.request_close(p)
            except Exception as e:
                logger.error("[EXIT] request_close failed path=%s err=%s", p, e, exc_info=True)
                self._exit_done += 1
                self._exit_failed.add(p)
                self._update_exit_progress()

        # 2) close 完了を待つ（Signal で進捗が進む）
        start = time.time()
        timeout_sec = 10.0
        while self._exit_done < self._exit_total:
            QApplication.processEvents()
            if time.time() - start > timeout_sec:
                logger.error("[EXIT] close wait timeout done=%s total=%s", self._exit_done, self._exit_total)
                break
            time.sleep(0.02)

        # 3) 最後に ExcelWorker を shutdown（Quit + COM破棄）
        try:
            logger.info("[EXIT] ExcelWorker.shutdown(confirm_save=True)")
            self._excel.shutdown(confirm_save=True)
        except Exception as e:
            logger.error("[EXIT] ExcelWorker.shutdown failed: %s", e, exc_info=True)

        # 4) thread 終了待ち（エンジン残留防止）
        wait_start = time.time()
        wait_timeout_sec = 10.0
        while True:
            QApplication.processEvents()
            try:
                if not self._excel.isRunning():
                    break
            except Exception:
                break

            if time.time() - wait_start > wait_timeout_sec:
                logger.error("[EXIT] ExcelWorker wait timeout")
                break
            time.sleep(0.02)

        if self._exit_progress:
            try:
                self._exit_progress.close()
            except Exception:
                pass
            self._exit_progress = None

        logger.info("[EXIT] shutdown_excel_on_exit done closed=%s failed=%s", len(self._exit_closed), len(self._exit_failed))
        self._is_shutting_down = False

    # =================================================
    # Macro play
    # =================================================
    def macro_play_dialog(self):
        logger.info(
            "[MACRO] play requested running=%s",
            bool(self._macro_play_thread and self._macro_play_thread.isRunning()),
        )
        if self._macro_play_thread and self._macro_play_thread.isRunning():
            QMessageBox.information(self, "Macro", "マクロは既に実行中です")
            return

        path, _ = QFileDialog.getOpenFileName(
            self, "Run Macro", "", "Macro JSON (*.json)"
        )
        if not path:
            return

        try:
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
        except Exception as e:
            QMessageBox.critical(self, "Macro", f"読み込み失敗:\n{e}")
            return

        self._macro_play_thread = _MacroPlayThread(self, data)
        self._macro_play_thread.start()

        logger.info("[MACRO] play start path=%s", path)

    def macro_stop_play(self):
        logger.info(
            "[MACRO] stop button pressed thread_exists=%s running=%s",
            bool(self._macro_play_thread),
            bool(self._macro_play_thread and self._macro_play_thread.isRunning()),
        )
        if self._macro_play_thread:
            self._macro_play_thread.stop()
            logger.info("[MACRO] stop requested")

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


class _MacroPlayThread(QThread):
    def __init__(self, tree: "LauncherTreeView", macro: Dict[str, Any]):
        super().__init__()
        self._tree = tree
        self._macro = macro
        self._stop = False

        self._log = get_logger("MacroPlayThread")
        self._log.info(
            "[INIT] thread created isRunning=%s",
            self.isRunning(),
        )

        self.finished.connect(self._on_finished)

    def stop(self):
        self._stop = True
        self._log.info(
            "[STOP] stop requested isRunning=%s",
            self.isRunning(),
        )

    def run(self):
        self._log.info("[RUN] start")

        steps = self._macro.get("steps", [])
        self._log.info("[RUN] steps count=%s", len(steps))

        for i, step in enumerate(steps):
            if self._stop:
                self._log.warning(
                    "[RUN] stop flag detected at step=%s",
                    i,
                )
                break

            op = step.get("op")
            args = step.get("args", {})

            self._log.info(
                "[STEP] idx=%s op=%s args=%s",
                i,
                op,
                args,
            )

            try:
                self._tree._engine_exec(op, source="macro", **args)
                self._log.info(
                    "[STEP] idx=%s op=%s DONE",
                    i,
                    op,
                )
            except Exception as e:
                self._log.error(
                    "[STEP] idx=%s FAILED op=%s err=%s",
                    i,
                    op,
                    e,
                    exc_info=True,
                )
                break

        self._log.info("[RUN] end stopped=%s", self._stop)

    def _on_finished(self):
        self._log.info(
            "[FINISHED] thread finished isRunning=%s",
            self.isRunning(),
        )
