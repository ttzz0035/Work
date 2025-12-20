# excel_transfer/ui/diff_dialog.py
from __future__ import annotations

import os
from typing import Any, Dict, List, Optional, Tuple

import xlwings as xw
from PySide6.QtWidgets import (
    QDialog,
    QVBoxLayout,
    QHBoxLayout,
    QLineEdit,
    QComboBox,
    QLabel,
    QTableView,
    QMenu,
    QMessageBox,
    QPushButton,
)
from PySide6.QtGui import QStandardItemModel, QStandardItem
from PySide6.QtCore import Qt, QPoint, QModelIndex

from logger import get_logger

logger = get_logger("DiffDialog")


# =====================================================
# Excel Jump (xlwings)
# =====================================================
def _get_or_create_app() -> xw.App:
    try:
        if xw.apps and len(xw.apps) > 0:
            app = xw.apps.active
            if app is not None:
                return app
    except Exception as ex:
        logger.warning("[JUMP] reuse active app failed: %s", ex)

    return xw.App(visible=True, add_book=False)


def _find_open_book(app: xw.App, file_path: str) -> Optional[xw.Book]:
    target = os.path.abspath(file_path)
    try:
        for b in app.books:
            try:
                if os.path.abspath(b.fullname) == target:
                    return b
            except Exception:
                continue
    except Exception as ex:
        logger.warning("[JUMP] iterate books failed: %s", ex)
    return None


def jump_to_cell(file_path: str, sheet_name: str, row: int, col: int) -> None:
    logger.info("[JUMP] file=%s sheet=%s r=%s c=%s", file_path, sheet_name, row, col)

    if not file_path or not os.path.exists(file_path):
        raise ValueError(f"file not found: {file_path}")
    if row is None or col is None:
        raise ValueError("invalid cell position")

    app = _get_or_create_app()
    book = _find_open_book(app, file_path)
    if book is None:
        book = app.books.open(file_path, read_only=False)

    sht = book.sheets[sheet_name] if sheet_name else book.sheets[0]

    try:
        try:
            book.app.activate(steal_focus=True)
        except Exception:
            book.app.activate()

        sht.activate()
        sht.range((int(row), int(col))).select()
    except Exception as ex:
        logger.error("[JUMP] failed: %s", ex)
        raise


# =====================================================
# Utility（ネスト関数禁止対応：全部トップレベル）
# =====================================================
def _str(v: Any) -> str:
    try:
        return "" if v is None else str(v)
    except Exception:
        return "<ERR>"


def _key_str(k: Any) -> str:
    try:
        return str(k)
    except Exception:
        return "<KEY>"


def _make_item(text: str) -> QStandardItem:
    it = QStandardItem(text)
    it.setEditable(False)
    return it


# =====================================================
# Diff Dialog
# =====================================================
class DiffDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)

        self.setWindowTitle("DIFF Results")
        self.resize(1200, 700)

        self._all_items: List[Dict[str, Any]] = []
        self._meta: Dict[str, Any] = {}

        self._model = QStandardItemModel()
        self._model.setHorizontalHeaderLabels(
            ["kind", "key", "column", "A", "B", "A(r,c)", "B(r,c)", "fileA", "sheetA", "fileB", "sheetB"]
        )

        self._table = QTableView()
        self._table.setModel(self._model)
        self._table.setSelectionBehavior(QTableView.SelectRows)
        self._table.setSelectionMode(QTableView.SingleSelection)
        self._table.setSortingEnabled(True)
        self._table.setContextMenuPolicy(Qt.CustomContextMenu)
        self._table.customContextMenuRequested.connect(self._on_context_menu)
        self._table.doubleClicked.connect(self._on_double_clicked)

        self._filter_kind = QComboBox()
        self._filter_kind.addItems(["ALL", "MOD", "ADD", "DEL"])
        self._filter_kind.currentTextChanged.connect(self._apply_filter)

        self._search = QLineEdit()
        self._search.setPlaceholderText("Search key/col/value ...")
        self._search.textChanged.connect(self._apply_filter)

        self._status = QLabel("")

        btn_close = QPushButton("Close")
        btn_close.clicked.connect(self.accept)

        top = QHBoxLayout()
        top.addWidget(QLabel("Filter:"))
        top.addWidget(self._filter_kind)
        top.addWidget(QLabel("Search:"))
        top.addWidget(self._search, 1)
        top.addWidget(btn_close)

        lay = QVBoxLayout()
        lay.addLayout(top)
        lay.addWidget(self._status)
        lay.addWidget(self._table, 1)
        self.setLayout(lay)

    # -----------------------------
    # Public
    # -----------------------------
    def set_results(self, items: List[Dict[str, Any]], meta: Dict[str, Any]) -> None:
        self._all_items = items or []
        self._meta = meta or {}
        self._apply_filter()

    # -----------------------------
    # Filtering
    # -----------------------------
    def _apply_filter(self) -> None:
        kind = self._filter_kind.currentText() or "ALL"
        q = (self._search.text() or "").strip().lower()

        self._model.removeRows(0, self._model.rowCount())

        shown = 0
        for it in self._all_items:
            if kind != "ALL" and it.get("kind") != kind:
                continue

            blob = " ".join(
                [
                    _str(it.get("kind")),
                    _key_str(it.get("key")),
                    _str(it.get("column")),
                    _str(it.get("a_value")),
                    _str(it.get("b_value")),
                ]
            ).lower()

            if q and q not in blob:
                continue

            a_rc = ""
            if it.get("a_row") is not None and it.get("a_col") is not None:
                a_rc = f"({it.get('a_row')},{it.get('a_col')})"
            b_rc = ""
            if it.get("b_row") is not None and it.get("b_col") is not None:
                b_rc = f"({it.get('b_row')},{it.get('b_col')})"

            row_items = [
                _make_item(_str(it.get("kind"))),
                _make_item(_key_str(it.get("key"))),
                _make_item(_str(it.get("column"))),
                _make_item(_str(it.get("a_value"))),
                _make_item(_str(it.get("b_value"))),
                _make_item(a_rc),
                _make_item(b_rc),
                _make_item(_str(it.get("file_a"))),
                _make_item(_str(it.get("sheet_a"))),
                _make_item(_str(it.get("file_b"))),
                _make_item(_str(it.get("sheet_b"))),
            ]

            # 元データを UserRole に埋める（ジャンプ用）
            for qi in row_items:
                qi.setData(it, Qt.ItemDataRole.UserRole + 10)

            self._model.appendRow(row_items)
            shown += 1

        self._status.setText(f"shown={shown} / total={len(self._all_items)}    {self._meta.get('sheet_a','')} ↔ {self._meta.get('sheet_b','')}")

        # file/sheet列は邪魔なら隠す（ただしデバッグ用に保持）
        self._table.setColumnHidden(7, True)
        self._table.setColumnHidden(8, True)
        self._table.setColumnHidden(9, True)
        self._table.setColumnHidden(10, True)

        self._table.resizeColumnsToContents()

    # -----------------------------
    # Row Data
    # -----------------------------
    def _current_item(self, idx: QModelIndex) -> Optional[Dict[str, Any]]:
        if not idx.isValid():
            return None
        it = idx.data(Qt.ItemDataRole.UserRole + 10)
        if isinstance(it, dict):
            return it
        return None

    # -----------------------------
    # Actions
    # -----------------------------
    def _jump_a(self, it: Dict[str, Any]) -> None:
        try:
            if it.get("a_row") is None or it.get("a_col") is None:
                QMessageBox.information(self, "Jump", "A側は座標が無いのでジャンプできません。")
                return
            jump_to_cell(it["file_a"], it["sheet_a"], int(it["a_row"]), int(it["a_col"]))
        except Exception as ex:
            QMessageBox.critical(self, "Jump", f"A側ジャンプ失敗:\n{ex}")

    def _jump_b(self, it: Dict[str, Any]) -> None:
        try:
            if it.get("b_row") is None or it.get("b_col") is None:
                QMessageBox.information(self, "Jump", "B側は座標が無いのでジャンプできません。")
                return
            jump_to_cell(it["file_b"], it["sheet_b"], int(it["b_row"]), int(it["b_col"]))
        except Exception as ex:
            QMessageBox.critical(self, "Jump", f"B側ジャンプ失敗:\n{ex}")

    # -----------------------------
    # UI events
    # -----------------------------
    def _on_double_clicked(self, idx: QModelIndex) -> None:
        it = self._current_item(idx)
        if not it:
            return
        # ダブルクリックはA側ジャンプ
        self._jump_a(it)

    def _on_context_menu(self, pos: QPoint) -> None:
        idx = self._table.indexAt(pos)
        it = self._current_item(idx)
        if not it:
            return

        menu = QMenu(self)
        menu.addAction("Jump to A", lambda: self._jump_a(it))
        menu.addAction("Jump to B", lambda: self._jump_b(it))
        menu.exec(self._table.viewport().mapToGlobal(pos))
