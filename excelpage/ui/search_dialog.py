from __future__ import annotations

from typing import List, Dict, Any, Optional

from PySide6.QtWidgets import (
    QDialog,
    QVBoxLayout,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QCheckBox,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QMessageBox,
    QHeaderView,
)
from PySide6.QtCore import Qt

from Logger import Logger
logger = Logger(
    name="SearchDialog",
    log_file_path="logs/app.log",
    level="DEBUG",
)

class SearchDialog(QDialog):
    """
    検索UI（最小）
    - keyword
    - case sensitive
    結果はテーブルで表示
    """

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Search")
        self.resize(980, 620)

        self._keyword = QLineEdit()
        self._keyword.setPlaceholderText("keyword...")

        self._case = QCheckBox("Case sensitive")
        self._case.setChecked(False)

        self._btn_run = QPushButton("Search")
        self._btn_close = QPushButton("Close")

        self._table = QTableWidget(0, 5)
        self._table.setHorizontalHeaderLabels(["File", "Sheet", "Cell", "Value", "Match"])
        self._table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        self._table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        self._table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeToContents)
        self._table.horizontalHeader().setSectionResizeMode(3, QHeaderView.Stretch)
        self._table.horizontalHeader().setSectionResizeMode(4, QHeaderView.ResizeToContents)
        self._table.setEditTriggers(QTableWidget.NoEditTriggers)
        self._table.setSelectionBehavior(QTableWidget.SelectRows)

        root = QVBoxLayout(self)

        form = QHBoxLayout()
        form.addWidget(QLabel("Keyword:"))
        form.addWidget(self._keyword, 1)
        form.addWidget(self._case)
        form.addWidget(self._btn_run)
        form.addWidget(self._btn_close)
        root.addLayout(form)

        root.addWidget(self._table, 1)

        self._btn_close.clicked.connect(self.close)

        self._runner = None

        logger.info("SearchDialog initialized")

    def bind_runner(self, runner) -> None:
        """
        runner(keyword:str, case_sensitive:bool) -> List[dict]
        dict keys: file, sheet, cell, value, match
        """
        self._runner = runner
        self._btn_run.clicked.connect(self._on_run)

    def _on_run(self) -> None:
        if self._runner is None:
            QMessageBox.warning(self, "Search", "Runner not bound.")
            return
        keyword = self._keyword.text().strip()
        if not keyword:
            QMessageBox.information(self, "Search", "Please input keyword.")
            return
        case_sensitive = self._case.isChecked()

        try:
            logger.info("Search run: keyword=%s case=%s", keyword, case_sensitive)
            results = self._runner(keyword, case_sensitive)
            self._set_results(results)
        except Exception as ex:
            logger.error("Search failed: %s", ex, exc_info=True)
            QMessageBox.critical(self, "Search", f"Search failed: {ex}")

    def _set_results(self, rows: List[Dict[str, Any]]) -> None:
        self._table.setRowCount(0)
        for r in rows:
            row = self._table.rowCount()
            self._table.insertRow(row)
            self._table.setItem(row, 0, QTableWidgetItem(str(r.get("file", ""))))
            self._table.setItem(row, 1, QTableWidgetItem(str(r.get("sheet", ""))))
            self._table.setItem(row, 2, QTableWidgetItem(str(r.get("cell", ""))))
            self._table.setItem(row, 3, QTableWidgetItem(str(r.get("value", ""))))
            self._table.setItem(row, 4, QTableWidgetItem(str(r.get("match", ""))))
        logger.info("Search results: %d", len(rows))
