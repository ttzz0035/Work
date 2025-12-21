# main.py
from __future__ import annotations

import sys
import json
import os
from typing import Dict, Any

from PySide6.QtWidgets import (
    QApplication,
    QMainWindow,
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QLabel,
    QToolButton,
    QMenu,
    QProgressDialog,
    QFileDialog,
    QMessageBox,
)
from PySide6.QtCore import Qt, QTimer

from ui.tree_view import LauncherTreeView
from logger import get_logger

logger = get_logger("App")


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Excel Workspace Launcher")
        self.resize(520, 860)

        # =========================
        # UI
        # =========================
        root = QWidget()
        self.setCentralWidget(root)

        layout = QVBoxLayout(root)
        layout.setContentsMargins(12, 10, 12, 12)

        # -------------------------
        # Header
        # -------------------------
        header = QWidget()
        header_layout = QHBoxLayout(header)
        header_layout.setContentsMargins(12, 10, 12, 10)

        title = QLabel("Workspace")
        title.setStyleSheet("font-size: 16px; font-weight: 600;")
        header_layout.addWidget(title)
        header_layout.addStretch(1)

        self.btn_plus = QToolButton()
        self.btn_plus.setText("+")
        self.btn_plus.setPopupMode(QToolButton.InstantPopup)

        plus_menu = QMenu(self.btn_plus)
        act_add_files = plus_menu.addAction("Add Files...")
        act_add_folder = plus_menu.addAction("Add Folder...")
        act_load_project = plus_menu.addAction("Load Project...")
        self.btn_plus.setMenu(plus_menu)

        header_layout.addWidget(self.btn_plus)

        layout.addWidget(header)

        # -------------------------
        # Tree
        # -------------------------
        self.tree = LauncherTreeView(self)
        layout.addWidget(self.tree, 1)

        # -------------------------
        # Wire
        # -------------------------
        act_add_files.triggered.connect(self.tree.add_files_dialog)
        act_add_folder.triggered.connect(self.tree.add_folder_dialog)
        act_load_project.triggered.connect(self.load_project_dialog)

        # -------------------------
        # 起動時プログレス
        # -------------------------
        self._startup_progress: QProgressDialog | None = None
        self._startup_total: int = 0
        self._startup_done: int = 0
        self._startup_books: set[str] = set()

        # Excel 起動進捗を監視
        excel = self.tree._excel
        if hasattr(excel, "book_closed"):
            excel.book_closed.connect(self._on_startup_book_event)
        if hasattr(excel, "book_close_failed"):
            excel.book_close_failed.connect(self._on_startup_book_event)

        logger.info("MainWindow initialized")

    # =================================================
    # Startup progress
    # =================================================
    def show_startup_progress(self, total_books: int):
        logger.info("[Startup] progress start total=%s", total_books)

        self._startup_total = total_books
        self._startup_done = 0
        self._startup_books.clear()

        self._startup_progress = QProgressDialog(
            "Starting Excel workspace...",
            None,
            0,
            max(1, total_books),
            self,
        )
        self._startup_progress.setWindowTitle("Starting")
        self._startup_progress.setWindowModality(Qt.ApplicationModal)
        self._startup_progress.setCancelButton(None)
        self._startup_progress.setMinimumDuration(0)
        self._startup_progress.show()

    def _on_startup_book_event(self, path: str, *_):
        """
        ExcelWorker からの通知で進捗を進める
        """
        if not self._startup_progress:
            return

        ap = os.path.abspath(path)
        if ap in self._startup_books:
            return

        self._startup_books.add(ap)
        self._startup_done += 1

        try:
            self._startup_progress.setValue(self._startup_done)
            self._startup_progress.setLabelText(
                f"Starting Excel workspace... {self._startup_done}/{self._startup_total}"
            )
        except Exception:
            pass

        logger.info(
            "[Startup] book ready %s (%s/%s)",
            ap,
            self._startup_done,
            self._startup_total,
        )

        if self._startup_done >= self._startup_total:
            try:
                self._startup_progress.close()
            except Exception:
                pass
            self._startup_progress = None
            logger.info("[Startup] progress finished")

    # =================================================
    # Project load
    # =================================================
    def load_project_dialog(self):
        path, _ = QFileDialog.getOpenFileName(
            self,
            "Load Project",
            "",
            "Project JSON (*.json)",
        )
        if not path:
            return

        try:
            with open(path, "r", encoding="utf-8") as f:
                data: Dict[str, Any] = json.load(f)
        except Exception as e:
            QMessageBox.critical(self, "Project", f"読み込み失敗:\n{e}")
            return

        # 起動対象ブック数を事前に数える
        total_books = self._count_books(data)
        self.show_startup_progress(total_books)

        try:
            self.tree.import_project(data)
        except Exception as e:
            logger.error("[Project] import failed: %s", e, exc_info=True)
            QMessageBox.critical(self, "Project", f"ロード失敗:\n{e}")
            if self._startup_progress:
                self._startup_progress.close()
                self._startup_progress = None

    def _count_books(self, node: Dict[str, Any]) -> int:
        cnt = 0
        if not isinstance(node, dict):
            return 0

        tag = node.get("tag", {})
        if isinstance(tag, dict) and tag.get("kind") == "file":
            cnt += 1

        for ch in node.get("children", []):
            cnt += self._count_books(ch)

        return cnt

    # =================================================
    # Close
    # =================================================
    def closeEvent(self, event):
        logger.info("[MainWindow] closeEvent")
        try:
            self.tree.shutdown_excel_on_exit()
        except Exception as e:
            logger.error("[MainWindow] shutdown failed: %s", e, exc_info=True)
        super().closeEvent(event)


def main():
    app = QApplication(sys.argv)

    # -------------------------
    # Style
    # -------------------------
    app.setStyleSheet(
        """
        QMainWindow { background: #111; }
        QLabel { color: #ddd; }
        QTreeView {
            background: #151515;
            color: #ddd;
            border: 1px solid #2a2a2a;
            border-radius: 12px;
            padding: 6px;
        }
        QTreeView::item {
            padding: 8px 8px;
            border-radius: 10px;
        }
        QTreeView::item:selected {
            background: #2d5fff;
            color: white;
        }
        QToolButton {
            background: #1f1f1f;
            color: #ddd;
            border: 1px solid #2a2a2a;
            border-radius: 12px;
            padding: 6px 12px;
            font-size: 14px;
            font-weight: 700;
        }
        QToolButton:hover { background: #252525; }
        QMenu {
            background: #1b1b1b;
            color: #ddd;
            border: 1px solid #2a2a2a;
        }
        QMenu::item:selected { background: #2d5fff; }
        QMessageBox { background: #1b1b1b; color: #ddd; }
        """
    )

    win = MainWindow()
    win.show()

    # GC 防止
    app._main_window = win

    sys.exit(app.exec())


if __name__ == "__main__":
    main()
