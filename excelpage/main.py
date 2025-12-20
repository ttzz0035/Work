import sys

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

        root = QWidget()
        self.setCentralWidget(root)

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

        # ---- + menu ----
        self.btn_plus = QToolButton()
        self.btn_plus.setText("+")
        self.btn_plus.setPopupMode(QToolButton.InstantPopup)

        plus_menu = QMenu(self.btn_plus)
        act_add_files = plus_menu.addAction("Add Files...")
        act_add_folder = plus_menu.addAction("Add Folder...")
        self.btn_plus.setMenu(plus_menu)

        header_layout.addWidget(self.btn_plus)

        # -------------------------
        # Tree
        # -------------------------
        self.tree = LauncherTreeView()

        layout = QVBoxLayout(root)
        layout.setContentsMargins(12, 10, 12, 12)
        layout.addWidget(header)
        layout.addWidget(self.tree, 1)

        # -------------------------
        # Wire
        # -------------------------
        act_add_files.triggered.connect(self.tree.add_files_dialog)
        act_add_folder.triggered.connect(self.tree.add_folder_dialog)

        logger.info("MainWindow initialized")


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

    # -------------------------
    # Startup progress
    # -------------------------
    progress = QProgressDialog(
        "Starting application...",
        None,
        0,
        0,
    )
    progress.setWindowTitle("Please wait")
    progress.setWindowModality(Qt.ApplicationModal)
    progress.setCancelButton(None)
    progress.setMinimumDuration(0)
    progress.show()

    logger.info("Startup progress shown")

    # -------------------------
    # Show main window
    # -------------------------
    def show_main():
        progress.close()
        logger.info("Startup progress closed")

        win = MainWindow()
        win.show()

        # GC防止
        app._main_window = win

    QTimer.singleShot(100, show_main)

    sys.exit(app.exec())


if __name__ == "__main__":
    main()
