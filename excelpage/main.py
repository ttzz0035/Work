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
    QFileDialog,
    QMessageBox,
)
from PySide6.QtCore import Qt, QTimer

from ui.tree_view import LauncherTreeView
from services.project_io import save_project, load_project
from logger import get_logger

logger = get_logger("App")


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Excel Workspace Launcher")
        self.resize(520, 860)

        self._project_path: str | None = None

        root = QWidget()
        self.setCentralWidget(root)

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
        self.btn_plus.setMenu(plus_menu)
        header_layout.addWidget(self.btn_plus)

        self.btn_menu = QToolButton()
        self.btn_menu.setText("≡")
        self.btn_menu.setPopupMode(QToolButton.InstantPopup)

        app_menu = QMenu(self.btn_menu)
        act_load_project = app_menu.addAction("Load Project...")
        act_save_project = app_menu.addAction("Save Project...")
        self.btn_menu.setMenu(app_menu)
        header_layout.addWidget(self.btn_menu)

        self.tree = LauncherTreeView()

        layout = QVBoxLayout(root)
        layout.setContentsMargins(12, 10, 12, 12)
        layout.addWidget(header)
        layout.addWidget(self.tree, 1)

        act_add_files.triggered.connect(self.tree.add_files_dialog)
        act_add_folder.triggered.connect(self.tree.add_folder_dialog)
        act_save_project.triggered.connect(self.save_project_dialog)
        act_load_project.triggered.connect(self.load_project_dialog)

        logger.info("MainWindow initialized")

    def save_project_dialog(self):
        path, _ = QFileDialog.getSaveFileName(
            self,
            "Save Project",
            self._project_path or "",
            "Project (*.json)",
        )
        if not path:
            logger.info("[Project] save canceled")
            return

        try:
            save_project(path, self.tree)
            self._project_path = path
            logger.info("[Project] saved path=%s", path)
        except Exception as e:
            logger.exception("[Project] save failed: %s", e)
            QMessageBox.critical(self, "Save Project", f"保存に失敗しました:\n{e}")

    def load_project_dialog(self):
        path, _ = QFileDialog.getOpenFileName(
            self,
            "Load Project",
            "",
            "Project (*.json)",
        )
        if not path:
            logger.info("[Project] load canceled")
            return

        try:
            load_project(path, self.tree)
            self._project_path = path
            logger.info("[Project] loaded path=%s", path)
        except Exception as e:
            logger.exception("[Project] load failed: %s", e)
            QMessageBox.critical(self, "Load Project", f"読み込みに失敗しました:\n{e}")

    def closeEvent(self, event):
        logger.info("[MainWindow] closeEvent begin")

        ret = QMessageBox.question(
            self,
            "終了確認",
            "プロジェクトを保存しますか？",
            QMessageBox.Yes | QMessageBox.No | QMessageBox.Cancel,
            QMessageBox.Yes,
        )

        if ret == QMessageBox.Cancel:
            logger.info("[MainWindow] close canceled")
            event.ignore()
            return

        if ret == QMessageBox.Yes:
            self.save_project_dialog()

        try:
            if self.tree:
                self.tree.shutdown_excel_on_exit()
        except Exception as e:
            logger.exception("[MainWindow] shutdown excel failed: %s", e)

        logger.info("[MainWindow] closeEvent -> super")
        super().closeEvent(event)


def show_main_window(app: QApplication, progress: QProgressDialog):
    try:
        progress.close()
        logger.info("Startup progress closed")
    except Exception as e:
        logger.error("Startup progress close failed: %s", e, exc_info=True)

    win = MainWindow()
    win.show()
    app._main_window = win
    logger.info("Main window shown")


def main():
    app = QApplication(sys.argv)

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

    QTimer.singleShot(100, lambda: show_main_window(app, progress))

    sys.exit(app.exec())


if __name__ == "__main__":
    main()
