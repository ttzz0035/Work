from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLineEdit, QLabel
)
from logger import get_logger

logger = get_logger("InspectorPanel")


class InspectorPanel(QWidget):
    """
    録画モード専用 操作盤
    ※ すべて engine.exec 経由（Excel直接操作なし）
    """

    def __init__(self, engine, parent=None):
        super().__init__(parent)
        self._engine = engine

        self.setWindowTitle("Inspector (Record Mode)")
        self.setMinimumWidth(260)

        root = QVBoxLayout(self)

        # --- Cell select ---
        row = QHBoxLayout()
        row.addWidget(QLabel("Cell"))
        self.cell_edit = QLineEdit("A1")
        row.addWidget(self.cell_edit)
        btn = QPushButton("Select")
        btn.clicked.connect(self._on_select)
        row.addWidget(btn)
        root.addLayout(row)

        # --- Value ---
        row = QHBoxLayout()
        row.addWidget(QLabel("Value"))
        self.value_edit = QLineEdit("")
        row.addWidget(self.value_edit)
        btn = QPushButton("Set")
        btn.clicked.connect(self._on_set)
        row.addWidget(btn)
        root.addLayout(row)

        # --- Move ---
        move = QHBoxLayout()
        for txt, d in [("↑", "up"), ("↓", "down"), ("←", "left"), ("→", "right")]:
            b = QPushButton(txt)
            b.clicked.connect(lambda _, dd=d: self._engine.exec(
                "move_cell", direction=dd, step=1
            ))
            move.addWidget(b)
        root.addLayout(move)

        # --- Copy / Paste ---
        row = QHBoxLayout()
        c = QPushButton("Copy")
        c.clicked.connect(lambda: self._engine.exec("copy_selection"))
        p = QPushButton("Paste")
        p.clicked.connect(lambda: self._engine.exec("paste_selection"))
        row.addWidget(c)
        row.addWidget(p)
        root.addLayout(row)

        root.addStretch()

        logger.info("InspectorPanel ready")

    def _on_select(self):
        cell = self.cell_edit.text().strip()
        if cell:
            self._engine.exec("select_cell", cell=cell)

    def _on_set(self):
        cell = self.cell_edit.text().strip()
        val = self.value_edit.text()
        if cell:
            self._engine.exec("set_cell_value", cell=cell, value=val)
