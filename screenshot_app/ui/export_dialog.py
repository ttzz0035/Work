from __future__ import annotations
from pathlib import Path
from typing import Optional
from PySide6 import QtCore, QtGui, QtWidgets

class ExportDialog(QtWidgets.QDialog):
    def __init__(self, parent=None, default_dir: Optional[Path]=None):
        super().__init__(parent)
        self.setWindowTitle("Export")
        self.setModal(True)

        self._default_dir = Path(default_dir) if default_dir else Path.cwd()

        form = QtWidgets.QFormLayout(self)

        self.edTitle = QtWidgets.QLineEdit()
        self.cmbFormat = QtWidgets.QComboBox()
        self.cmbFormat.addItems(["Excel (.xlsx)", "HTML (.html)"])

        self.edPath = QtWidgets.QLineEdit()
        self.btnBrowse = QtWidgets.QToolButton(text="…")
        self.btnBrowse.clicked.connect(self._browse)

        pathLay = QtWidgets.QHBoxLayout()
        pathLay.addWidget(self.edPath, 1)
        pathLay.addWidget(self.btnBrowse, 0)

        form.addRow("Title (optional):", self.edTitle)
        form.addRow("Format:", self.cmbFormat)
        form.addRow("Output file:", pathLay)

        self.cmbFormat.currentIndexChanged.connect(self._update_suggested_path)
        self._update_suggested_path()

        btns = QtWidgets.QDialogButtonBox(QtWidgets.QDialogButtonBox.Ok | QtWidgets.QDialogButtonBox.Cancel)
        btns.accepted.connect(self.accept)
        btns.rejected.connect(self.reject)
        form.addRow(btns)

        self.resize(520, 160)

    def _update_suggested_path(self):
        idx = self.cmbFormat.currentIndex()
        ext = ".xlsx" if idx == 0 else ".html"
        cur = Path(self.edPath.text().strip() or "")
        if not cur.suffix:
            self.edPath.setText(str(self._default_dir / f"captures_export{ext}"))
        else:
            self.edPath.setText(str(cur.with_suffix(ext)))

    def _browse(self):
        idx = self.cmbFormat.currentIndex()
        if idx == 0:
            fn, _ = QtWidgets.QFileDialog.getSaveFileName(self, "Save Excel", str(self._default_dir), "Excel (*.xlsx)")
        else:
            fn, _ = QtWidgets.QFileDialog.getSaveFileName(self, "Save HTML", str(self._default_dir), "HTML (*.html)")
        if fn:
            self.edPath.setText(fn)

    # 結果取得
    def result_title(self) -> str:
        return self.edTitle.text().strip()

    def result_format(self) -> str:
        return "excel" if self.cmbFormat.currentIndex() == 0 else "html"

    def result_path(self) -> Path:
        return Path(self.edPath.text().strip())
