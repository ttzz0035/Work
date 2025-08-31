from __future__ import annotations
from typing import Dict, Optional
from PySide6 import QtCore, QtGui, QtWidgets

# 設定に表示する項目（ホットキー名 ←→ ラベル）
LABELS: Dict[str, str] = {
    "capture":         "Capture",
    "add_rect":        "Add rectangle",
    "remove_selected": "Remove selected",
    "pick_new_color":  "Pick new color",
    "show_hotkeys":    "Show hotkeys",
    "exit_app":        "Exit app",
    # 録画系（追加）
    "rec_start":       "Record: Start",
    "rec_stop":        "Record: Stop",
    "rec_play":        "Record: Play last",
}

class SettingsDialog(QtWidgets.QDialog):
    def __init__(self, current_keys: Dict[str, str], parent=None):
        super().__init__(parent)
        self.setWindowTitle("Settings - Hotkeys")
        self.setWindowFlags(self.windowFlags() | QtCore.Qt.Tool)
        self.setModal(True)

        self._edits: Dict[str, QtWidgets.QKeySequenceEdit] = {}
        form = QtWidgets.QFormLayout()
        form.setLabelAlignment(QtCore.Qt.AlignRight)

        for key, label in LABELS.items():
            edit = QtWidgets.QKeySequenceEdit()
            edit.setKeySequence(QtGui.QKeySequence(current_keys.get(key, "")))
            self._edits[key] = edit
            form.addRow(label, edit)

        btns = QtWidgets.QDialogButtonBox(QtWidgets.QDialogButtonBox.Ok | QtWidgets.QDialogButtonBox.Cancel)
        btns.accepted.connect(self._on_ok)
        btns.rejected.connect(self.reject)

        lay = QtWidgets.QVBoxLayout(self)
        lay.addLayout(form)
        lay.addWidget(btns)

        self._result: Optional[Dict[str, str]] = None
        self.resize(480, 320)

    def _on_ok(self):
        new_map = {k: e.keySequence().toString() for k, e in self._edits.items()}
        used = [v for v in new_map.values() if v]
        dups = sorted({x for x in used if used.count(x) > 1})
        if dups:
            QtWidgets.QMessageBox.warning(self, "Conflict", f"Hotkey duplicated: {', '.join(dups)}")
            return
        self._result = new_map
        self.accept()

    def result_keys(self) -> Optional[Dict[str, str]]:
        return self._result