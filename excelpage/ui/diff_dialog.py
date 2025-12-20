from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout,
    QCheckBox, QLabel, QPushButton
)


class DiffOptionDialog(QDialog):
    def __init__(
        self,
        parent=None,
        *,
        key_cols=None,
        compare_formula=False,
        include_context=True,
        compare_shapes=False,
    ):
        super().__init__(parent)
        self.setWindowTitle("Diff オプション")

        self._key_cols = key_cols or []

        layout = QVBoxLayout(self)

        info = QLabel("Diff オプションを指定してください")
        layout.addWidget(info)

        self.chk_formula = QCheckBox("数式も比較")
        self.chk_formula.setChecked(compare_formula)
        layout.addWidget(self.chk_formula)

        self.chk_context = QCheckBox("ジャンプリンク / コンテキストを含める")
        self.chk_context.setChecked(include_context)
        layout.addWidget(self.chk_context)

        self.chk_shapes = QCheckBox("図・画像も比較")
        self.chk_shapes.setChecked(compare_shapes)
        layout.addWidget(self.chk_shapes)

        btns = QHBoxLayout()
        btns.addStretch()

        ok = QPushButton("OK")
        cancel = QPushButton("キャンセル")
        ok.clicked.connect(self.accept)
        cancel.clicked.connect(self.reject)

        btns.addWidget(ok)
        btns.addWidget(cancel)
        layout.addLayout(btns)

    def get_options(self) -> dict:
        return {
            "key_cols": list(self._key_cols),
            "compare_formula": self.chk_formula.isChecked(),
            "include_context": self.chk_context.isChecked(),
            "compare_shapes": self.chk_shapes.isChecked(),
        }
