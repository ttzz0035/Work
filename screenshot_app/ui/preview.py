# ui/preview.py
from __future__ import annotations
from dataclasses import dataclass
from pathlib import Path
from typing import Optional, List
import json
import logging

from PySide6 import QtCore, QtGui, QtWidgets

from core.render import render_annotated  # (src_png:Path, meta:dict, folder:Path) -> Path

log = logging.getLogger("ui.preview")


# --- フォーカス外れ検知用 PlainTextEdit ---
class FocusSavePlainTextEdit(QtWidgets.QPlainTextEdit):
    focusLost = QtCore.Signal()

    def focusOutEvent(self, e: QtGui.QFocusEvent) -> None:
        try:
            self.focusLost.emit()
        finally:
            super().focusOutEvent(e)


@dataclass
class CardData:
    json_path: Path
    image_path: Path
    folder: Path
    comment: str
    display_title: str
    ann_png_path: Optional[Path] = None

    @property
    def title(self) -> str:
        return self.display_title

    def load_meta(self) -> dict:
        return json.loads(self.json_path.read_text(encoding="utf-8"))

    def save_meta(self, meta: dict):
        self.json_path.write_text(json.dumps(meta, ensure_ascii=False, indent=2), encoding="utf-8")

    def base_png_abs(self) -> Path:
        p = self.image_path
        return (self.folder / p) if not p.is_absolute() else p

    def regenerate_ann(self) -> Path:
        meta = self.load_meta()
        png_abs = self.base_png_abs()
        self.ann_png_path = render_annotated(png_abs, meta, self.folder)
        return self.ann_png_path

    def delete_files(self):
        # ann
        try:
            if self.ann_png_path and self.ann_png_path.exists():
                self.ann_png_path.unlink()
                log.info("annPng Delete: %s", self.ann_png_path.name)
        except Exception as e:
            log.warning("annPng Delete Failed: %s (%s)", self.ann_png_path, e)
        # base
        try:
            bp = self.base_png_abs()
            if bp.exists():
                bp.unlink()
                log.info("basePNG Delete: %s", bp.name)
        except Exception as e:
            log.warning("basePNG Delete Failed: %s", e)
        # json
        try:
            if self.json_path.exists():
                self.json_path.unlink()
                log.info("JSON Delete: %s", self.json_path.name)
        except Exception as e:
            log.warning("JSON Delete Failed: %s", e)


class CardWidget(QtWidgets.QFrame):
    requestRemove = QtCore.Signal(object)
    requestRefresh = QtCore.Signal(object)

    def __init__(self, data: CardData, parent=None):
        super().__init__(parent)
        self.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.setObjectName("Card")
        self.setStyleSheet("""
        QFrame#Card{border:1px solid #ddd;border-radius:8px;background:#fff;}
        QToolButton{background:#f7f7f7;border:1px solid #ccc;border-radius:6px;padding:4px 8px;}
        QToolButton:hover{background:#eee;}
        QLabel.head{font-weight:600;padding:6px 8px;border-bottom:1px solid #eee;background:#fafafa;border-top-left-radius:8px;border-top-right-radius:8px;}
        QLineEdit#titleEdit{padding:4px 6px;border:1px solid #ccc;border-radius:6px;min-width:260px;}
        """)

        self.data = data

        # Header
        self.head_lbl = QtWidgets.QLabel("Title:")
        self.head_lbl.setProperty("class", "head")
        self.title_edit = QtWidgets.QLineEdit(self.data.display_title)
        self.title_edit.setObjectName("titleEdit")
        self.title_edit.setToolTip("HTML/Excel Auto Save")

        # Buttons: OpenImage / Save / Delete
        self.btn_open_img = QtWidgets.QToolButton(self); self.btn_open_img.setText("OpenImage")
        self.btn_save     = QtWidgets.QToolButton(self); self.btn_save.setText("Save")
        self.btn_delete   = QtWidgets.QToolButton(self); self.btn_delete.setText("Delete")

        # Image
        self.image_lbl = QtWidgets.QLabel()
        self.image_lbl.setAlignment(QtCore.Qt.AlignCenter)
        self.image_lbl.setMinimumSize(320, 180)

        # Comment（フォーカス外れで保存）
        self.comment_edit = FocusSavePlainTextEdit(self.data.comment)
        self.comment_edit.setPlaceholderText("comment…")

        # Layout
        header = QtWidgets.QHBoxLayout()
        header.addWidget(self.head_lbl)
        header.addWidget(self.title_edit, 1)
        header.addStretch()
        header.addWidget(self.btn_open_img)
        header.addWidget(self.btn_save)
        header.addWidget(self.btn_delete)

        body = QtWidgets.QHBoxLayout()
        body.addWidget(self.image_lbl, 0)
        body.addWidget(self.comment_edit, 1)

        lay = QtWidgets.QVBoxLayout(self)
        lay.setContentsMargins(0,0,0,0)
        lay.addLayout(header)
        lay.addLayout(body)

        # Signals
        self.btn_open_img.clicked.connect(self._on_open_image)
        self.btn_save.clicked.connect(self._save_now)
        self.btn_delete.clicked.connect(self._on_delete)

        self.title_edit.editingFinished.connect(self._save_title_only)   # ← フォーカス外れ/Enterで保存
        self.comment_edit.focusLost.connect(self._save_comment_only)     # ← フォーカス外れで保存

        # 初期表示
        self.refresh_image()

    # --- UI 更新 ---
    def refresh_image(self):
        try:
            ann = self.data.regenerate_ann()
        except Exception as e:
            log.error("Ann Image Create Failed: %s", e)
            ann = None

        if ann and ann.exists():
            pm = QtGui.QPixmap(str(ann))
        else:
            pm = QtGui.QPixmap(self.data.base_png_abs().as_posix())

        if not pm.isNull():
            pm = pm.scaled(600, 600, QtCore.Qt.KeepAspectRatio, QtCore.Qt.SmoothTransformation)
        self.image_lbl.setPixmap(pm)

    # --- OpenImage ---
    def _on_open_image(self):
        png = self.data.base_png_abs()
        if not png.exists():
            QtWidgets.QMessageBox.information(self, "Open image", f"Not found: {png.name}")
            return
        QtGui.QDesktopServices.openUrl(QtCore.QUrl.fromLocalFile(str(png)))

    # --- Save(手動) ---
    def _save_now(self):
        title = self.title_edit.text().strip()
        if not title:
            QtWidgets.QMessageBox.information(self, "Save", "Title Empty")
            self.title_edit.setText(self.data.display_title)
            return
        try:
            meta = self.data.load_meta()
            meta["display_title"] = title
            meta["comment"] = self.comment_edit.toPlainText()
            self.data.display_title = title
            self.data.comment = meta["comment"]
            self.data.save_meta(meta)

            # 再描画（必要に応じ ann再生成）
            try:
                if self.data.ann_png_path and self.data.ann_png_path.exists():
                    self.data.ann_png_path.unlink()
                self.refresh_image()
            except Exception:
                pass

            self.requestRefresh.emit(self)
            log.info("Saved: title='%s', json=%s", title, self.data.json_path.name)
        except Exception as e:
            QtWidgets.QMessageBox.warning(self, "Save failed", f"{e}")

    # --- Title 自動保存（フォーカス外れ/Enter） ---
    def _save_title_only(self):
        title = self.title_edit.text().strip()
        if not title or title == self.data.display_title:
            return
        try:
            meta = self.data.load_meta()
            meta["display_title"] = title
            self.data.display_title = title
            self.data.save_meta(meta)
            self.requestRefresh.emit(self)
            log.info("AutoSaved (title): %s", title)
        except Exception as e:
            log.warning("AutoSave(title) Failed: %s", e)
            self.title_edit.setText(self.data.display_title)

    # --- Comment 自動保存（フォーカス外れ） ---
    def _save_comment_only(self):
        text = self.comment_edit.toPlainText()
        if text == self.data.comment:
            return
        try:
            meta = self.data.load_meta()
            meta["comment"] = text
            self.data.comment = text
            self.data.save_meta(meta)
            self.requestRefresh.emit(self)
            log.info("AutoSaved (comment)")
        except Exception as e:
            log.warning("AutoSave(comment) Failed: %s", e)

    # --- 削除 ---
    def _on_delete(self):
        self.data.delete_files()
        self.requestRemove.emit(self)


class PreviewPane(QtWidgets.QScrollArea):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWidgetResizable(True)

        self.container = QtWidgets.QWidget()
        self.vbox = QtWidgets.QVBoxLayout(self.container)
        self.vbox.setContentsMargins(8,8,8,8)
        self.vbox.setSpacing(12)
        self.vbox.addStretch(1)

        self.setWidget(self.container)
        self.cards: List[CardWidget] = []

    def _add_card_widget(self, cd: CardData):
        w = CardWidget(cd, self)
        w.requestRemove.connect(self._on_remove_card)
        w.requestRefresh.connect(self._on_refresh_card)
        self.vbox.insertWidget(self.vbox.count()-1, w)
        self.cards.append(w)

    def add_capture(self, json_path: Path):
        try:
            meta = json.loads(json_path.read_text(encoding="utf-8"))
        except Exception as e:
            log.warning("メタ読込失敗: %s (%s)", json_path, e)
            return
        folder = json_path.parent
        img_name = meta.get("image_path") or json_path.with_suffix(".png").name
        comment = meta.get("comment", "") or ""
        display_title = meta.get("display_title") or Path(img_name).stem  # ← 初回は PNG stem

        cd = CardData(json_path=json_path,
                      image_path=Path(img_name),
                      folder=folder,
                      comment=comment,
                      display_title=display_title)
        self._add_card_widget(cd)

    # signals
    def _on_remove_card(self, w: CardWidget):
        try:
            self.cards.remove(w)
        except ValueError:
            pass
        w.setParent(None); w.deleteLater()

    def _on_refresh_card(self, _w: CardWidget):
        pass

    def clear_all(self):
        for w in list(self.cards):
            w.setParent(None); w.deleteLater()
        self.cards.clear()
