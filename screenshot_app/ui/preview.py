# ui/preview.py
from __future__ import annotations
from dataclasses import dataclass
from pathlib import Path
from typing import Optional, List
import json
import logging

from PySide6 import QtCore, QtGui, QtWidgets

from core.render import render_annotated

log = logging.getLogger("ui.preview")


# ==================================================
# Theme helpers (Qt Palette based)
# ==================================================
def _pal(c: QtWidgets.QWidget) -> QtGui.QPalette:
    return c.palette()


def UI_BG(w: QtWidgets.QWidget) -> QtGui.QColor:
    return _pal(w).color(QtGui.QPalette.Base)


def UI_PANEL(w: QtWidgets.QWidget) -> QtGui.QColor:
    return _pal(w).color(QtGui.QPalette.Window)


def UI_TEXT(w: QtWidgets.QWidget) -> QtGui.QColor:
    return _pal(w).color(QtGui.QPalette.Text)


def UI_TEXT_DIM(w: QtWidgets.QWidget) -> QtGui.QColor:
    return _pal(w).color(QtGui.QPalette.PlaceholderText)


def UI_BORDER(w: QtWidgets.QWidget) -> QtGui.QColor:
    return _pal(w).color(QtGui.QPalette.Mid)


def UI_BTN_BG(w: QtWidgets.QWidget) -> QtGui.QColor:
    return _pal(w).color(QtGui.QPalette.Button)


def UI_BTN_BG_HOVER(w: QtWidgets.QWidget) -> QtGui.QColor:
    return _pal(w).color(QtGui.QPalette.Light)


def UI_ACCENT(w: QtWidgets.QWidget) -> QtGui.QColor:
    return _pal(w).color(QtGui.QPalette.Highlight)


# ==================================================
# Focus helper
# ==================================================
class FocusSavePlainTextEdit(QtWidgets.QPlainTextEdit):
    focusLost = QtCore.Signal()

    def focusOutEvent(self, e: QtGui.QFocusEvent) -> None:
        try:
            self.focusLost.emit()
        finally:
            super().focusOutEvent(e)


# ==================================================
# Overlay image view
# ==================================================
class OverlayImageView(QtWidgets.QLabel):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setScaledContents(False)
        self._pixmap: Optional[QtGui.QPixmap] = None
        self._img_w = 0
        self._img_h = 0
        self._rects: list[dict] = []
        self.setMinimumSize(320, 180)

    def set_image(self, pm: QtGui.QPixmap, img_w_px: int, img_h_px: int):
        self._pixmap = pm
        self._img_w = int(img_w_px)
        self._img_h = int(img_h_px)
        self.update()

    def set_rects_img_px(self, rects: Optional[List[dict]]):
        self._rects = rects or []
        self.update()

    def _fit_rect(self, vw: int, vh: int) -> QtCore.QRect:
        if self._img_w <= 0 or self._img_h <= 0 or vw <= 0 or vh <= 0:
            return QtCore.QRect(0, 0, 0, 0)
        sx = vw / self._img_w
        sy = vh / self._img_h
        s = min(sx, sy)
        w = int(round(self._img_w * s))
        h = int(round(self._img_h * s))
        x = (vw - w) // 2
        y = (vh - h) // 2
        return QtCore.QRect(x, y, w, h)

    def paintEvent(self, e: QtGui.QPaintEvent):
        p = QtGui.QPainter(self)
        p.setRenderHint(QtGui.QPainter.SmoothPixmapTransform, True)

        # 背景（テーマ追従）
        p.fillRect(self.rect(), UI_BG(self))

        vw, vh = self.width(), self.height()
        fit = self._fit_rect(vw, vh)

        if self._pixmap and not self._pixmap.isNull() and fit.width() > 0 and fit.height() > 0:
            p.drawPixmap(fit, self._pixmap)

            sx = fit.width() / max(1, self._img_w)
            sy = fit.height() / max(1, self._img_h)

            p.setRenderHint(QtGui.QPainter.Antialiasing, True)
            for r in self._rects:
                x = fit.left() + int(round(r.get("x", 0) * sx))
                y = fit.top() + int(round(r.get("y", 0) * sy))
                w = int(round(r.get("w", 0) * sx))
                h = int(round(r.get("h", 0) * sy))

                color = r.get("color", UI_ACCENT(self).name())
                stroke = max(1, int(round(r.get("stroke", 2) * sx)))
                pen = QtGui.QPen(QtGui.QColor(color))
                pen.setWidth(stroke)
                pen.setJoinStyle(QtCore.Qt.MiterJoin)
                p.setPen(pen)
                p.setBrush(QtCore.Qt.NoBrush)
                p.drawRect(QtCore.QRect(x, y, w, h))


# ==================================================
# Card data
# ==================================================
@dataclass
class CardData:
    json_path: Path
    image_path: Path
    folder: Path
    comment: str
    display_title: str
    ann_png_path: Optional[Path] = None

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
        if self.ann_png_path and self.ann_png_path.exists():
            self.ann_png_path.unlink()

        bp = self.base_png_abs()
        if bp.exists():
            bp.unlink()

        if self.json_path.exists():
            self.json_path.unlink()


# ==================================================
# Card widget
# ==================================================
class CardWidget(QtWidgets.QFrame):
    requestRemove = QtCore.Signal(object)
    requestRefresh = QtCore.Signal(object)

    def __init__(self, data: CardData, parent=None):
        super().__init__(parent)
        self.data = data

        self.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.setAutoFillBackground(True)
        pal = self.palette()
        pal.setColor(QtGui.QPalette.Window, UI_PANEL(self))
        pal.setColor(QtGui.QPalette.WindowText, UI_TEXT(self))
        self.setPalette(pal)

        # Header
        self.head_lbl = QtWidgets.QLabel("Title:")
        self.title_edit = QtWidgets.QLineEdit(self.data.display_title)

        # Buttons
        self.btn_open_img = QtWidgets.QToolButton(self); self.btn_open_img.setText("Open")
        self.btn_save = QtWidgets.QToolButton(self); self.btn_save.setText("Save")
        self.btn_delete = QtWidgets.QToolButton(self); self.btn_delete.setText("Delete")

        for b in (self.btn_open_img, self.btn_save, self.btn_delete):
            b.setAutoRaise(False)

        # Image
        self.image_view = OverlayImageView(self)
        self.image_view.setAlignment(QtCore.Qt.AlignCenter)

        # Comment
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
        body.addWidget(self.image_view, 0)
        body.addWidget(self.comment_edit, 1)

        lay = QtWidgets.QVBoxLayout(self)
        lay.setContentsMargins(8, 8, 8, 8)
        lay.addLayout(header)
        lay.addLayout(body)

        # Signals
        self.btn_open_img.clicked.connect(self._on_open_image)
        self.btn_save.clicked.connect(self._save_now)
        self.btn_delete.clicked.connect(self._on_delete)
        self.title_edit.editingFinished.connect(self._save_title_only)
        self.comment_edit.focusLost.connect(self._save_comment_only)

        self.refresh_image()

    # -------------------------------------------------
    def refresh_image(self):
        base_png = self.data.base_png_abs()
        try:
            meta = self.data.load_meta()
        except Exception:
            pm = QtGui.QPixmap(str(base_png))
            self.image_view.set_image(pm, pm.width(), pm.height())
            self.image_view.set_rects_img_px([])
            return

        rects_img_px = meta.get("rects_img_px")
        img_px = meta.get("image_px") or {}
        img_w = int(img_px.get("width") or 0)
        img_h = int(img_px.get("height") or 0)

        pm = QtGui.QPixmap(str(base_png))
        self.image_view.set_image(pm, img_w or pm.width(), img_h or pm.height())
        self.image_view.set_rects_img_px(rects_img_px or [])

    # -------------------------------------------------
    def _on_open_image(self):
        png = self.data.base_png_abs()
        if png.exists():
            QtGui.QDesktopServices.openUrl(QtCore.QUrl.fromLocalFile(str(png)))

    def _save_now(self):
        try:
            meta = self.data.load_meta()
            meta["display_title"] = self.title_edit.text().strip()
            meta["comment"] = self.comment_edit.toPlainText()
            self.data.save_meta(meta)
            self.requestRefresh.emit(self)
        except Exception as e:
            QtWidgets.QMessageBox.warning(self, "Save failed", str(e))

    def _save_title_only(self):
        t = self.title_edit.text().strip()
        if not t or t == self.data.display_title:
            return
        meta = self.data.load_meta()
        meta["display_title"] = t
        self.data.display_title = t
        self.data.save_meta(meta)
        self.requestRefresh.emit(self)

    def _save_comment_only(self):
        txt = self.comment_edit.toPlainText()
        if txt == self.data.comment:
            return
        meta = self.data.load_meta()
        meta["comment"] = txt
        self.data.comment = txt
        self.data.save_meta(meta)
        self.requestRefresh.emit(self)

    def _on_delete(self):
        try:
            self.data.delete_files()
        finally:
            # ★ UI から削除
            self.requestRemove.emit(self)


# ==================================================
# Preview pane
# ==================================================
class PreviewPane(QtWidgets.QScrollArea):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWidgetResizable(True)

        self.container = QtWidgets.QWidget()
        self.vbox = QtWidgets.QVBoxLayout(self.container)
        self.vbox.setContentsMargins(8, 8, 8, 8)
        self.vbox.setSpacing(12)
        self.vbox.addStretch(1)

        self.setWidget(self.container)
        self.cards: List[CardWidget] = []

    def _add_card_widget(self, cd: CardData):
        w = CardWidget(cd, self)
        w.requestRemove.connect(self._on_remove_card)
        w.requestRefresh.connect(self._on_refresh_card)
        self.vbox.insertWidget(self.vbox.count() - 1, w)
        self.cards.append(w)

    def add_capture(self, json_path: Path):
        try:
            meta = json.loads(json_path.read_text(encoding="utf-8"))
        except Exception:
            return

        folder = json_path.parent
        img_name = meta.get("image_path") or json_path.with_suffix(".png").name
        comment = meta.get("comment", "") or ""
        display_title = meta.get("display_title") or Path(img_name).stem

        cd = CardData(
            json_path=json_path,
            image_path=Path(img_name),
            folder=folder,
            comment=comment,
            display_title=display_title,
        )
        self._add_card_widget(cd)

    def _on_remove_card(self, w: CardWidget):
        try:
            self.cards.remove(w)
        except ValueError:
            pass
        w.setParent(None)
        w.deleteLater()

    def _on_refresh_card(self, _w: CardWidget):
        pass

    def clear_all(self):
        for w in list(self.cards):
            w.setParent(None)
            w.deleteLater()
        self.cards.clear()
