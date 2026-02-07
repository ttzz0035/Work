# ui/toolbar.py
from __future__ import annotations

import logging
from typing import Dict, Optional

from PySide6 import QtCore, QtGui, QtWidgets


# ==================================================
# Logger (no root pollution)
# ==================================================
def _get_logger() -> logging.Logger:
    logger = logging.getLogger("ui.toolbar")
    if not logger.handlers:
        logger.setLevel(logging.DEBUG)
        h = logging.StreamHandler()
        h.setLevel(logging.DEBUG)
        fmt = logging.Formatter("%(asctime)s [%(levelname)s] %(name)s: %(message)s")
        h.setFormatter(fmt)
        logger.addHandler(h)
        logger.propagate = False
    return logger


log = _get_logger()


class FocusFriendlySpinBox(QtWidgets.QSpinBox):
    def __init__(self, *a, **kw):
        super().__init__(*a, **kw)
        self.setFocusPolicy(QtCore.Qt.NoFocus)

    def focusInEvent(self, e: QtGui.QFocusEvent):
        try:
            self.clearFocus()
        except Exception:
            pass
        e.ignore()

    def event(self, ev: QtCore.QEvent) -> bool:
        if ev.type() in (QtCore.QEvent.FocusIn, QtCore.QEvent.FocusAboutToChange):
            return True
        return super().event(ev)

    def keyPressEvent(self, e: QtGui.QKeyEvent):
        if e.modifiers() != QtCore.Qt.NoModifier or e.key() == QtCore.Qt.Key_Delete:
            e.ignore()
            return
        super().keyPressEvent(e)

    def wheelEvent(self, e: QtGui.QWheelEvent):
        if not self.hasFocus():
            e.ignore()
            return
        super().wheelEvent(e)


class MiniToolbar(QtWidgets.QWidget):
    shotClicked = QtCore.Signal()
    rectClicked = QtCore.Signal()
    colorClicked = QtCore.Signal()
    settingsClicked = QtCore.Signal()

    recToggleClicked = QtCore.Signal()
    playToggleClicked = QtCore.Signal()

    newColorPicked = QtCore.Signal(QtGui.QColor)
    newStrokePicked = QtCore.Signal(int)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setObjectName("MiniToolbar")

        # ------------------------------
        # Internal state (single source of truth)
        # ------------------------------
        self._current_color: QtGui.QColor = QtGui.QColor("#FF3B30")
        self._current_stroke: int = 2

        lay = QtWidgets.QHBoxLayout(self)
        lay.setContentsMargins(8, 6, 8, 6)
        lay.setSpacing(6)

        self.btnShot     = QtWidgets.QToolButton(self); self.btnShot.setText("Capture")
        self.btnRect     = QtWidgets.QToolButton(self); self.btnRect.setText("AddRect")
        self.btnColor    = QtWidgets.QToolButton(self); self.btnColor.setText("Color")
        self.btnSettings = QtWidgets.QToolButton(self); self.btnSettings.setText("Settings")

        self.btnRec  = QtWidgets.QToolButton(self); self.btnRec.setCheckable(True);  self.btnRec.setText("● Rec")
        self.btnPlay = QtWidgets.QToolButton(self); self.btnPlay.setCheckable(True); self.btnPlay.setText("▶ Play")

        self.spinStroke = FocusFriendlySpinBox(self)
        self.spinStroke.setRange(1, 20)
        self.spinStroke.setValue(self._current_stroke)

        self._all_buttons = (
            self.btnShot,
            self.btnRect,
            self.btnColor,
            self.btnSettings,
            self.btnRec,
            self.btnPlay,
        )

        for w in (*self._all_buttons, self.spinStroke):
            lay.addWidget(w)

        # signals
        self.btnShot.clicked.connect(self.shotClicked.emit)
        self.btnRect.clicked.connect(self.rectClicked.emit)
        self.btnSettings.clicked.connect(self.settingsClicked.emit)
        self.btnColor.clicked.connect(self._on_pick_color)
        self.spinStroke.valueChanged.connect(self._on_stroke_changed)
        self.btnRec.toggled.connect(self._on_rec_toggled)
        self.btnPlay.toggled.connect(self._on_play_toggled)

        # Initial swatch sync
        self._apply_color(self._current_color, emit_signal=False, update_ui=True)

        log.debug("MiniToolbar init color=%s stroke=%d", self.current_color_name(), self.current_stroke())

    # -------------------------------------------------
    # Public getters (caller must use these)
    # -------------------------------------------------
    def current_color(self) -> QtGui.QColor:
        return QtGui.QColor(self._current_color)

    def current_color_name(self) -> str:
        # always #RRGGBB
        return QtGui.QColor(self._current_color).name()

    def current_stroke(self) -> int:
        return int(self._current_stroke)

    # -------------------------------------------------
    # 色をツールバー全体に反映
    # -------------------------------------------------
    def update_swatch(self, color: QtGui.QColor):
        # UI only; state update must go through _apply_color
        if not isinstance(color, QtGui.QColor):
            return

        col = color.name()
        text_col = "#111" if self._luma(color) > 0.5 else "#fff"

        btn_style = (
            "QToolButton{"
            f"background:{col};"
            f"color:{text_col};"
            "border:1px solid rgba(0,0,0,0.4);"
            "border-radius:6px;"
            "padding:4px 8px;"
            "}"
            "QToolButton:hover{"
            f"background:{col};"
            "border:1px solid rgba(0,0,0,0.6);"
            "}"
        )

        spin_style = (
            "QSpinBox{"
            f"background:{col};"
            f"color:{text_col};"
            "border:1px solid rgba(0,0,0,0.4);"
            "border-radius:6px;"
            "padding:2px 6px;"
            "}"
            "QSpinBox::up-button, QSpinBox::down-button{"
            "width:14px; border:none;"
            "}"
        )

        for b in self._all_buttons:
            b.setStyleSheet(btn_style)

        self.spinStroke.setStyleSheet(spin_style)

    # -------------------------------------------------
    # Internal state apply
    # -------------------------------------------------
    def _apply_color(self, color: QtGui.QColor, *, emit_signal: bool, update_ui: bool):
        if not isinstance(color, QtGui.QColor):
            return
        if not color.isValid():
            return

        old = self._current_color.name()
        self._current_color = QtGui.QColor(color)

        if update_ui:
            self.update_swatch(self._current_color)

        log.debug("toolbar color change %s -> %s", old, self._current_color.name())

        if emit_signal:
            try:
                self.newColorPicked.emit(QtGui.QColor(self._current_color))
            except Exception as e:
                log.exception("emit newColorPicked failed: %s", e)

    def _apply_stroke(self, v: int, *, emit_signal: bool):
        old = int(self._current_stroke)
        self._current_stroke = int(v)
        log.debug("toolbar stroke change %d -> %d", old, self._current_stroke)

        if emit_signal:
            try:
                self.newStrokePicked.emit(int(self._current_stroke))
            except Exception as e:
                log.exception("emit newStrokePicked failed: %s", e)

    # -------------------------------------------------
    def _on_pick_color(self):
        self.colorClicked.emit()
        col = QtWidgets.QColorDialog.getColor(initial=self._current_color, parent=self, title="Pick Color")
        if col.isValid():
            self._apply_color(col, emit_signal=True, update_ui=True)

    def _on_stroke_changed(self, v: int):
        try:
            self._apply_stroke(int(v), emit_signal=True)
        except Exception as e:
            log.exception("stroke change failed: %s", e)

    def _on_rec_toggled(self, on: bool):
        self.btnRec.setText("■ Rec" if on else "● Rec")
        self.recToggleClicked.emit()

    def _on_play_toggled(self, on: bool):
        self.btnPlay.setText("■ Stop" if on else "▶ Play")
        self.playToggleClicked.emit()

    @staticmethod
    def _luma(c: QtGui.QColor) -> float:
        r, g, b = c.redF(), c.greenF(), c.blueF()
        return 0.2126 * r + 0.7152 * g + 0.0722 * b
