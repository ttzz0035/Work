# ui/region/constants.py
from __future__ import annotations

from PySide6 import QtGui, QtWidgets

# ==================================================
# Geometry / Handles
# ==================================================
HANDLE_SIZE = 10
RECT_CLOSE_SIZE = 16

MIN_ANNO_W = 12
MIN_ANNO_H = 12

MIN_REGION_W = 120
MIN_REGION_H = 90


# ==================================================
# Layout
# ==================================================
TOP_MARGIN = 36
BOTTOM_MARGIN = 60
SIDE_MARGIN = 12


# ==================================================
# Theme helpers (PALETTE BASED)
# ==================================================
def _app_palette() -> QtGui.QPalette:
    app = QtWidgets.QApplication.instance()
    if not app:
        return QtGui.QPalette()
    return app.palette()


def UI_BG() -> QtGui.QColor:
    return _app_palette().color(QtGui.QPalette.Window)


def UI_PANEL_BG() -> QtGui.QColor:
    return _app_palette().color(QtGui.QPalette.Base)


def UI_BORDER() -> QtGui.QColor:
    return _app_palette().color(QtGui.QPalette.Mid)


def UI_TEXT() -> QtGui.QColor:
    return _app_palette().color(QtGui.QPalette.WindowText)


def UI_TEXT_DIM() -> QtGui.QColor:
    return _app_palette().color(QtGui.QPalette.Disabled, QtGui.QPalette.WindowText)


def UI_BTN_BG() -> QtGui.QColor:
    return _app_palette().color(QtGui.QPalette.Button)


def UI_BTN_BG_HOVER() -> QtGui.QColor:
    # hover は Button を少し明るく
    c = UI_BTN_BG()
    return c.lighter(110)


def UI_BTN_BORDER() -> QtGui.QColor:
    return _app_palette().color(QtGui.QPalette.Dark)


def UI_HANDLE_BG() -> QtGui.QColor:
    return UI_BTN_BG()


def UI_HANDLE_BORDER() -> QtGui.QColor:
    return UI_BTN_BORDER()


def UI_ACCENT() -> QtGui.QColor:
    return _app_palette().color(QtGui.QPalette.Highlight)


# ==================================================
# Annotation defaults
# ==================================================
def DEFAULT_NEW_RECT_COLOR() -> QtGui.QColor:
    return UI_ACCENT()


DEFAULT_NEW_RECT_STROKE = 2


# ==================================================
# Frame
# ==================================================
def FRAME_PEN_COLOR() -> QtGui.QColor:
    return UI_ACCENT()


FRAME_PEN_WIDTH = 4.0
