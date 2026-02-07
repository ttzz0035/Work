from __future__ import annotations

import logging
import tempfile
import shutil
from pathlib import Path
from typing import Any, Dict, Optional, List, Tuple

from PySide6 import QtCore, QtGui, QtWidgets

from core.calibration import Calibration

# ==================================================
# Logger (no root pollution)
# ==================================================
def _get_logger() -> logging.Logger:
    logger = logging.getLogger("Calibration.Panel")
    if not logger.handlers:
        logger.setLevel(logging.DEBUG)
        h = logging.StreamHandler()
        h.setLevel(logging.DEBUG)
        fmt = logging.Formatter("%(asctime)s [%(levelname)s] %(name)s: %(message)s")
        h.setFormatter(fmt)
        logger.addHandler(h)
        logger.propagate = False
    return logger


logger = _get_logger()


# ==================================================
# RectImageView (zoom + layers, improved visibility)
# ==================================================
class RectImageView(QtWidgets.QWidget):
    handlesChanged = QtCore.Signal()

    def __init__(self, *, view_kind: str):
        super().__init__()
        self.view_kind = view_kind
        self.setMouseTracking(True)

        self._pix: Optional[QtGui.QPixmap] = None
        self._zoom: float = 1.0
        self._auto_fit: bool = True

        self.base_rects: List[Dict[str, Any]] = []
        self.dst_rects: List[Dict[str, Any]] = []
        self.base_handles: List[Dict[str, int]] = []
        self.dst_handles: List[Dict[str, int]] = []

        self._sel_handle = 0
        self._dragging = False
        self._last_img_pos: Optional[QtCore.QPoint] = None

        self._bg = QtGui.QColor(40, 40, 40)

    # --------------------------------------------------
    def set_image(self, pm: QtGui.QPixmap):
        self._pix = pm
        self._auto_fit = True
        self._fit_to_view()
        self.update()
        if pm:
            logger.debug("[%s] set_image %dx%d zoom=%.4f", self.view_kind, pm.width(), pm.height(), self._zoom)

    def set_layers(
        self,
        *,
        base_rects,
        base_handles,
        dst_rects,
        dst_handles,
        sel_handle,
    ):
        self.base_rects = base_rects
        self.base_handles = base_handles
        self.dst_rects = dst_rects
        self.dst_handles = dst_handles
        self._sel_handle = int(sel_handle or 0)
        self.update()

    # --------------------------------------------------
    def fit(self) -> None:
        self._auto_fit = True
        self._fit_to_view()
        self.update()
        logger.debug("[%s] fit zoom=%.4f", self.view_kind, self._zoom)

    def zoom_value(self) -> float:
        return float(self._zoom)

    # --------------------------------------------------
    def resizeEvent(self, e: QtGui.QResizeEvent):
        super().resizeEvent(e)
        if self._auto_fit:
            self._fit_to_view()
            self.update()

    def _fit_to_view(self):
        if not self._pix:
            return
        if self.width() <= 0 or self.height() <= 0:
            return
        sx = self.width() / max(1, self._pix.width())
        sy = self.height() / max(1, self._pix.height())
        self._zoom = max(0.05, min(8.0, min(sx, sy)))

    # --------------------------------------------------
    def wheelEvent(self, e: QtGui.QWheelEvent):
        if not self._pix:
            return
        self._auto_fit = False

        delta = e.angleDelta().y()
        before = self._zoom
        if delta > 0:
            self._zoom *= 1.15
        else:
            self._zoom /= 1.15
        self._zoom = max(0.05, min(12.0, self._zoom))
        self.update()

        logger.debug("[%s] wheel delta=%d zoom %.4f -> %.4f", self.view_kind, int(delta), before, self._zoom)

    # --------------------------------------------------
    def _canvas_origin(self) -> Tuple[float, float, float, float]:
        if not self._pix:
            return 0.0, 0.0, 0.0, 0.0
        sw = self._pix.width() * self._zoom
        sh = self._pix.height() * self._zoom
        ox = (self.width() - sw) / 2
        oy = (self.height() - sh) / 2
        return ox, oy, sw, sh

    def _tr(self, x: float, y: float, ox: float, oy: float) -> Tuple[float, float]:
        return ox + x * self._zoom, oy + y * self._zoom

    def _to_img_pos(self, p: QtCore.QPoint) -> Optional[QtCore.QPoint]:
        if not self._pix:
            return None
        ox, oy, sw, sh = self._canvas_origin()
        if sw <= 0 or sh <= 0:
            return None
        x = (p.x() - ox) / self._zoom
        y = (p.y() - oy) / self._zoom
        if 0 <= x < self._pix.width() and 0 <= y < self._pix.height():
            return QtCore.QPoint(int(x), int(y))
        return None

    # --------------------------------------------------
    def mousePressEvent(self, e: QtGui.QMouseEvent):
        p = self._to_img_pos(e.pos())
        if not p:
            return

        hit = None
        for i, h in enumerate(self.dst_handles):
            if abs(int(h["x"]) - p.x()) <= 12 and abs(int(h["y"]) - p.y()) <= 12:
                hit = i
                break

        if hit is None:
            return

        self._sel_handle = int(hit)
        self._dragging = True
        self._last_img_pos = p
        self.update()

        logger.debug("[%s] press handle=%d img=(%d,%d)", self.view_kind, self._sel_handle, p.x(), p.y())

    def mouseMoveEvent(self, e: QtGui.QMouseEvent):
        if not self._dragging or self._last_img_pos is None:
            return
        p = self._to_img_pos(e.pos())
        if not p:
            return

        dx = p.x() - self._last_img_pos.x()
        dy = p.y() - self._last_img_pos.y()

        h = self.dst_handles[self._sel_handle]
        h["x"] = int(h["x"]) + int(dx)
        h["y"] = int(h["y"]) + int(dy)

        self._last_img_pos = p
        self.handlesChanged.emit()
        self.update()

        logger.debug("[%s] drag handle=%d dx=%d dy=%d -> (%d,%d)",
                     self.view_kind, self._sel_handle, int(dx), int(dy), int(h["x"]), int(h["y"]))

    def mouseReleaseEvent(self, e: QtGui.QMouseEvent):
        if self._dragging:
            logger.debug("[%s] release handle=%d", self.view_kind, self._sel_handle)
        self._dragging = False
        self._last_img_pos = None

    # --------------------------------------------------
    def paintEvent(self, e):
        if not self._pix:
            return

        painter = QtGui.QPainter(self)
        painter.setRenderHint(QtGui.QPainter.Antialiasing, True)

        painter.fillRect(self.rect(), self._bg)

        ox, oy, sw, sh = self._canvas_origin()

        painter.drawPixmap(
            QtCore.QRect(int(ox), int(oy), int(sw), int(sh)),
            self._pix
        )

        base_col = QtGui.QColor("#ff00ff")
        if self.view_kind == "png":
            dst_col = QtGui.QColor("#00ff00")
            dst_fill = QtGui.QColor(0, 255, 0, 80)
            dst_name = "PNG Output (adjust)"
        else:
            dst_col = QtGui.QColor("#ffd000")
            dst_fill = QtGui.QColor(255, 208, 0, 90)
            dst_name = "Excel-like Output (adjust)"

        # ---- base rects (thick, dashed for separation)
        pen_base = QtGui.QPen(base_col, 4, QtCore.Qt.PenStyle.DashLine)
        painter.setPen(pen_base)
        painter.setBrush(QtCore.Qt.BrushStyle.NoBrush)
        for r in self.base_rects:
            x, y = self._tr(float(r["x"]), float(r["y"]), ox, oy)
            w = float(r["w"]) * self._zoom
            h = float(r["h"]) * self._zoom
            painter.drawRect(QtCore.QRectF(x, y, w, h))

        # ---- dst rects
        pen_dst = QtGui.QPen(dst_col, 3)
        painter.setPen(pen_dst)
        painter.setBrush(QtCore.Qt.BrushStyle.NoBrush)
        for r in self.dst_rects:
            x, y = self._tr(float(r["x"]), float(r["y"]), ox, oy)
            w = float(r["w"]) * self._zoom
            h = float(r["h"]) * self._zoom
            painter.drawRect(QtCore.QRectF(x, y, w, h))

        # ---- base handles (bigger squares)
        pen_bh = QtGui.QPen(base_col, 3, QtCore.Qt.PenStyle.DashLine)
        painter.setPen(pen_bh)
        painter.setBrush(QtCore.Qt.BrushStyle.NoBrush)
        for h in self.base_handles:
            hx, hy = self._tr(float(h["x"]), float(h["y"]), ox, oy)
            painter.drawRect(QtCore.QRectF(hx - 9, hy - 9, 18, 18))

        # ---- dst handles (filled + crosshair, selection emphasized)
        for i, h in enumerate(self.dst_handles):
            is_sel = (i == self._sel_handle)
            size = 26 if is_sel else 20
            half = size / 2.0
            lw = 6 if is_sel else 4

            hx, hy = self._tr(float(h["x"]), float(h["y"]), ox, oy)

            painter.setPen(QtGui.QPen(dst_col, lw))
            painter.setBrush(QtGui.QBrush(dst_fill))
            painter.drawRect(QtCore.QRectF(hx - half, hy - half, size, size))

            painter.setPen(QtGui.QPen(dst_col, lw))
            painter.drawLine(QtCore.QPointF(hx - half - 6, hy), QtCore.QPointF(hx + half + 6, hy))
            painter.drawLine(QtCore.QPointF(hx, hy - half - 6), QtCore.QPointF(hx, hy + half + 6))

        # ---- legend (inside view)
        painter.setPen(QtGui.QPen(QtGui.QColor(240, 240, 240), 1))
        font = painter.font()
        font.setPointSize(max(9, int(font.pointSize())))
        painter.setFont(font)

        legend = (
            "MAGENTA = App Base (fixed)   |   "
            f"{dst_name}   |   "
            "Mouse: drag handle   Keys: arrows(1px) Shift+arrows(10px) Tab(switch handle)   Wheel: zoom"
        )
        painter.drawText(QtCore.QRectF(10, 8, self.width() - 20, 24), legend)

        # ---- zoom indicator
        painter.drawText(QtCore.QRectF(10, 30, 300, 20), f"zoom={self._zoom:.3f}")

        painter.end()


# ==================================================
# Calibration Window
# ==================================================
class CalibrationWindow(QtWidgets.QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Calibration")
        self.resize(1200, 900)

        self.cal_png = Calibration()
        self.cal_excel = Calibration()

        self._tmpdir = Path(tempfile.mkdtemp(prefix="calib_"))

        self._base_rects = [
            {"x": 120, "y": 120, "w": 24, "h": 24},
            {"x": 1500, "y": 800, "w": 24, "h": 24},
        ]
        self._base_handles = [
            {"x": 120, "y": 120},
            {"x": 1524, "y": 824},
        ]

        self._dst_png = [dict(p) for p in self._base_handles]
        self._dst_excel = [dict(p) for p in self._base_handles]

        self._sel = 0

        self._build_ui()
        self._load_pattern()

        logger.debug("[WIN] init done")

    def closeEvent(self, e: QtGui.QCloseEvent):
        logger.debug("[WIN] closeEvent")
        try:
            shutil.rmtree(self._tmpdir, ignore_errors=True)
        except Exception as ex:
            logger.debug("[WIN] tmp cleanup failed: %s", ex)
        super().closeEvent(e)

    # --------------------------------------------------
    def _build_ui(self):
        cw = QtWidgets.QWidget()
        self.setCentralWidget(cw)
        root = QtWidgets.QVBoxLayout(cw)

        top = QtWidgets.QHBoxLayout()
        root.addLayout(top)

        self.lbl_param = QtWidgets.QLineEdit()
        self.lbl_param.setReadOnly(True)

        btn_fit = QtWidgets.QPushButton("Fit")
        btn_fit.clicked.connect(self._fit_current)

        top.addWidget(self.lbl_param, 1)
        top.addWidget(btn_fit)

        self.tabs = QtWidgets.QTabWidget()
        root.addWidget(self.tabs, 1)

        self.view_png = RectImageView(view_kind="png")
        self.view_excel = RectImageView(view_kind="excel")

        self.view_png.handlesChanged.connect(self._on_png_changed)
        self.view_excel.handlesChanged.connect(self._on_excel_changed)

        self.tabs.addTab(self.view_png, "PNG")
        self.tabs.addTab(self.view_excel, "Excel-like")

        self.tabs.currentChanged.connect(self._on_tab_changed)

        self.setFocusPolicy(QtCore.Qt.FocusPolicy.StrongFocus)

    def _fit_current(self) -> None:
        if self.tabs.currentIndex() == 0:
            self.view_png.fit()
        else:
            self.view_excel.fit()
        logger.debug("[WIN] fit_current tab=%d", int(self.tabs.currentIndex()))

    def _on_tab_changed(self, idx: int) -> None:
        self._update_param()
        self._fit_current()
        logger.debug("[WIN] tab_changed idx=%d", int(idx))

    # --------------------------------------------------
    def _load_pattern(self):
        pm = QtGui.QPixmap(1600, 900)
        pm.fill(QtGui.QColor(230, 230, 230))
        self.view_png.set_image(pm)
        self.view_excel.set_image(pm)
        self._refresh()
        logger.debug("[WIN] pattern loaded")

    # --------------------------------------------------
    def keyPressEvent(self, e: QtGui.QKeyEvent):
        step = 10 if (e.modifiers() & QtCore.Qt.KeyboardModifier.ShiftModifier) else 1

        if e.key() == QtCore.Qt.Key.Key_F:
            self._fit_current()
            return

        if e.key() == QtCore.Qt.Key.Key_Tab:
            self._sel = 1 - int(self._sel)
            self._refresh()
            logger.debug("[KEY] switch handle sel=%d", int(self._sel))
            return

        dx = dy = 0
        if e.key() == QtCore.Qt.Key.Key_Left:
            dx = -step
        elif e.key() == QtCore.Qt.Key.Key_Right:
            dx = step
        elif e.key() == QtCore.Qt.Key.Key_Up:
            dy = -step
        elif e.key() == QtCore.Qt.Key.Key_Down:
            dy = step
        else:
            return

        if self.tabs.currentIndex() == 0:
            self._dst_png[self._sel]["x"] += int(dx)
            self._dst_png[self._sel]["y"] += int(dy)
            self._solve_png()
            logger.debug("[KEY] png move sel=%d dx=%d dy=%d -> (%d,%d)",
                         int(self._sel), int(dx), int(dy),
                         int(self._dst_png[self._sel]["x"]), int(self._dst_png[self._sel]["y"]))
        else:
            self._dst_excel[self._sel]["x"] += int(dx)
            self._dst_excel[self._sel]["y"] += int(dy)
            self._solve_excel()
            logger.debug("[KEY] excel move sel=%d dx=%d dy=%d -> (%d,%d)",
                         int(self._sel), int(dx), int(dy),
                         int(self._dst_excel[self._sel]["x"]), int(self._dst_excel[self._sel]["y"]))

        self._refresh()

    # --------------------------------------------------
    def _solve_png(self):
        self.cal_png = self._solve(self._dst_png)
        logger.debug("[CAL] png scale=%.6f off=(%.2f,%.2f)", self.cal_png.scale, self.cal_png.off_x, self.cal_png.off_y)

    def _solve_excel(self):
        self.cal_excel = self._solve(self._dst_excel)
        logger.debug("[CAL] excel scale=%.6f off=(%.2f,%.2f)", self.cal_excel.scale, self.cal_excel.off_x, self.cal_excel.off_y)

    def _solve(self, dst):
        p0, p1 = self._base_handles
        q0, q1 = dst

        dx_b = float(p1["x"] - p0["x"])
        dy_b = float(p1["y"] - p0["y"])
        dx_o = float(q1["x"] - q0["x"])
        dy_o = float(q1["y"] - q0["y"])

        sx = (dx_o / dx_b) if dx_b else 1.0
        sy = (dy_o / dy_b) if dy_b else 1.0
        scale = (sx + sy) / 2.0
        if scale == 0.0:
            scale = 1.0

        off_x = float(q0["x"]) - float(p0["x"]) * scale
        off_y = float(q0["y"]) - float(p0["y"]) * scale

        return Calibration(scale, off_x, off_y)

    # --------------------------------------------------
    def _refresh(self):
        self.view_png.set_layers(
            base_rects=self._base_rects,
            base_handles=self._base_handles,
            dst_rects=self._apply(self.cal_png),
            dst_handles=self._dst_png,
            sel_handle=self._sel,
        )
        self.view_excel.set_layers(
            base_rects=self._base_rects,
            base_handles=self._base_handles,
            dst_rects=self._apply(self.cal_excel),
            dst_handles=self._dst_excel,
            sel_handle=self._sel,
        )
        self._update_param()

    def _apply(self, cal):
        out = []
        for r in self._base_rects:
            out.append(
                {
                    "x": int(r["x"] * cal.scale + cal.off_x),
                    "y": int(r["y"] * cal.scale + cal.off_y),
                    "w": max(1, int(r["w"] * cal.scale)),
                    "h": max(1, int(r["h"] * cal.scale)),
                }
            )
        return out

    def _update_param(self):
        cal = self.cal_png if self.tabs.currentIndex() == 0 else self.cal_excel
        self.lbl_param.setText(f"scale={cal.scale:.6f} off=({cal.off_x:.2f},{cal.off_y:.2f})")

    def _on_png_changed(self):
        self._dst_png = [dict(p) for p in self.view_png.dst_handles]
        self._sel = int(self.view_png._sel_handle)
        self._solve_png()
        self._refresh()
        logger.debug("[SIG] png_changed sel=%d", int(self._sel))

    def _on_excel_changed(self):
        self._dst_excel = [dict(p) for p in self.view_excel.dst_handles]
        self._sel = int(self.view_excel._sel_handle)
        self._solve_excel()
        self._refresh()
        logger.debug("[SIG] excel_changed sel=%d", int(self._sel))
