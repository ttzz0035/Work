# ui/region/region_window.py
from __future__ import annotations

import logging
import time
from pathlib import Path
from typing import Optional, Tuple

from PySide6 import QtCore, QtGui, QtWidgets

from core.config import Config, HotkeyManager, load_last_state, save_last_state
from core.recorder import InputRecorder, InputPlayer
from core.capture_service import (
    CaptureService,
    CaptureRequest,
    CaptureAnnoRect,
    CaptureRegionLocal,
    CaptureGlobalTopLeft,
)

from ui.toolbar import MiniToolbar
from ui.preview import PreviewPane
from ui.settings_dialog import SettingsDialog

from ui.region.constants import (
    HANDLE_SIZE,
    RECT_CLOSE_SIZE,
    MIN_ANNO_W,
    MIN_ANNO_H,
    MIN_REGION_W,
    MIN_REGION_H,
    FRAME_PEN_WIDTH,
    TOP_MARGIN,
    BOTTOM_MARGIN,
    SIDE_MARGIN,
    DEFAULT_NEW_RECT_STROKE,

    # --- UI theme ---
    UI_BTN_BG,
    UI_BTN_BORDER,
    UI_TEXT,
    UI_ACCENT,
)

from ui.region.geometry import (
    handle_rects,
    clamp_inside,
    rect_close_rect,
)
from ui.region.annotation import AnnotationManager


# ==================================================
# Logger (no root pollution)
# ==================================================
def _get_logger() -> logging.Logger:
    logger = logging.getLogger("UI.RegionWindow")
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
# Utils
# ==================================================
def _to_int_flag(v) -> int:
    """
    PySide6 の Qt.* 列挙（MouseButton/MouseButtons/KeyboardModifiers など）は
    環境によって int() 変換に失敗することがあるため、.value を優先的に用いて整数化する。
    """
    try:
        return int(v)
    except (TypeError, ValueError):
        return int(getattr(v, "value", 0))


# ==================================================
# RegionWindow
# ==================================================
class RegionWindow(QtWidgets.QWidget):
    def __init__(self, preview: Optional[PreviewPane] = None, save_dir: Optional[Path] = None):
        super().__init__()
        self.logger = logger
        self.capture_service = CaptureService()

        self.setWindowFlags(
            QtCore.Qt.FramelessWindowHint
            | QtCore.Qt.WindowStaysOnTopHint
        )
        self.setAttribute(QtCore.Qt.WA_TranslucentBackground, True)
        self.setAttribute(QtCore.Qt.WA_NoSystemBackground, True)
        self.setMouseTracking(True)
        self.setFocusPolicy(QtCore.Qt.StrongFocus)

        self.save_dir: Path = Path(save_dir) if save_dir else Path(__file__).resolve().parents[2] / "out"
        self.records_dir: Path = self.save_dir / "_records"
        self.records_dir.mkdir(parents=True, exist_ok=True)

        scr = QtGui.QGuiApplication.primaryScreen().availableGeometry()
        w = min(1000, scr.width() - 200)
        h = min(560, scr.height() - 200)
        self.resize(w, h)
        self.move(scr.center() - QtCore.QPoint(self.width() // 2, self.height() // 2))

        # ==================================================
        # Config
        # ==================================================
        self.conf = Config()
        self.conf.load()
        self.logger.debug("config loaded toolbar=%s ui=%s record=%s", self.conf.toolbar, getattr(self.conf, "ui", {}), getattr(self.conf, "record", {}))

        # record last_dir -> records_dir
        try:
            last_dir = ""
            if hasattr(self.conf, "record") and isinstance(self.conf.record, dict):
                last_dir = str(self.conf.record.get("last_dir") or "")
            if last_dir:
                self.records_dir = Path(last_dir)
                self.records_dir.mkdir(parents=True, exist_ok=True)
                self.logger.debug("records_dir overridden by config: %s", self.records_dir)
        except Exception as e:
            self.logger.exception("apply record.last_dir failed: %s", e)

        # toast duration
        self._toast_default_sec = 1.2
        try:
            if hasattr(self.conf, "ui") and isinstance(self.conf.ui, dict):
                v = self.conf.ui.get("toast_duration", 1.2)
                self._toast_default_sec = float(v)
            self.logger.debug("toast_duration=%.3f", float(self._toast_default_sec))
        except Exception as e:
            self.logger.exception("apply ui.toast_duration failed: %s", e)
            self._toast_default_sec = 1.2

        self.ann = AnnotationManager(self)
        self.ann.changed.connect(self.update)

        # toolbar defaults from config
        try:
            col_hex = str(self.conf.toolbar.get("rect_color", "#FF3B30"))
        except Exception:
            col_hex = "#FF3B30"
        try:
            stroke_v = int(self.conf.toolbar.get("rect_stroke", DEFAULT_NEW_RECT_STROKE))
        except Exception:
            stroke_v = int(DEFAULT_NEW_RECT_STROKE)

        self._new_rect_color = QtGui.QColor(col_hex)
        if not self._new_rect_color.isValid():
            self._new_rect_color = QtGui.QColor("#FF3B30")
        self._new_rect_stroke = max(1, int(stroke_v))

        self.drag_mode: Optional[str] = None
        self.resize_handle: Optional[str] = None
        self.drag_start_local = QtCore.QPoint()
        self.drag_start_global = QtCore.QPoint()
        self.orig_window = QtCore.QRect()
        self.orig_region = QtCore.QRect()
        self.orig_rect = QtCore.QRect()
        self._dragging = False

        self.preview = preview

        self.toolbar = MiniToolbar(self)
        self.toolbar.shotClicked.connect(self.capture_now)
        self.toolbar.rectClicked.connect(self._add_rect_from_preset)
        self.toolbar.newColorPicked.connect(self._set_new_color)
        self.toolbar.newStrokePicked.connect(self._set_new_stroke)
        self.toolbar.settingsClicked.connect(lambda: self._open_settings(SettingsDialog))
        self.toolbar.recToggleClicked.connect(self._toggle_record)
        self.toolbar.playToggleClicked.connect(self._toggle_play)

        # reflect config -> toolbar ui
        try:
            self.toolbar.update_swatch(self._new_rect_color)
            self.toolbar.spinStroke.setValue(int(self._new_rect_stroke))
        except Exception as e:
            self.logger.exception("apply toolbar UI defaults failed: %s", e)

        self.toolbar.show()

        self.close_btn = QtWidgets.QToolButton(self)
        self.close_btn.setText("×")
        self.close_btn.setFixedSize(22, 22)
        self.close_btn.setStyleSheet(
            "QToolButton{background:white;color:#111;border:1px solid #ccc;border-radius:4px;}"
            "QToolButton:hover{background:#f2f2f2;}"
        )
        self.close_btn.clicked.connect(self._close_region)
        self.close_btn.show()

        self.recorder = InputRecorder(self)
        self.player: Optional[InputPlayer] = None
        self._recording = False
        self._playing = False

        self._restore_last_state()
        self._place_children()

        self.hotkeys = HotkeyManager(
            self,
            {
                "capture": self.capture_now,
                "add_rect": self._add_rect_from_preset,
                "remove_selected": self.ann.remove_selected,
                "pick_new_color": self._pick_new_color,
                "show_hotkeys": self._show_hotkeys_dialog,
                "exit_app": self._close_region,
                "rec_start": self._toggle_record,
                "rec_stop": self._toggle_record,
                "rec_play": self._toggle_play,
            },
            self.conf,
        )
        self.hotkeys.apply()

        self._toast = ""
        self._toast_until = 0.0

    # -------------------------------------------------
    # Geometry helpers
    # -------------------------------------------------
    def _region_rect(self) -> QtCore.QRect:
        return QtCore.QRect(
            SIDE_MARGIN,
            TOP_MARGIN,
            max(1, self.width() - 2 * SIDE_MARGIN),
            max(1, self.height() - TOP_MARGIN - BOTTOM_MARGIN),
        )

    def _frame_rects(self) -> Tuple[QtCore.QRectF, QtCore.QRect]:
        reg = self._region_rect()
        inner_f = QtCore.QRectF(reg).adjusted(
            FRAME_PEN_WIDTH / 2,
            FRAME_PEN_WIDTH / 2,
            -FRAME_PEN_WIDTH / 2,
            -FRAME_PEN_WIDTH / 2,
        )
        return inner_f, inner_f.toAlignedRect()

    # -------------------------------------------------
    # State
    # -------------------------------------------------
    def _restore_last_state(self):
        self.logger.debug("restore_last_state start")
        data = load_last_state()
        rg = data.get("region")
        if isinstance(rg, dict):
            left = int(rg.get("left", self.x()))
            top = int(rg.get("top", self.y()))
            width = max(MIN_REGION_W, int(rg.get("width", self.width())))
            height = max(MIN_REGION_H, int(rg.get("height", self.height())))
            self.setGeometry(QtCore.QRect(left, top, width, height))
            self.logger.debug("restore region geometry=%s", (left, top, width, height))

        rects = data.get("rects", [])
        if isinstance(rects, list):
            self.ann.annos.clear()
            for r in rects:
                try:
                    qr = clamp_inside(
                        QtCore.QRect(
                            int(r.get("x", 16)),
                            int(r.get("y", 16)),
                            int(r.get("w", 160)),
                            int(r.get("h", 90)),
                        ),
                        self._region_rect(),
                        min_w=MIN_ANNO_W,
                        min_h=MIN_ANNO_H,
                    )
                    self.ann.add(
                        qr.x(),
                        qr.y(),
                        qr.width(),
                        qr.height(),
                        color=str(r.get("color", self._new_rect_color.name())),
                        stroke=int(r.get("stroke", self._new_rect_stroke)),
                    )
                except Exception:
                    continue
            self.ann.selected = None
            self.logger.debug("restore rects count=%d", len(self.ann.annos))

        self.logger.debug("restore_last_state end")

    def _save_last_state(self):
        self.logger.debug("save_last_state start")
        g = self.frameGeometry()
        data = {
            "region": {
                "left": g.left(),
                "top": g.top(),
                "width": g.width(),
                "height": g.height(),
            },
            "rects": [
                {
                    "x": a.x,
                    "y": a.y,
                    "w": a.w,
                    "h": a.h,
                    "color": a.color,
                    "stroke": a.stroke,
                }
                for a in self.ann.annos
            ],
        }
        save_last_state(data)
        self.logger.debug("save_last_state end region=%s rects=%d", data.get("region"), len(data.get("rects", [])))

    # -------------------------------------------------
    # Layout
    # -------------------------------------------------
    def _place_children(self):
        self.toolbar.adjustSize()
        tb_w, tb_h = self.toolbar.width(), self.toolbar.height()
        self.toolbar.move(
            max(0, (self.width() - tb_w) // 2),
            self.height() - BOTTOM_MARGIN + (BOTTOM_MARGIN - tb_h) // 2,
        )

        self.close_btn.move(
            self.width() - SIDE_MARGIN - self.close_btn.width(),
            (TOP_MARGIN - self.close_btn.height()) // 2,
        )

    def resizeEvent(self, e: QtGui.QResizeEvent):
        super().resizeEvent(e)
        self._place_children()

    def showEvent(self, e: QtGui.QShowEvent):
        super().showEvent(e)
        self._place_children()

        self.activateWindow()
        self.raise_()
        self.setFocus(QtCore.Qt.ActiveWindowFocusReason)

    def closeEvent(self, e: QtGui.QCloseEvent):
        try:
            self._save_last_state()
        finally:
            super().closeEvent(e)

    # -------------------------------------------------
    # Paint
    # -------------------------------------------------
    def paintEvent(self, _e):
        p = QtGui.QPainter(self)
        p.setRenderHint(QtGui.QPainter.Antialiasing, not self._dragging)

        inner_f, frame = self._frame_rects()

        pal = self.palette()

        frame_color = pal.color(QtGui.QPalette.Highlight)
        btn_border  = pal.color(QtGui.QPalette.Mid)
        btn_bg      = pal.color(QtGui.QPalette.Button)
        text_color  = pal.color(QtGui.QPalette.Text)

        # --- region frame（赤固定） ---
        pen = QtGui.QPen(QtGui.QColor("#FF3B30"))
        pen.setWidthF(FRAME_PEN_WIDTH)
        p.setPen(pen)
        p.setBrush(QtCore.Qt.NoBrush)
        p.drawRect(inner_f)

        # --- window resize handles ---
        p.setPen(QtGui.QPen(btn_border, 1))
        p.setBrush(btn_bg)
        for hr in handle_rects(frame, HANDLE_SIZE).values():
            p.drawRect(hr)

        # --- annotations ---
        bounds = self._region_rect()
        for a in self.ann.annos:
            r = clamp_inside(
                QtCore.QRect(a.x, a.y, a.w, a.h),
                bounds,
                min_w=MIN_ANNO_W,
                min_h=MIN_ANNO_H,
            )

            # 矩形本体（色は a.color）
            p.setPen(QtGui.QPen(QtGui.QColor(a.color), a.stroke))
            p.setBrush(QtCore.Qt.NoBrush)
            p.drawRect(r)

            # close button
            rc = rect_close_rect(r, RECT_CLOSE_SIZE)
            self._paint_close_icon(p, rc)

            # rect resize handles
            p.setPen(QtGui.QPen(btn_border, 1))
            p.setBrush(btn_bg)
            for hrect in handle_rects(r, HANDLE_SIZE).values():
                p.drawRect(hrect)

    def _paint_close_icon(self, p: QtGui.QPainter, rc: QtCore.QRect):
        pal = self.palette()

        btn_border = pal.color(QtGui.QPalette.Mid)
        btn_bg     = pal.color(QtGui.QPalette.Button)
        text_color = pal.color(QtGui.QPalette.Text)

        p.setPen(QtGui.QPen(btn_border, 1))
        p.setBrush(btn_bg)
        p.drawRect(rc)

        p.setPen(QtGui.QPen(text_color, 2))
        p.drawLine(
            rc.left() + 4, rc.top() + 4,
            rc.right() - 4, rc.bottom() - 4,
        )
        p.drawLine(
            rc.left() + 4, rc.bottom() - 4,
            rc.right() - 4, rc.top() + 4,
        )

    # -------------------------------------------------
    # Input (mouse / key) + Recorder hooks
    # -------------------------------------------------
    def wheelEvent(self, e: QtGui.QWheelEvent):
        if self._recording:
            delta = e.angleDelta().y() or e.angleDelta().x() or 0
            self.recorder.on_mouse(
                "wheel",
                e.position().toPoint(),
                e.globalPosition().toPoint(),
                _to_int_flag(e.buttons()),
                0,
                int(delta),
            )
        e.accept()

    def mousePressEvent(self, e: QtGui.QMouseEvent):
        if self._recording:
            self.recorder.on_mouse(
                "press",
                e.position().toPoint(),
                e.globalPosition().toPoint(),
                _to_int_flag(e.buttons()),
                _to_int_flag(e.button()),
            )

        pos_local = e.position().toPoint()
        pos_global = e.globalPosition().toPoint()
        _, frame = self._frame_rects()
        reg = self._region_rect()

        # ----------------------------
        # close button (annotation)
        # ----------------------------
        idx = self.ann.hit_close(
            pos_local,
            close_rect_fn=lambda r: rect_close_rect(r),
        )
        if idx is not None:
            self.ann.remove_at(idx)
            return

        # ----------------------------
        # resize annotation
        # ----------------------------
        idx, h = self.ann.hit_handle(
            pos_local,
            handle_rects_fn=lambda r: handle_rects(r, HANDLE_SIZE),
        )
        if idx is not None and h:
            self.ann.selected = idx
            self.drag_mode = "resize_anno"
            self.resize_handle = h
            self.drag_start_local = pos_local
            self.orig_rect = self.ann.qrect(idx)
            self._start_drag()
            return

        # ----------------------------
        # move annotation
        #   ※ 判定を少し広げる（pad=10）
        # ----------------------------
        idx = (
            self.ann.hit_body(pos_local)
            or self.ann.hit_body_expanded(pos_local, pad=10)
        )
        if idx is not None:
            self.ann.selected = idx
            self.drag_mode = "move_anno"
            self.drag_start_local = pos_local
            self.orig_rect = self.ann.qrect(idx)
            self._start_drag()
            return

        # ----------------------------
        # resize window
        # ----------------------------
        for k, hr in handle_rects(frame, HANDLE_SIZE).items():
            if hr.contains(pos_local):
                self.drag_mode = "resize_win"
                self.resize_handle = k
                self.drag_start_global = pos_global
                self.orig_window = self.frameGeometry()
                self._start_drag(grab_mouse=True)
                return

        # ----------------------------
        # move window
        # ----------------------------
        if reg.contains(pos_local):
            self.drag_mode = "move_win"
            self.drag_start_global = pos_global
            self.orig_window = self.frameGeometry()
            self._start_drag(grab_mouse=True)
            return

        # ----------------------------
        # reset
        # ----------------------------
        self.drag_mode = None
        self.ann.selected = None
        self.update()

    def mouseMoveEvent(self, e: QtGui.QMouseEvent):
        if self._recording:
            self.recorder.on_mouse(
                "move",
                e.position().toPoint(),
                e.globalPosition().toPoint(),
                _to_int_flag(e.buttons()),
                0,
            )

        pos_local = e.position().toPoint()
        pos_global = e.globalPosition().toPoint()
        reg = self._region_rect()

        if not self.drag_mode:
            self.update()
            return

        if self.drag_mode == "move_win":
            d = pos_global - self.drag_start_global
            g = QtCore.QRect(self.orig_window)
            g.translate(d)
            self.move(g.topLeft())
            return

        if self.drag_mode == "resize_win":
            dx = pos_global.x() - self.drag_start_global.x()
            dy = pos_global.y() - self.drag_start_global.y()
            ow = QtCore.QRect(self.orig_window)

            nx, ny, nr, nb = ow.left(), ow.top(), ow.right(), ow.bottom()
            h = self.resize_handle or ""

            if "e" in h:
                nr += dx
            if "w" in h:
                nx += dx
            if "s" in h:
                nb += dy
            if "n" in h:
                ny += dy

            w = max(MIN_REGION_W, nr - nx + 1)
            h2 = max(MIN_REGION_H, nb - ny + 1)
            self.setGeometry(QtCore.QRect(nx, ny, w, h2))
            return

        dx = pos_local.x() - self.drag_start_local.x()
        dy = pos_local.y() - self.drag_start_local.y()
        bounds = self._region_rect()

        if self.drag_mode == "move_anno" and self.ann.selected is not None:
            r = QtCore.QRect(self.orig_rect)
            r.translate(dx, dy)
            r = clamp_inside(r, bounds, MIN_ANNO_W, MIN_ANNO_H)
            self.ann.move_to(self.ann.selected, r)
            return

        if self.drag_mode == "resize_anno" and self.ann.selected is not None:
            r = QtCore.QRect(self.orig_rect)
            h = self.resize_handle or ""
            if "n" in h:
                r.setTop(r.top() + dy)
            if "s" in h:
                r.setBottom(r.bottom() + dy)
            if "w" in h:
                r.setLeft(r.left() + dx)
            if "e" in h:
                r.setRight(r.right() + dx)
            r = clamp_inside(r, bounds, MIN_ANNO_W, MIN_ANNO_H)
            self.ann.move_to(self.ann.selected, r)
            return

    def mouseReleaseEvent(self, e: QtGui.QMouseEvent):
        if self._recording:
            self.recorder.on_mouse(
                "release",
                e.position().toPoint(),
                e.globalPosition().toPoint(),
                _to_int_flag(e.buttons()),
                _to_int_flag(e.button()),
            )
        self.drag_mode = None
        self.resize_handle = None
        self._dragging = False
        if self.mouseGrabber():
            self.releaseMouse()
        self.update()

    def keyPressEvent(self, e: QtGui.QKeyEvent):
        self.logger.debug(
            "keyPressEvent key=%s mod=%s focus=%s",
            e.key(),
            _to_int_flag(e.modifiers()),
            self.hasFocus(),
        )

        if self._recording:
            self.recorder.on_key(
                "keyPress",
                e.key(),
                _to_int_flag(e.modifiers()),
                e.text(),
            )

        super().keyPressEvent(e)

    def keyReleaseEvent(self, e: QtGui.QKeyEvent):
        if self._recording:
            self.recorder.on_key(
                "keyRelease",
                e.key(),
                _to_int_flag(e.modifiers()),
                e.text(),
            )
        super().keyReleaseEvent(e)

    # -------------------------------------------------
    # Preview sink (UI wiring)
    # -------------------------------------------------
    def set_preview_sink(self, preview: Optional[PreviewPane]) -> None:
        """
        PreviewPane を差し替えるための API。
        main_window 側から呼ばれることを前提とする。

        - preview=None の場合は切断
        - 既存キャプチャ処理との互換性維持
        """
        self.logger.debug(
            "set_preview_sink preview=%s",
            preview.__class__.__name__ if preview else None,
        )
        self.preview = preview

    # -------------------------------------------------
    # Capture
    # -------------------------------------------------
    def capture_now(self):
        self.logger.debug("=== RegionWindow.capture_now start ===")
        self.save_dir.mkdir(parents=True, exist_ok=True)

        self.setWindowOpacity(0.0)
        QtWidgets.QApplication.processEvents()
        time.sleep(0.06)

        # -----------------------------
        # region (logical px)
        # -----------------------------
        reg_local = self._region_rect()
        top_left = self.mapToGlobal(reg_local.topLeft())

        # -----------------------------
        # screen & DPR (per-monitor)
        # -----------------------------
        screen = QtGui.QGuiApplication.screenAt(top_left)
        dpr = screen.devicePixelRatio() if screen else 1.0
        self.logger.debug("capture screen=%s dpr=%.4f", screen.name() if screen else None, dpr)

        # -----------------------------
        # logical px -> physical px
        # -----------------------------
        def lp(v: int) -> int:
            return int(round(v * dpr))

        region_physical = CaptureRegionLocal(
            lp(reg_local.x()),
            lp(reg_local.y()),
            lp(reg_local.width()),
            lp(reg_local.height()),
        )

        global_physical = CaptureGlobalTopLeft(
            lp(top_left.x()),
            lp(top_left.y()),
        )

        annos_physical = []
        for i, a in enumerate(self.ann.annos):
            annos_physical.append(
                CaptureAnnoRect(
                    lp(a.x),
                    lp(a.y),
                    lp(a.w),
                    lp(a.h),
                    a.color,
                    a.stroke,
                )
            )
            self.logger.debug(
                "[anno %d] logical=(%d,%d,%d,%d) -> physical=(%d,%d,%d,%d)",
                i, a.x, a.y, a.w, a.h,
                lp(a.x), lp(a.y), lp(a.w), lp(a.h),
            )

        # -----------------------------
        # build request (physical px only)
        # -----------------------------
        req = CaptureRequest(
            save_dir=self.save_dir,
            region_local=region_physical,
            global_top_left=global_physical,
            device_pixel_ratio=1.0,  # ★ もはや使わない
            annos=annos_physical,
            version=3,
        )

        result = self.capture_service.capture(req)
        self.setWindowOpacity(1.0)

        if not result.ok:
            self._toast_msg(f"Capture failed: {result.message}", 2.0)
            return

        # --- realtime UI update ---
        if self.preview:
            self.preview.add_capture(result.json_path)
            self.preview.show()
            self.preview.raise_()

        if result.png_path:
            self._toast_msg(f"Saved: {result.png_path.name}")

        self.logger.debug("=== RegionWindow.capture_now end ===")

    # -------------------------------------------------
    # Commands / UI helpers
    # -------------------------------------------------
    def _add_rect_from_preset(self):
        rr = self._region_rect()
        self.ann.add(
            rr.x() + 16,
            rr.y() + 16,
            160,
            90,
            color=self._new_rect_color.name(),
            stroke=self._new_rect_stroke,
        )
        self._toast_msg("+Rect")

    def _pick_new_color(self):
        col = QtWidgets.QColorDialog.getColor(self._new_rect_color, self)
        if col.isValid():
            self._set_new_color(col)

    def _set_new_color(self, col: QtGui.QColor):
        self._new_rect_color = col
        self.toolbar.update_swatch(col)

        try:
            self.conf.toolbar["rect_color"] = col.name()
            self.conf.save()
            self.logger.debug("config save toolbar.rect_color=%s", col.name())
        except Exception as e:
            self.logger.exception("save toolbar.rect_color failed: %s", e)

    def _set_new_stroke(self, stroke: int):
        self._new_rect_stroke = max(1, int(stroke))

        try:
            self.conf.toolbar["rect_stroke"] = int(self._new_rect_stroke)
            self.conf.save()
            self.logger.debug("config save toolbar.rect_stroke=%d", int(self._new_rect_stroke))
        except Exception as e:
            self.logger.exception("save toolbar.rect_stroke failed: %s", e)

    def _show_hotkeys_dialog(self):
        from core.config import DEFAULT_KEYS
        msg = "\n".join(
            f"{k}: {self.conf.hotkeys.get(k, DEFAULT_KEYS[k])}"
            for k in DEFAULT_KEYS
        )
        QtWidgets.QMessageBox.information(self, "Hotkeys", msg)

    def _open_settings(self, DialogCls):
        dlg = DialogCls(self.conf.hotkeys, self)
        if dlg.exec() == QtWidgets.QDialog.Accepted:
            self.conf.hotkeys.update(dlg.result_keys())
            self.conf.save()
            self.hotkeys.apply()

    def _close_region(self):
        self.close()

    def _start_drag(self, grab_mouse: bool = False):
        self._dragging = True

        # キーイベントを確保
        self.activateWindow()
        self.setFocus(QtCore.Qt.ActiveWindowFocusReason)

        if grab_mouse:
            self.grabMouse()

    def _toast_msg(self, msg: str, sec: Optional[float] = None):
        if sec is None:
            sec = float(self._toast_default_sec)
        self._toast = msg
        self._toast_until = time.time() + float(sec)
        self.update()

    # -------------------------------------------------
    # Record / Play
    # -------------------------------------------------
    def _toggle_record(self):
        if self._recording:
            out = self.recorder.stop()
            self._recording = False
            self.toolbar.setRecording(False)
            if out:
                self._toast_msg(f"REC saved: {Path(out).name}")
            return

        ts = time.strftime("%Y%m%d_%H%M%S")
        path, _ = QtWidgets.QFileDialog.getSaveFileName(
            self,
            "Save record",
            str(self.records_dir / f"rec_{ts}.jsonl"),
            "Record (*.jsonl)",
        )
        if not path:
            return

        # save last_dir
        try:
            p = Path(path)
            self.records_dir = p.parent
            self.records_dir.mkdir(parents=True, exist_ok=True)
            if hasattr(self.conf, "record") and isinstance(self.conf.record, dict):
                self.conf.record["last_dir"] = str(self.records_dir)
                self.conf.save()
                self.logger.debug("config save record.last_dir=%s", str(self.records_dir))
        except Exception as e:
            self.logger.exception("save record.last_dir failed: %s", e)

        self.recorder.start(Path(path), self.frameGeometry())
        self._recording = True
        self.toolbar.setRecording(True)
        self.activateWindow()
        self.raise_()
        self.setFocus()
        self._toast_msg("REC start")

    def _toggle_play(self):
        if self._playing:
            if self.player:
                self.player.stop()
            return

        path, _ = QtWidgets.QFileDialog.getOpenFileName(
            self,
            "Open record",
            str(self.records_dir),
            "Record (*.jsonl)",
        )
        if not path:
            return

        # save last_dir
        try:
            p = Path(path)
            self.records_dir = p.parent
            self.records_dir.mkdir(parents=True, exist_ok=True)
            if hasattr(self.conf, "record") and isinstance(self.conf.record, dict):
                self.conf.record["last_dir"] = str(self.records_dir)
                self.conf.save()
                self.logger.debug("config save record.last_dir=%s", str(self.records_dir))
        except Exception as e:
            self.logger.exception("save record.last_dir failed: %s", e)

        self.player = InputPlayer(self)
        self.player.load(Path(path))
        self.player.finished.connect(self._on_play_finished)
        self._playing = True
        self.toolbar.setPlaying(True)
        self.player.start()
        self._toast_msg("PLAY")

    def _on_play_finished(self):
        self._playing = False
        self.toolbar.setPlaying(False)
        self._toast_msg("PLAY done")
