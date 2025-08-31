from __future__ import annotations
from dataclasses import dataclass
from pathlib import Path
import json, time
from typing import List, Optional, Tuple

from PySide6 import QtCore, QtGui, QtWidgets
from mss import mss, tools

from core.config import Config, HotkeyManager, load_last_state, save_last_state
from core.recording import InputRecorder, InputPlayer
from ui.toolbar import MiniToolbar
from ui.preview import PreviewPane
from ui.settings_dialog import SettingsDialog

# ── 見た目・動作定数
HANDLE = 10
FRAME_PEN_WIDTH = 4.0
FRAME_PEN_COLOR = "#FF3B30"          # 領域枠 = 赤固定
DEFAULT_NEW_RECT_COLOR = "#FF3B30"   # 新規矩形 = 赤
DEFAULT_NEW_RECT_STROKE = 2
RECT_CLOSE_SIZE = 16

# 領域外に置く UI のためのウィンドウ内マージン
TOP_MARGIN    = 36   # 上余白（×ボタン置き場）
BOTTOM_MARGIN = 60   # 下余白（ツールバー置き場）
SIDE_MARGIN   = 12   # 左右余白（見栄え用）

# 領域最小
MIN_REGION_W, MIN_REGION_H = 120, 90
MIN_WINDOW_W = MIN_REGION_W + 2*SIDE_MARGIN
MIN_WINDOW_H = MIN_REGION_H + TOP_MARGIN + BOTTOM_MARGIN

@dataclass
class AnnoRect:
    x:int; y:int; w:int; h:int
    color:str
    stroke:int

def handle_rects(rect: QtCore.QRect) -> dict:
    hs = HANDLE; cx = rect.x() + rect.width() // 2; cy = rect.y() + rect.height() // 2
    return {
        "nw": QtCore.QRect(rect.x()-hs//2, rect.y()-hs//2, hs, hs),
        "n" : QtCore.QRect(cx-hs//2, rect.y()-hs//2, hs, hs),
        "ne": QtCore.QRect(rect.right()-hs//2, rect.y()-hs//2, hs, hs),
        "e" : QtCore.QRect(rect.right()-hs//2, cy-hs//2, hs, hs),
        "se": QtCore.QRect(rect.right()-hs//2, rect.bottom()-hs//2, hs, hs),
        "s" : QtCore.QRect(cx-hs//2, rect.bottom()-hs//2, hs, hs),
        "sw": QtCore.QRect(rect.x()-hs//2, rect.bottom()-hs//2, hs, hs),
        "w" : QtCore.QRect(rect.x()-hs//2, cy-hs//2, hs, hs),
    }

def clamp_inside(r: QtCore.QRect, bounds: QtCore.QRect) -> QtCore.QRect:
    if r.width() < 12: r.setWidth(12)
    if r.height() < 12: r.setHeight(12)
    if r.left()   < bounds.left():   r.moveLeft(bounds.left())
    if r.top()    < bounds.top():    r.moveTop(bounds.top())
    if r.right()  > bounds.right():  r.moveRight(bounds.right())
    if r.bottom() > bounds.bottom(): r.moveBottom(bounds.bottom())
    return r

def rect_close_rect(r: QtCore.QRect) -> QtCore.QRect:
    return QtCore.QRect(r.right()-RECT_CLOSE_SIZE+1, r.top(), RECT_CLOSE_SIZE, RECT_CLOSE_SIZE)

class AnnotationManager(QtCore.QObject):
    changed = QtCore.Signal()
    def __init__(self, parent=None):
        super().__init__(parent)
        self.annos: List[AnnoRect] = []
        self.selected: Optional[int] = None
    def add(self, x=16, y=16, w=160, h=90, *, color: str, stroke:int):
        self.annos.append(AnnoRect(x, y, w, h, color, stroke))
        self.selected = len(self.annos)-1
        self.changed.emit()
    def remove_selected(self):
        if self.selected is not None and 0 <= self.selected < len(self.annos):
            del self.annos[self.selected]; self.selected = None; self.changed.emit()
    def remove_at(self, idx:int):
        if 0 <= idx < len(self.annos):
            del self.annos[idx]
            if self.selected == idx: self.selected = None
            elif self.selected is not None and self.selected > idx: self.selected -= 1
            self.changed.emit()
    def qrect(self, idx:int) -> QtCore.QRect:
        a = self.annos[idx]; return QtCore.QRect(a.x,a.y,a.w,a.h)
    def hit_handle(self, pos: QtCore.QPoint):
        for i in reversed(range(len(self.annos))):
            r = self.qrect(i)
            for k,hr in handle_rects(r).items():
                if hr.contains(pos): return i, k
        return None, None
    def hit_body(self, pos: QtCore.QPoint) -> Optional[int]:
        for i in reversed(range(len(self.annos))):
            if self.qrect(i).contains(pos): return i
        return None
    def hit_body_expanded(self, pos: QtCore.QPoint, pad:int=6) -> Optional[int]:
        for i in reversed(range(len(self.annos))):
            r = self.qrect(i).adjusted(-pad, -pad, pad, pad)
            if r.contains(pos): return i
        return None
    def hit_close(self, pos: QtCore.QPoint) -> Optional[int]:
        for i in reversed(range(len(self.annos))):
            if rect_close_rect(self.qrect(i)).contains(pos): return i
        return None
    def move_to(self, idx:int, rect: QtCore.QRect):
        a = self.annos[idx]
        a.x,a.y,a.w,a.h = rect.x(), rect.y(), rect.width(), rect.height()
        self.changed.emit()

class RegionWindow(QtWidgets.QWidget):
    def __init__(self, preview: Optional[PreviewPane] = None, save_dir: Optional[Path] = None):
        super().__init__()
        self.setWindowFlags(QtCore.Qt.FramelessWindowHint | QtCore.Qt.WindowStaysOnTopHint | QtCore.Qt.Tool)
        self.setAttribute(QtCore.Qt.WA_TranslucentBackground, True)
        self.setAttribute(QtCore.Qt.WA_NoSystemBackground, True)
        self.setMouseTracking(True)

        self.save_dir: Path = Path(save_dir) if save_dir else Path(__file__).resolve().parents[1] / "out"

        scr = QtGui.QGuiApplication.primaryScreen().availableGeometry()
        w, h = min(1000, scr.width()-200), min(560, scr.height()-200)
        self.resize(w, h); self.move(scr.center() - QtCore.QPoint(self.width()//2, self.height()//2))

        self.conf = Config(); self.conf.load()

        self.ann = AnnotationManager(self); self.ann.changed.connect(self.update)
        self._new_rect_color = QtGui.QColor(DEFAULT_NEW_RECT_COLOR)
        self._new_rect_stroke = int(DEFAULT_NEW_RECT_STROKE)

        # ドラッグ状態
        self.drag_mode: Optional[str] = None
        self.resize_handle: Optional[str] = None
        self.drag_start_local = QtCore.QPoint()
        self.drag_start_global = QtCore.QPoint()
        self.orig_window = QtCore.QRect()
        self.orig_region = QtCore.QRect()
        self._dragging = False

        # プレビュー
        self.preview = preview if preview is not None else None

        # ==== 子ウィジェット（同一ウィンドウ、領域“外”に配置）====
        self.toolbar = MiniToolbar(self)
        self.toolbar.shotClicked.connect(self.capture_now)
        self.toolbar.rectClicked.connect(self._add_rect_from_preset)
        self.toolbar.newColorPicked.connect(self._set_new_color)
        self.toolbar.newStrokePicked.connect(self._set_new_stroke)
        self.toolbar.settingsClicked.connect(lambda: self._open_settings(SettingsDialog))
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

        self._toast = ""; self._toast_until = 0

        # 前回状態の復元
        self._restore_last_state()
        self._place_children()

        # ホットキー
        self.hotkeys = HotkeyManager(self, {
            "capture": self.capture_now,
            "add_rect": self._add_rect_from_preset,
            "remove_selected": self.ann.remove_selected,
            "pick_new_color": self._pick_new_color,
            "show_hotkeys": self._show_hotkeys_dialog,
            "exit_app": self._close_region,
        }, self.conf)
        self.hotkeys.apply()

        self.recorder = InputRecorder(self)
        self.player = InputPlayer(self)
        self.player_running = False
        self.player.finished.connect(self._on_play_finished)

        self.toolbar.recToggleClicked.connect(self._rec_toggle)
        self.toolbar.playToggleClicked.connect(self._play_toggle)

    # ── 領域（赤枠）矩形：ウィンドウ内の実キャプチャ領域
    def _region_rect(self) -> QtCore.QRect:
        return QtCore.QRect(
            SIDE_MARGIN,
            TOP_MARGIN,
            max(1, self.width()  - 2*SIDE_MARGIN),
            max(1, self.height() - TOP_MARGIN - BOTTOM_MARGIN)
        )

    def set_save_dir(self, p: Path):
        self.save_dir = Path(p)
        self._toast_msg(f"Save dir: {self.save_dir}", 1.0)

    def set_preview_sink(self, preview: Optional[PreviewPane]):
        self.preview = preview

    # ---- 状態復元/保存 ----
    def _restore_last_state(self):
        data = load_last_state()
        rg = data.get("region")
        if isinstance(rg, dict):
            left = int(rg.get("left", self.x()));  top = int(rg.get("top", self.y()))
            width = max(200, int(rg.get("width", self.width())))
            height = max(180, int(rg.get("height", self.height())))
            self.setGeometry(QtCore.QRect(left, top, width, height))

        rects = data.get("rects", [])
        if isinstance(rects, list):
            self.ann.annos.clear()
            for r in rects:
                try:
                    x = int(r.get("x", 16)); y = int(r.get("y", 16))
                    w = int(r.get("w", 160)); h = int(r.get("h", 90))
                    color = str(r.get("color", DEFAULT_NEW_RECT_COLOR))
                    stroke = int(r.get("stroke", DEFAULT_NEW_RECT_STROKE))
                    qr = clamp_inside(QtCore.QRect(x,y,w,h), self._region_rect())
                    self.ann.add(qr.x(), qr.y(), qr.width(), qr.height(), color=color, stroke=stroke)
                except Exception:
                    continue
            self.ann.selected = None
            self.update()

    def _save_last_state(self):
        g = self.frameGeometry()
        data = {
            "region": {"left": g.left(), "top": g.top(), "width": g.width(), "height": g.height()},
            "rects": [
                {"x": a.x, "y": a.y, "w": a.w, "h": a.h, "color": a.color, "stroke": a.stroke}
                for a in self.ann.annos
            ],
        }
        save_last_state(data)

    # --- 子UIの配置（親の矩形に追従：領域の外側へ）
    def _place_children(self):
        self.toolbar.adjustSize()
        tb_w, tb_h = self.toolbar.width(), self.toolbar.height()
        tx = (self.width() - tb_w) // 2
        ty = self.height() - BOTTOM_MARGIN + (BOTTOM_MARGIN - tb_h)//2
        self.toolbar.move(max(0, tx), max(0, ty))

        cx = self.width() - SIDE_MARGIN - self.close_btn.width()
        cy = (TOP_MARGIN - self.close_btn.height()) // 2
        self.close_btn.move(max(0, cx), max(0, cy))

    # --- Qt events
    def resizeEvent(self, e: QtGui.QResizeEvent):
        super().resizeEvent(e)
        self._place_children()

    def closeEvent(self, e: QtGui.QCloseEvent):
        try:
            self._save_last_state()
        finally:
            return super().closeEvent(e)

    # --- frame / paint
    def _frame_rects(self) -> Tuple[QtCore.QRectF, QtCore.QRect]:
        reg = self._region_rect()
        inner_f = QtCore.QRectF(reg).adjusted(FRAME_PEN_WIDTH/2, FRAME_PEN_WIDTH/2,
                                              -FRAME_PEN_WIDTH/2, -FRAME_PEN_WIDTH/2)
        return inner_f, inner_f.toAlignedRect()

    def paintEvent(self, _e):
        p = QtGui.QPainter(self)
        p.setRenderHint(QtGui.QPainter.Antialiasing, not self._dragging)
        inner_f, frame = self._frame_rects()

        pen = QtGui.QPen(QtGui.QColor(FRAME_PEN_COLOR)); pen.setWidthF(FRAME_PEN_WIDTH); pen.setJoinStyle(QtCore.Qt.MiterJoin)
        p.setPen(pen); p.setBrush(QtCore.Qt.NoBrush); p.drawRect(inner_f)

        p.setPen(QtGui.QPen(QtGui.QColor(50,50,50,220), 1)); p.setBrush(QtGui.QColor(255,255,255,235))
        for hr in handle_rects(frame).values(): p.drawRect(hr)

        region_bounds = self._region_rect()
        for a in self.ann.annos:
            r = QtCore.QRect(a.x,a.y,a.w,a.h)
            col = QtGui.QColor(a.color)
            r = clamp_inside(r, region_bounds)
            p.setPen(QtGui.QPen(col, a.stroke if a.stroke>0 else 2))
            p.setBrush(QtCore.Qt.NoBrush)
            p.drawRect(r)
            rc = rect_close_rect(r); self._paint_close_icon(p, rc)
            p.setPen(QtGui.QPen(QtGui.QColor(50,50,50,220), 1)); p.setBrush(QtGui.QColor(255,255,255,235))
            for hrect in handle_rects(r).values(): p.drawRect(hrect)

        if self._toast and time.time() < self._toast_until:
            p.setOpacity(0.95)
            box = QtCore.QRect(10, self.height()-34-10, 280, 34)
            p.fillRect(box, QtGui.QColor(30,30,30,220))
            p.setPen(QtGui.QColor("white")); p.drawText(box.adjusted(8,0,0,0), QtCore.Qt.AlignVCenter, self._toast)

    def _paint_close_icon(self, p: QtGui.QPainter, rc: QtCore.QRect):
        p.setPen(QtGui.QPen(QtGui.QColor(180,180,180), 1))
        p.setBrush(QtGui.QColor(255,255,255))
        p.drawRect(rc)
        p.setPen(QtGui.QPen(QtGui.QColor(20,20,20), 2))
        p.drawLine(rc.left()+4, rc.top()+4, rc.right()-4, rc.bottom()-4)
        p.drawLine(rc.left()+4, rc.bottom()-4, rc.right()-4, rc.top()+4)

    # --- 入力
    def wheelEvent(self, e: QtGui.QWheelEvent): e.accept()

    def mousePressEvent(self, e: QtGui.QMouseEvent):
        self.recorder.on_mouse(
            "press",
            e.position().toPoint(),
            e.globalPosition().toPoint(),
            e.buttons().value,
            e.button().value
        )
        pos_local = e.position().toPoint()
        pos_global = e.globalPosition().toPoint()
        _, frame = self._frame_rects()
        reg = self._region_rect()

        if (idx_close := self.ann.hit_close(pos_local)) is not None:
            self.ann.remove_at(idx_close); return

        idx, h = self.ann.hit_handle(pos_local)
        if idx is not None and h:
            self.ann.selected = idx; self.drag_mode='resize_anno'; self.resize_handle=h
            self.drag_start_local = pos_local; self.orig_region = reg
            self.orig_window = self.frameGeometry()
            self.orig_rect = self.ann.qrect(idx)
            self._start_drag(); return

        idx = self.ann.hit_body(pos_local) or self.ann.hit_body_expanded(pos_local, pad=6)
        if idx is not None:
            self.ann.selected = idx; self.drag_mode='move_anno'
            self.drag_start_local = pos_local; self.orig_region = reg
            self.orig_window = self.frameGeometry()
            self.orig_rect = self.ann.qrect(idx)
            self._start_drag(); return

        for k,hr in handle_rects(frame).items():
            if hr.contains(pos_local):
                self.drag_mode='resize_win'; self.resize_handle=k
                self.drag_start_local = pos_local
                self.drag_start_global = pos_global
                self.orig_window = self.frameGeometry()
                self.orig_region = reg
                self._start_drag(grab_mouse=True); return

        if reg.contains(pos_local):
            self.drag_mode='move_win'
            self.drag_start_global = pos_global; self.orig_window=self.frameGeometry()
            self._start_drag(grab_mouse=True); return

        self.drag_mode=None; self.ann.selected=None; self.update()

    def mouseMoveEvent(self, e: QtGui.QMouseEvent):
        self.recorder.on_mouse(
            "move",
            e.position().toPoint(),
            e.globalPosition().toPoint(),
            e.buttons().value
        )
        pos_local  = e.position().toPoint()
        pos_global = e.globalPosition().toPoint()
        _, frame = self._frame_rects()
        reg = self._region_rect()

        cursor = QtCore.Qt.ArrowCursor
        if self.ann.hit_close(pos_local) is not None: cursor = QtCore.Qt.PointingHandCursor
        else:
            _, ah = self.ann.hit_handle(pos_local)
            if ah in ('n','s'): cursor = QtCore.Qt.SizeVerCursor
            elif ah in ('e','w'): cursor = QtCore.Qt.SizeHorCursor
            elif ah in ('ne','sw'): cursor = QtCore.Qt.SizeBDiagCursor
            elif ah in ('nw','se'): cursor = QtCore.Qt.SizeFDiagCursor
            else:
                if self.ann.hit_body(pos_local) is not None or reg.contains(pos_local): cursor = QtCore.Qt.SizeAllCursor
        self.setCursor(cursor)

        if not self.drag_mode:
            self.update(); return

        if self.drag_mode == 'move_win':
            dx = pos_global.x() - self.drag_start_global.x()
            dy = pos_global.y() - self.drag_start_global.y()
            self.move(self.orig_window.x() + dx, self.orig_window.y() + dy)
            return

        if self.drag_mode == 'resize_win':
            # 掴んだエッジ「だけ」を動かす。差分はグローバル座標で算出してブレ防止。
            dx = e.globalPosition().toPoint().x() - self.drag_start_global.x()
            dy = e.globalPosition().toPoint().y() - self.drag_start_global.y()

            ow = QtCore.QRect(self.orig_window)
            h  = self.resize_handle or ""

            nx, ny = ow.x(), ow.y()
            nr, nb = ow.right(), ow.bottom()

            if 'e' in h:
                nr = max(ow.left() + MIN_WINDOW_W - 1, ow.right() + dx)
            if 'w' in h:
                nx = min(ow.right() - (MIN_WINDOW_W - 1), ow.x() + dx)
            if 's' in h:
                nb = max(ow.top() + MIN_WINDOW_H - 1, ow.bottom() + dy)
            if 'n' in h:
                ny = min(ow.bottom() - (MIN_WINDOW_H - 1), ow.y() + dy)

            nw = max(MIN_WINDOW_W, nr - nx + 1)
            nh = max(MIN_WINDOW_H, nb - ny + 1)
            if 'w' in h: nx = nr - nw + 1
            if 'n' in h: ny = nb - nh + 1

            self.setGeometry(QtCore.QRect(nx, ny, nw, nh))
            return

        # 注釈（ローカル差分）
        dx = pos_local.x() - self.drag_start_local.x()
        dy = pos_local.y() - self.drag_start_local.y()
        bounds = self._region_rect()

        if self.drag_mode == 'move_anno' and self.ann.selected is not None:
            r0 = self.orig_rect
            r = QtCore.QRect(r0.x()+dx, r0.y()+dy, r0.width(), r0.height())
            r = clamp_inside(r, bounds)
            self.ann.move_to(self.ann.selected, r); return

        if self.drag_mode == 'resize_anno' and self.ann.selected is not None:
            r = QtCore.QRect(self.orig_rect); h = self.resize_handle
            if 'n' in h: r.setTop(r.top()+dy)
            if 's' in h: r.setBottom(r.bottom()+dy)
            if 'w' in h: r.setLeft(r.left()+dx)
            if 'e' in h: r.setRight(r.right()+dx)
            r = clamp_inside(r, bounds)
            self.ann.move_to(self.ann.selected, r); return

    def mouseReleaseEvent(self, e):
        self.recorder.on_mouse(
            "release",
            e.position().toPoint(),
            e.globalPosition().toPoint(),
            e.buttons().value,
            e.button().value
        )
        if self.mouseGrabber(): self.releaseMouse()
        self._dragging = False
        self.drag_mode=None; self.resize_handle=None
        self.update()

    def wheelEvent(self, e: QtGui.QWheelEvent):
        self.recorder.on_mouse(
            "wheel",
            e.position().toPoint(),
            e.globalPosition().toPoint(),
            e.buttons().value,
            0,                        # button = 0 （wheelなので押下ボタン無し）
            e.angleDelta().y()        # delta: 上下スクロール量
        )
        e.accept()

    def _mods_to_int(mods) -> int:
        # Qt.KeyboardModifier / Qt.KeyboardModifiers の両対応
        return 

    def keyPressEvent(self, e: QtGui.QKeyEvent):
        mods = e.modifiers().value if hasattr(e.modifiers(), "value") else int(e.modifiers())
        self.recorder.on_key("keyPress", e.key(), mods, e.text())
        super().keyPressEvent(e)

    def keyReleaseEvent(self, e: QtGui.QKeyEvent):
        mods = e.modifiers().value if hasattr(e.modifiers(), "value") else int(e.modifiers())
        self.recorder.on_key("keyRelease", e.key(), mods, e.text())
        super().keyReleaseEvent(e)

    # --- capture
    def capture_now(self):
        self.save_dir.mkdir(exist_ok=True, parents=True)
        self.setWindowOpacity(0.0); QtWidgets.QApplication.processEvents(); time.sleep(0.06)

        reg_local = self._region_rect()
        reg_global_top_left = self.mapToGlobal(reg_local.topLeft())
        gleft, gtop = reg_global_top_left.x(), reg_global_top_left.y()
        scr = QtGui.QGuiApplication.screenAt(reg_global_top_left)
        scale = scr.devicePixelRatio() if scr else 1.0
        bbox = {
            "left":   int(gleft * scale),
            "top":    int(gtop  * scale),
            "width":  int(reg_local.width()  * scale),
            "height": int(reg_local.height() * scale),
        }
        with mss() as sct:
            img = sct.grab(bbox)
            ts = time.strftime("%Y%m%d_%H%M%S")
            png_path = self.save_dir / f"capture_{ts}.png"
            tools.to_png(img.rgb, img.size, output=str(png_path))
        self.setWindowOpacity(1.0)

        meta = {
            "timestamp": ts,
            "region": {
                "left_global": gleft, "top_global": gtop,
                "width": reg_local.width(), "height": reg_local.height(),
                "device_pixel_ratio": scale
            },
            "rects": [ {"x":r.x, "y":r.y, "w":r.w, "h":r.h, "color":r.color, "stroke":r.stroke} for r in self.ann.annos ],
            "image_path": png_path.name,
            "comment": "",
            "version": 2
        }
        json_path = png_path.with_suffix(".json")
        json_path.write_text(json.dumps(meta, ensure_ascii=False, indent=2), encoding="utf-8")

        if self.preview:
            self.preview.add_capture(json_path)
            if not self.preview.isVisible():
                self.preview.show(); self.preview.raise_()

        self._toast_msg(f"Saved: {png_path.name}")

    # --- commands
    def _add_rect_from_preset(self):
        c = self._new_rect_color.name()
        s = int(self._new_rect_stroke)
        rr = self._region_rect()
        x = rr.x() + 16; y = rr.y() + 16
        self.ann.add(x=x, y=y, w=160, h=90, color=c, stroke=s)
        self._toast_msg(f"+Rect {c},{s}px", 0.8)

    def _pick_new_color(self):
        col = QtWidgets.QColorDialog.getColor(parent=self, title="Pick Color (for next Rect)")
        if col.isValid(): self._set_new_color(col)
    def _set_new_color(self, col: QtGui.QColor):
        self._new_rect_color = col; self.toolbar.update_swatch(col); self._toast_msg(f"New-Color: {col.name()}", 0.8)
    def _set_new_stroke(self, stroke:int):
        self._new_rect_stroke = max(1, int(stroke)); self._toast_msg(f"New-Width: {self._new_rect_stroke}px", 0.8)

    def _show_hotkeys_dialog(self):
        from core.config import DEFAULT_KEYS
        ms = "\n".join([f"{k}: {self.conf.hotkeys.get(k, DEFAULT_KEYS[k]) or '(disabled)'}" for k in DEFAULT_KEYS.keys()])
        QtWidgets.QMessageBox.information(self, "Hotkeys", ms)

    def _open_settings(self, DialogCls=SettingsDialog):
        dlg = DialogCls(self.conf.hotkeys, self)
        if dlg.exec() == QtWidgets.QDialog.Accepted:
            new_keys = dlg.result_keys()
            if new_keys:
                self.conf.hotkeys.update(new_keys)
                self.conf.save()
                self.hotkeys.apply()
                # ★ 追加：Toolbar のヒントも更新
                self.toolbar.update_hotkey_hints(self.conf.hotkeys)
                self._toast_msg("Settings saved", 1.0)

    def _close_region(self): self.close()

    def _start_drag(self, grab_mouse: bool=False):
        self._dragging = True
        if grab_mouse: self.grabMouse()

    def _toast_msg(self, msg:str, sec:float=1.2):
        self._toast = msg; self._toast_until = time.time()+sec; self.update()

    def _rec_file_path(self) -> Path:
        self.save_dir.mkdir(parents=True, exist_ok=True)
        ts = time.strftime("%Y%m%d_%H%M%S")
        return self.save_dir / f"record_{ts}.ndjson"

    def _rec_start(self):
        if self.recorder.is_active():
            self._toast_msg("Recording already active", 0.8); return
        path = self._rec_file_path()
        self.recorder.start(path, self.frameGeometry())
        self.toolbar.setRecording(True)        # ●を点灯
        self._toast_msg(f"REC ● {path.name}", 1.2)

    def _rec_stop(self):
        p = self.recorder.stop()
        self.toolbar.setRecording(False)       # ●を消灯
        self._toast_msg(f"REC ⏹ saved: {p.name if p else 'n/a'}", 1.2)

    def _rec_play(self):
        files = sorted(self.save_dir.glob("record_*.ndjson"))
        if not files:
            self._toast_msg("No record file", 1.0); return
        latest = files[-1]
        if self.player is None:
            self.player = InputPlayer(self)
            self.player.finished.connect(lambda: self._toast_msg("Play done", 0.8))
        self.player.load(latest)
        self._toast_msg(f"PLAY ▶ {latest.name}", 1.0)
        self.player.start()

    # ハンドラ群
    def _rec_toggle(self):
        if not self.recorder.is_active():
            self._rec_start()
        else:
            self._rec_stop()

    def _play_toggle(self):
        if not self.player_running:
            # 直近の record_*.ndjson を再生
            files = sorted(self.save_dir.glob("record_*.ndjson"))
            if not files:
                self._toast_msg("No record file", 1.0); return
            latest = files[-1]
            try:
                self.player.load(latest)
                self.player_running = True
                self.toolbar.setPlaying(True)     # ▶ → ■
                self._toast_msg(f"PLAY ▶ {latest.name}", 1.0)
                self.player.start()
            except Exception as ex:
                self.player_running = False
                self.toolbar.setPlaying(False)
                self._toast_msg(f"Play error: {ex}", 2.0)
        else:
            self._play_stop()

    def _play_stop(self):
        # InputPlayer は QTimer 駆動なので止める
        try:
            if hasattr(self.player, "_timer"):
                self.player._timer.stop()
            self.player._i = len(getattr(self.player, "_events", []))  # 消化済みにする
        except Exception:
            pass
        self.player_running = False
        self.toolbar.setPlaying(False)          # ■ → ▶
        self._toast_msg("Play stopped", 0.8)

    def _on_play_finished(self):
        self.player_running = False
        self.toolbar.setPlaying(False)          # ■ → ▶
        self._toast_msg("Play done", 0.8)
