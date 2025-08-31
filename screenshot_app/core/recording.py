# core/recording.py
from __future__ import annotations
from dataclasses import dataclass, asdict
from pathlib import Path
from typing import Optional, List, Dict, Any
import json, time, logging

from PySide6 import QtCore, QtGui, QtWidgets

log = logging.getLogger("recorder")

@dataclass
class RecEvent:
    t: float                 # 相対秒（record開始からの秒）
    kind: str                # "meta" | "mouse" | "key"
    etype: str               # meta: "start" / mouse: "press|move|release|wheel" / key: "keyPress|keyRelease"
    ts: float = 0.0          # 絶対時刻（epoch秒）
    # マウス
    pos_l: Optional[List[int]] = None  # ローカル [x,y]
    pos_g: Optional[List[int]] = None  # グローバル [x,y]
    button: Optional[int] = None       # Qt.MouseButton
    buttons: Optional[int] = None      # Qt.MouseButtons
    delta: Optional[int] = None        # wheel delta
    # キー
    key: Optional[int] = None          # Qt.Key
    mods: Optional[int] = None         # Qt.KeyboardModifiers
    text: Optional[str] = None
    # 追加情報
    extra: Optional[Dict[str, Any]] = None

class InputRecorder(QtCore.QObject):
    """ RegionWindow から通知を受けて入力イベントを記録 """
    def __init__(self, parent=None):
        super().__init__(parent)
        self._base_t: Optional[float] = None
        self._events: List[RecEvent] = []
        self._active: bool = False
        self._out_path: Optional[Path] = None

    def is_active(self) -> bool: return self._active

    def start(self, out_path: Path, base_win: QtCore.QRect):
        self._events.clear()
        self._base_t = time.perf_counter()
        self._active = True
        self._out_path = Path(out_path)
        # メタ（開始時ウィンドウ矩形）
        self._events.append(RecEvent(
            t=0.0, ts=time.time(), kind="meta", etype="start",
            extra={"base_win": [base_win.x(), base_win.y(), base_win.width(), base_win.height()]}
        ))
        log.info(f"[REC] started -> {self._out_path} base_win={base_win}")

    def stop(self) -> Optional[Path]:
        self._active = False
        if not self._out_path:
            return None
        p = self._out_path
        with p.open("w", encoding="utf-8") as f:
            for ev in self._events:
                f.write(json.dumps(asdict(ev), ensure_ascii=False) + "\n")
        log.info(f"[REC] saved {len(self._events)} events -> {p}")
        return p

    def _now(self) -> float:
        if self._base_t is None: return 0.0
        return time.perf_counter() - self._base_t

    # ---- call from RegionWindow on each input ----
    def on_mouse(self, etype: str, pos_local: QtCore.QPoint, pos_global: QtCore.QPoint,
                 buttons: int, button: int = 0, delta: int = 0):
        if not self._active: return
        self._events.append(RecEvent(
            t=self._now(), ts=time.time(), kind="mouse", etype=etype,
            pos_l=[pos_local.x(), pos_local.y()],
            pos_g=[pos_global.x(), pos_global.y()],
            button=button, buttons=buttons, delta=delta
        ))

    def on_key(self, etype: str, key: int, mods: int, text: str = ""):
        if not self._active: return
        self._events.append(RecEvent(
            t=self._now(), ts=time.time(), kind="key", etype=etype,
            key=key, mods=mods, text=text
        ))

class InputPlayer(QtCore.QObject):
    """ NDJSON を読み込み、RegionWindow に対して復元→postEvent で再生 """
    progress = QtCore.Signal(float)
    finished = QtCore.Signal()

    def __init__(self, target_widget: QtWidgets.QWidget, parent=None):
        super().__init__(parent)
        self._target = target_widget
        self._events: List[RecEvent] = []
        self._timer = QtCore.QTimer(self)
        self._timer.setSingleShot(True)
        self._timer.timeout.connect(self._step)
        self._i = 0
        self._t0 = 0.0
        self._base_win: Optional[QtCore.QRect] = None

    def load(self, path: Path):
        self._events.clear()
        self._base_win = None
        for line in Path(path).read_text(encoding="utf-8").splitlines():
            if not line.strip(): continue
            d = json.loads(line)
            ev = RecEvent(**d)
            # メタ
            if ev.kind == "meta" and ev.etype == "start":
                bw = (ev.extra or {}).get("base_win")
                if isinstance(bw, list) and len(bw) == 4:
                    self._base_win = QtCore.QRect(bw[0], bw[1], bw[2], bw[3])
            self._events.append(ev)
        self._events.sort(key=lambda e: e.t)
        self._i = 0

    def start(self):
        # 再生前にベース位置へ復元（存在すれば）
        if self._base_win is not None:
            self._target.setGeometry(self._base_win)
            QtWidgets.QApplication.processEvents()
        if not self._events:
            self.finished.emit(); return
        self._t0 = time.perf_counter()
        self._schedule_next()

    def _schedule_next(self):
        if self._i >= len(self._events):
            self.finished.emit(); return
        now = time.perf_counter() - self._t0
        delay = max(0.0, self._events[self._i].t - now)
        self._timer.start(int(delay * 1000))

    def _step(self):
        if self._i >= len(self._events):
            self.finished.emit(); return
        ev = self._events[self._i]
        self._i += 1
        self._dispatch(ev)
        self.progress.emit(ev.t)
        self._schedule_next()

    def _dispatch(self, ev: RecEvent):
        if ev.kind == "mouse":
            self._dispatch_mouse(ev)
        elif ev.kind == "key":
            self._dispatch_key(ev)

    def _dispatch_mouse(self, ev: RecEvent):
        # まずローカルを優先。無ければグローバル→ローカルへ変換。
        if ev.pos_l:
            pos = QtCore.QPoint(ev.pos_l[0], ev.pos_l[1])
            gpos = self._target.mapToGlobal(pos)
        elif ev.pos_g:
            gpos = QtCore.QPoint(ev.pos_g[0], ev.pos_g[1])
            pos = self._target.mapFromGlobal(gpos)
        else:
            pos = QtCore.QPoint(0, 0); gpos = self._target.mapToGlobal(pos)

        if ev.etype == "move":
            me = QtGui.QMouseEvent(QtCore.QEvent.MouseMove, pos, gpos,
                                   QtCore.Qt.MouseButton.NoButton,
                                   QtCore.Qt.MouseButtons(ev.buttons or 0),
                                   QtCore.Qt.KeyboardModifiers(0))
            QtGui.QGuiApplication.postEvent(self._target, me)
        elif ev.etype == "press":
            me = QtGui.QMouseEvent(QtCore.QEvent.MouseButtonPress, pos, gpos,
                                   QtCore.Qt.MouseButton(ev.button or 0),
                                   QtCore.Qt.MouseButtons(ev.buttons or 0),
                                   QtCore.Qt.KeyboardModifiers(0))
            QtGui.QGuiApplication.postEvent(self._target, me)
        elif ev.etype == "release":
            me = QtGui.QMouseEvent(QtCore.QEvent.MouseButtonRelease, pos, gpos,
                                   QtCore.Qt.MouseButton(ev.button or 0),
                                   QtCore.Qt.MouseButtons(ev.buttons or 0),
                                   QtCore.Qt.KeyboardModifiers(0))
            QtGui.QGuiApplication.postEvent(self._target, me)
        elif ev.etype == "wheel":
            we = QtGui.QWheelEvent(pos, gpos, QtCore.QPoint(0, ev.delta or 0),
                                   QtCore.QPoint(0, ev.delta or 0),
                                   int(ev.delta or 0), QtCore.Qt.Vertical,
                                   QtCore.Qt.MouseButtons(ev.buttons or 0),
                                   QtCore.Qt.KeyboardModifiers(0))
            QtGui.QGuiApplication.postEvent(self._target, we)

    def _dispatch_key(self, ev: RecEvent):
        kp = QtGui.QKeyEvent(QtCore.QEvent.KeyPress, ev.key or 0, QtCore.Qt.KeyboardModifiers(ev.mods or 0), ev.text or "")
        kr = QtGui.QKeyEvent(QtCore.QEvent.KeyRelease, ev.key or 0, QtCore.Qt.KeyboardModifiers(ev.mods or 0), ev.text or "")
        QtGui.QGuiApplication.postEvent(self._target, kp)
        QtGui.QGuiApplication.postEvent(self._target, kr)
