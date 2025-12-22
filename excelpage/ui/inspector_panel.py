# ui/inspector_panel.py
# UI完成形固定・機能完全互換・修正容易化リファクタ版（説明なし）

from __future__ import annotations

import time
import traceback
from collections import deque
from typing import Optional, Any, Dict

from PySide6.QtWidgets import (
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QTextEdit,
    QFileDialog,
    QSizePolicy,
    QApplication,
)
from PySide6.QtCore import Qt, QEvent, QTimer
from PySide6.QtGui import QKeyEvent


# =================================================
# Logger
# =================================================
try:
    from Logger import Logger

    logger = Logger(
        name="Inspector",
        log_file_path="logs/app.log",
        level="DEBUG",
    )
except ModuleNotFoundError:
    import logging

    def get_logger(name: str):
        lg = logging.getLogger(name)
        if not lg.handlers:
            lg.setLevel(logging.DEBUG)
            h = logging.StreamHandler()
            fmt = logging.Formatter("%(asctime)s - %(name)s - %(levelname)s - %(message)s")
            h.setFormatter(fmt)
            lg.addHandler(h)
            lg.propagate = False
        return lg

    logger = get_logger("Inspector")


# =================================================
# MacroRecorder
# =================================================
try:
    from services.macro_recorder import get_macro_recorder
except ModuleNotFoundError:
    import json
    import datetime

    class _DummyMacroRecorder:
        def __init__(self):
            self._recording = False
            self._steps = []

        def is_recording(self) -> bool:
            return self._recording

        def start(self):
            self._recording = True

        def stop(self):
            self._recording = False

        def steps_count(self) -> int:
            return len(self._steps)

        def save_json(self, path: str):
            data = {
                "version": 1,
                "name": "dummy-macro",
                "created_at": datetime.datetime.now().isoformat(),
                "steps": self._steps,
            }
            with open(path, "w", encoding="utf-8") as f:
                json.dump(data, f, indent=2, ensure_ascii=False)

    _DUMMY_MACRO = _DummyMacroRecorder()

    def get_macro_recorder():
        return _DUMMY_MACRO


# =================================================
# Helpers
# =================================================
def now_ms() -> int:
    return int(time.time() * 1000)


def key_to_name(key: int) -> str:
    return {
        Qt.Key_Up: "Up",
        Qt.Key_Down: "Down",
        Qt.Key_Left: "Left",
        Qt.Key_Right: "Right",
        Qt.Key_F2: "F2",
        Qt.Key_Return: "Return",
        Qt.Key_Enter: "Enter",
        Qt.Key_Escape: "Escape",
        Qt.Key_A: "A",
        Qt.Key_C: "C",
        Qt.Key_X: "X",
        Qt.Key_V: "V",
        Qt.Key_Z: "Z",
        Qt.Key_Y: "Y",
        Qt.Key_R: "R",
        Qt.Key_S: "S",
        Qt.Key_F4: "F4",
    }.get(key, f"Key({key})")


def mod_to_str(mod: Qt.KeyboardModifiers) -> str:
    parts = []
    if mod & Qt.ControlModifier:
        parts.append("Ctrl")
    if mod & Qt.ShiftModifier:
        parts.append("Shift")
    if mod & Qt.AltModifier:
        parts.append("Alt")
    if mod & Qt.MetaModifier:
        parts.append("Meta")
    return "+".join(parts) if parts else "None"


# =================================================
# UI Log Buffer
# =================================================
class UILog:
    def __init__(self, view: QTextEdit, max_len: int = 10):
        self._buf = deque(maxlen=max_len)
        self._view = view

    def add(self, msg: str, color: str = "#ddd"):
        self._buf.appendleft(f'<span style="color:{color}">▸ {msg}</span>')
        self._view.setHtml("<br>".join(self._buf))


# =================================================
# Key Dispatcher
# =================================================
class KeyDispatcher:
    def __init__(self, panel: "InspectorPanel"):
        self.p = panel

    def handle(self, e: QKeyEvent, trace_id: int):
        key = e.key()
        mod = e.modifiers()

        if (mod & Qt.ControlModifier) and (mod & Qt.ShiftModifier):
            if key == Qt.Key_R:
                self.p.toggle_record()
                return
            if key == Qt.Key_S:
                self.p.save_macro()
                return

        if (mod & Qt.AltModifier) and key == Qt.Key_F4:
            self.p.close()
            return

        if key == Qt.Key_F2:
            self.p.enter_edit(trace_id)
            return

        if self.p.edit_mode:
            self.p.handle_edit_keys(e, trace_id)
            return

        if (mod & Qt.ControlModifier) and (mod & Qt.ShiftModifier):
            if key in self.p.ARROWS:
                self.p.exec_and_log("select_edge", trace_id, direction=self.p.dir(key),
                                    msg="Select edge (Ctrl+Shift+Arrow)", color="#7fd7ff")
                return

        if mod & Qt.ControlModifier:
            self.p.handle_ctrl(key, trace_id)
            return

        if mod & Qt.ShiftModifier:
            if key in self.p.ARROWS:
                self.p.exec_and_log("select_move", trace_id, direction=self.p.dir(key),
                                    msg="Select move (Shift+Arrow)", color="#7fd7ff")
                return

        if key in self.p.ARROWS:
            d = self.p.dir(key)
            self.p.exec_and_log("move_cell", trace_id, direction=d, step=1,
                                msg=f"Move {d} (Arrow)", color="#aaa")
            return


# =================================================
# InspectorPanel
# =================================================
class InspectorPanel(QWidget):
    MAX_LOG = 10
    POLL_MS = 400
    ARROWS = (Qt.Key_Up, Qt.Key_Down, Qt.Key_Left, Qt.Key_Right)

    def __init__(self):
        super().__init__(None)

        self._tree = None
        self._macro = get_macro_recorder()
        self._trace = 0
        self.edit_mode = False
        self._last_ctx: Optional[str] = None

        self.setWindowTitle("Excel Inspector")
        self.setWindowFlags(Qt.Window | Qt.WindowStaysOnTopHint)
        self.resize(620, 420)
        self.setStyleSheet("QWidget { background-color:#0f0f0f; }")
        self.setFocusPolicy(Qt.StrongFocus)

        root = QVBoxLayout(self)
        root.setContentsMargins(10, 10, 10, 10)
        root.setSpacing(8)

        self.addr_label = QLabel("—")
        self.addr_label.setAlignment(Qt.AlignCenter)
        self.addr_label.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        self.addr_label.setStyleSheet(
            "background:#1b1b1b;color:#7fd7ff;font-size:14px;font-weight:700;padding:6px;border-radius:6px;"
        )

        self.rec_label = QLabel("")
        self.rec_label.setFixedWidth(56)
        self.rec_label.setAlignment(Qt.AlignCenter)
        self.rec_label.setStyleSheet("color:#ff4d4d;font-weight:900;")

        self.hint_label = QLabel("Ctrl+Shift+R:REC   Ctrl+Shift+S:SAVE")
        self.hint_label.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        self.hint_label.setStyleSheet("color:#666;font-size:11px;")
        self.hint_label.setFixedWidth(230)

        header = QHBoxLayout()
        header.addWidget(self.addr_label, 1)
        header.addWidget(self.rec_label)
        header.addWidget(self.hint_label)
        root.addLayout(header)

        bar = QHBoxLayout()
        fx = QLabel("fx")
        fx.setFixedWidth(26)
        fx.setAlignment(Qt.AlignCenter)
        fx.setStyleSheet("color:#6cf;font-weight:700;")

        self.editor = QLineEdit()
        self.editor.setReadOnly(True)
        self.editor.setFocusPolicy(Qt.NoFocus)
        self.editor.setStyleSheet(
            "background:#151515;color:#eee;border:1px solid #2a2a2a;border-radius:6px;padding:8px;font-size:14px;"
        )

        bar.addWidget(fx)
        bar.addWidget(self.editor, 1)
        root.addLayout(bar)

        self.log_view = QTextEdit()
        self.log_view.setReadOnly(True)
        self.log_view.setFixedHeight(170)
        self.log_view.setStyleSheet("background:#101010;color:#ccc;border-radius:6px;padding:6px;")
        root.addWidget(self.log_view)

        self.ui_log = UILog(self.log_view, self.MAX_LOG)
        self.ui_log.add("Help: Ctrl+Shift+R = REC  /  Ctrl+Shift+S = SAVE", "#777")

        self.dispatcher = KeyDispatcher(self)

        self.installEventFilter(self)
        self.editor.installEventFilter(self)

        self._poll = QTimer(self)
        self._poll.timeout.connect(self.poll_context)
        self._poll.start(self.POLL_MS)

        # --- Key debounce state ---
        self.KEY_GUARD_MS = 120
        self._last_key_sig: Optional[tuple[int, int]] = None
        self._last_key_at_ms: int = 0

    # -------------------------------------------------
    # EventFilter
    # -------------------------------------------------
    def eventFilter(self, obj, event):
        if event.type() == QEvent.KeyPress:
            e: QKeyEvent = event  # type: ignore
            if e.isAutoRepeat():
                return True

            # ★ 編集中 + editor 由来はデバウンスしない
            if self.edit_mode and obj is self.editor:
                return False

            # --- 重複防止（Inspector 操作のみ） ---
            now = now_ms()
            sig = (int(e.key()), int(e.modifiers().value))
            if self._last_key_sig == sig and (now - self._last_key_at_ms) < self.KEY_GUARD_MS:
                logger.debug(
                    f"[KEY] debounce skip key={key_to_name(e.key())} "
                    f"mod={mod_to_str(e.modifiers())} dt={now - self._last_key_at_ms}ms"
                )
                return True

            self._last_key_sig = sig
            self._last_key_at_ms = now

            self._trace += 1
            logger.info(
                f"[KEY] trace={self._trace} key={key_to_name(e.key())} "
                f"mod={mod_to_str(e.modifiers())}"
            )

            self.dispatcher.handle(e, self._trace)
            return True

        return super().eventFilter(obj, event)

    # -------------------------------------------------
    # Actions
    # -------------------------------------------------
    def dir(self, key: int) -> str:
        return {
            Qt.Key_Up: "up",
            Qt.Key_Down: "down",
            Qt.Key_Left: "left",
            Qt.Key_Right: "right",
        }[key]

    def exec(self, op: str, trace_id: int, **kw):
        if self._tree:
            self._tree._engine_exec(op, source="inspector", **kw)

    def exec_and_log(self, op: str, trace_id: int, msg: str, color: str, **kw):
        self.exec(op, trace_id, **kw)
        self.ui_log.add(msg, color)

    def handle_ctrl(self, key: int, trace_id: int):
        mapping = {
            Qt.Key_A: ("select_all", "Select All (Ctrl+A)"),
            Qt.Key_C: ("copy", "Copy (Ctrl+C)"),
            Qt.Key_X: ("cut", "Cut (Ctrl+X)"),
            Qt.Key_V: ("paste", "Paste (Ctrl+V)"),
            Qt.Key_Z: ("undo", "Undo (Ctrl+Z)"),
            Qt.Key_Y: ("redo", "Redo (Ctrl+Y)"),
        }
        if key in mapping:
            op, msg = mapping[key]
            self.exec_and_log(op, trace_id, msg, "#6cf")
            return
        if key in self.ARROWS:
            self.exec_and_log("move_edge", trace_id,
                              "Move edge (Ctrl+Arrow)", "#aaa",
                              direction=self.dir(key))

    def enter_edit(self, trace_id: int):
        self.edit_mode = True
        self.editor.setReadOnly(False)
        self.editor.setFocusPolicy(Qt.StrongFocus)
        self.editor.setFocus()
        self.editor.selectAll()
        self.ui_log.add("Edit (F2)", "#7fd7ff")

    def handle_edit_keys(self, e: QKeyEvent, trace_id: int):
        if e.key() in (Qt.Key_Return, Qt.Key_Enter):
            val = self.editor.text()
            self.exec("set_cell_value", trace_id, cell="*", value=val)
            self.ui_log.add(f"Set = {val}", "#ffb347")
            self.editor.clear()
            self.exit_edit()
            return
        if e.key() == Qt.Key_Escape:
            self.editor.clear()
            self.ui_log.add("Edit cancel (Esc)", "#aaa")
            self.exit_edit()

    def exit_edit(self):
        self.edit_mode = False
        self.editor.setReadOnly(True)
        self.editor.setFocusPolicy(Qt.NoFocus)
        self.setFocus()

    def toggle_record(self):
        if not self._macro.is_recording():
            self._macro.start()
            self.rec_label.setText("● REC")
            self.ui_log.add("REC START (Ctrl+Shift+R)", "#ff4d4d")
        else:
            self._macro.stop()
            self.rec_label.setText("")
            self.ui_log.add("REC STOP (Ctrl+Shift+R)", "#ff4d4d")

    def save_macro(self):
        cnt = self._macro.steps_count()
        if cnt == 0:
            self.ui_log.add("No macro steps (Ctrl+Shift+S)", "#aaa")
            return
        path, _ = QFileDialog.getSaveFileName(self, "Save Macro", "", "Macro JSON (*.json)")
        if not path:
            self.ui_log.add("Save canceled (Ctrl+Shift+S)", "#777")
            return
        self._macro.save_json(path)
        self.ui_log.add(f"Saved macro ({cnt} steps) (Ctrl+Shift+S)", "#7fd7ff")

    # -------------------------------------------------
    # Poll
    # -------------------------------------------------
    def poll_context(self):
        if not self._tree:
            return
        try:
            ctx = self._tree._engine_exec("get_active_context")
        except Exception:
            return
        if not isinstance(ctx, dict):
            return
        addr = str(ctx.get("address", "")).replace("$", "")
        sheet = str(ctx.get("sheet", ""))
        label = f"{sheet}!{addr}" if sheet and addr else "—"
        if label != self._last_ctx:
            self._last_ctx = label
            self.addr_label.setText(label)

    # -------------------------------------------------
    # External bind
    # -------------------------------------------------
    def set_tree(self, tree):
        self._tree = tree
        logger.info(f"[Inspector] set_tree tree={tree}")

    def set_current_cell(self, cell: str):
        logger.info(f"[Inspector] set_current_cell cell={cell}")

    # -------------------------------------------------
    # Focus helpers
    # -------------------------------------------------
    def showEvent(self, event):
        super().showEvent(event)
        try:
            self.raise_()
            self.activateWindow()
            self.setFocus(Qt.ActiveWindowFocusReason)
        except Exception:
            pass

    def mousePressEvent(self, event):
        try:
            if not self.edit_mode:
                self.setFocus(Qt.ActiveWindowFocusReason)
        except Exception:
            pass
        super().mousePressEvent(event)

    # -------------------------------------------------
    # Validation helpers (dev only)
    # -------------------------------------------------
class _FakeTree:
    def __init__(self):
        self.calls = []

    def _engine_exec(self, op: str, **kw):
        self.calls.append((op, dict(kw)))
        logger.info(f"[VALIDATE] engine_exec op={op} kw={kw}")
        if op == "get_active_context":
            return {"address": "$A$1", "sheet": "Sheet1"}
        return None

    def clear(self):
        self.calls.clear()

    def last(self):
        if not self.calls:
            raise AssertionError("engine_exec が呼ばれていません")
        return self.calls[-1]


def _assert_last(tree: _FakeTree, op: str, **expect):
    got_op, got_kw = tree.last()
    assert got_op == op, f"op mismatch: {got_op} != {op}"
    for k, v in expect.items():
        assert got_kw.get(k) == v, f"{k} mismatch: {got_kw.get(k)} != {v}"


def validate_keys(panel: InspectorPanel, tree: _FakeTree):
    from PySide6.QtTest import QTest

    wait_ms = panel.KEY_GUARD_MS + 10

    tree.clear()
    QTest.keyClick(panel, Qt.Key_Down)
    QTest.qWait(wait_ms)
    _assert_last(tree, "move_cell", direction="down", step=1, source="inspector")

    tree.clear()
    QTest.keyClick(panel, Qt.Key_Right, Qt.ShiftModifier)
    QTest.qWait(wait_ms)
    _assert_last(tree, "select_move", direction="right", source="inspector")

    tree.clear()
    QTest.keyClick(panel, Qt.Key_Left, Qt.ControlModifier)
    QTest.qWait(wait_ms)
    _assert_last(tree, "move_edge", direction="left", source="inspector")

    tree.clear()
    QTest.keyClick(panel, Qt.Key_Up, Qt.ControlModifier | Qt.ShiftModifier)
    QTest.qWait(wait_ms)
    _assert_last(tree, "select_edge", direction="up", source="inspector")

def validate_edit(panel: InspectorPanel, tree: _FakeTree):
    from PySide6.QtTest import QTest

    wait_ms = panel.KEY_GUARD_MS + 10

    tree.clear()

    # F2
    QTest.keyClick(panel, Qt.Key_F2)
    QTest.qWait(wait_ms)

    # 入力（ここは editor が処理するのでデバウンス非対象）
    QTest.keyClicks(panel.editor, "abc")

    # ★ 人間は必ず Enter の前に一瞬止まる
    QTest.qWait(wait_ms)

    # Enter
    QTest.keyClick(panel.editor, Qt.Key_Return)
    QTest.qWait(wait_ms)

    _assert_last(
        tree,
        "set_cell_value",
        cell="*",
        value="abc",
        source="inspector",
    )

    assert panel.edit_mode is False
    assert panel.editor.isReadOnly() is True

def validate_autorepeat(panel: InspectorPanel):
    from PySide6.QtGui import QKeyEvent
    from PySide6.QtCore import QEvent

    tree: _FakeTree = panel._tree  # type: ignore
    tree.clear()

    ev = QKeyEvent(
        QEvent.KeyPress,
        Qt.Key_Down,
        Qt.NoModifier,
        "",
        True,  # autoRepeat
        1,
    )
    QApplication.sendEvent(panel, ev)

    # autoRepeat はデバウンス以前に無視される
    assert len(tree.calls) == 0


def validate_user_input(panel: InspectorPanel):
    from PySide6.QtTest import QTest

    logger.info("=== VALIDATE: user key input ===")
    logger.info("↓ 今から実キーを押してください ↓")
    logger.info("  Arrow / Shift+Arrow / Ctrl+Arrow / F2 → Enter")
    logger.info("  （10秒以内・デバウンス有効）")

    QTest.qWait(10000)

    logger.info("=== VALIDATE: user key input DONE ===")

if __name__ == "__main__":
    import sys
    from PySide6.QtWidgets import QApplication
    from PySide6.QtTest import QTest

    app = QApplication(sys.argv)

    panel = InspectorPanel()
    fake_tree = _FakeTree()
    panel.set_tree(fake_tree)
    panel._poll.stop()

    panel.show()
    QTest.qWaitForWindowExposed(panel)
    panel.activateWindow()
    panel.setFocus()

    validate_keys(panel, fake_tree)
    validate_edit(panel, fake_tree)
    validate_autorepeat(panel)
    validate_user_input(panel)

    logger.info("=== ALL VALIDATION PASSED ===")
