from __future__ import annotations

import time
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
        logger = logging.getLogger(name)
        if not logger.handlers:
            logger.setLevel(logging.DEBUG)
            h = logging.StreamHandler()
            fmt = logging.Formatter("%(asctime)s - %(name)s - %(levelname)s - %(message)s")
            h.setFormatter(fmt)
            logger.addHandler(h)
            logger.propagate = False
        return logger

# =================================================
# MacroRecorder import absorb (run-as-script safe)
# =================================================
try:
    from services.macro_recorder import get_macro_recorder
except ModuleNotFoundError:
    import json
    import datetime

    class _DummyMacroRecorder:
        """
        Fallback MacroRecorder for direct execution.
        - API compatible (minimum)
        - No external dependency
        """

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
        """
        Fallback getter.
        Same signature as services.macro_recorder.get_macro_recorder
        """
        return _DUMMY_MACRO


# =================================================
# Top-level helpers (no nested functions)
# =================================================
def now_ms() -> int:
    return int(time.time() * 1000)


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


def key_to_name(key: int) -> str:
    m = {
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
    }
    return m.get(key, f"Key({key})")


def obj_name(obj: Any) -> str:
    try:
        if obj is None:
            return "None"
        n = obj.objectName()
        if n:
            return n
        return obj.__class__.__name__
    except Exception:
        return str(type(obj))


def focus_snapshot() -> str:
    try:
        fw = QApplication.focusWidget()
        aw = QApplication.activeWindow()
        return f"focus={obj_name(fw)} activeWindow={obj_name(aw)}"
    except Exception:
        return "focus=(unknown)"


def safe_native_info(e: QKeyEvent) -> str:
    try:
        # Windowsなら nativeVirtualKey / nativeScanCode が見える
        return f"native_vk={e.nativeVirtualKey()} native_sc={e.nativeScanCode()} native_mod={e.nativeModifiers()}"
    except Exception:
        return "native=(n/a)"


def kv_compact(d: Dict[str, Any]) -> str:
    try:
        return " ".join([f"{k}={d[k]!r}" for k in d.keys()])
    except Exception:
        return str(d)


# =================================================
# InspectorPanel
# =================================================
class InspectorPanel(QWidget):
    """
    Excel Inspector (Excel-like, Keyboard-first)

    - F2 : edit cell
    - Enter : commit
    - Esc : cancel edit only
    - Ctrl / Shift / Arrow : Excel compatible
    - Ctrl+Shift+R : Macro record start/stop
    - Ctrl+Shift+S : Macro save
    - Alt+F4 : close window

    ★ Debug tracing:
      - KeyPress/KeyRelease に trace_id を採番
      - _exec に trace_id を付与して TreeView/ExcelWorker と突合可能にする
    """

    MAX_LOG = 10
    POLL_MS = 400

    def __init__(self):
        super().__init__(None)

        self._tree = None
        self._macro = get_macro_recorder()

        self._log_buf = deque(maxlen=self.MAX_LOG)
        self._edit_mode = False
        self._last_ctx: Optional[str] = None
        self._active_cell: Optional[str] = None

        # ---- tracing
        self._trace_seq: int = 0
        self._last_exec_trace: Optional[int] = None
        self._last_exec_at_ms: int = 0

        # ---- window
        self.setWindowTitle("Excel Inspector")
        self.setWindowFlags(Qt.Window | Qt.WindowStaysOnTopHint)
        self.resize(620, 420)
        self.setStyleSheet("QWidget { background-color:#0f0f0f; }")

        # ★ Inspector 自体は常にキーを取れる状態にする
        self.setFocusPolicy(Qt.StrongFocus)

        root = QVBoxLayout(self)
        root.setContentsMargins(10, 10, 10, 10)
        root.setSpacing(8)

        # =================================================
        # Header (Address + REC slot + hint)
        # =================================================
        header = QHBoxLayout()
        header.setSpacing(8)

        self.addr_label = QLabel("—")
        self.addr_label.setAlignment(Qt.AlignCenter)
        self.addr_label.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        self.addr_label.setStyleSheet(
            """
            QLabel {
                background:#1b1b1b;
                color:#7fd7ff;
                font-size:14px;
                font-weight:700;
                padding:6px;
                border-radius:6px;
            }
            """
        )
        self.addr_label.setFocusPolicy(Qt.NoFocus)
        self.addr_label.setTextInteractionFlags(Qt.NoTextInteraction)

        self.rec_label = QLabel("")
        self.rec_label.setFixedWidth(56)
        self.rec_label.setAlignment(Qt.AlignCenter)
        self.rec_label.setStyleSheet(
            """
            QLabel {
                color:#ff4d4d;
                font-weight:900;
            }
            """
        )
        self.rec_label.setFocusPolicy(Qt.NoFocus)
        self.rec_label.setTextInteractionFlags(Qt.NoTextInteraction)

        self.hint_label = QLabel("Ctrl+Shift+R:REC   Ctrl+Shift+S:SAVE")
        self.hint_label.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        self.hint_label.setStyleSheet(
            """
            QLabel {
                color:#666;
                font-size:11px;
                padding-right:4px;
            }
            """
        )
        self.hint_label.setFixedWidth(230)
        self.hint_label.setFocusPolicy(Qt.NoFocus)
        self.hint_label.setTextInteractionFlags(Qt.NoTextInteraction)

        header.addWidget(self.addr_label, 1)
        header.addWidget(self.rec_label, 0)
        header.addWidget(self.hint_label, 0)
        root.addLayout(header)

        # =================================================
        # Formula bar
        # =================================================
        bar = QHBoxLayout()
        bar.setSpacing(8)

        fx = QLabel("fx")
        fx.setFixedWidth(26)
        fx.setAlignment(Qt.AlignCenter)
        fx.setStyleSheet("color:#6cf; font-weight:700;")
        fx.setFocusPolicy(Qt.NoFocus)
        fx.setTextInteractionFlags(Qt.NoTextInteraction)

        self.editor = QLineEdit()
        self.editor.setPlaceholderText("F2 to edit")
        self.editor.setStyleSheet(
            """
            QLineEdit {
                background:#151515;
                color:#eee;
                border:1px solid #2a2a2a;
                border-radius:6px;
                padding:8px;
                font-size:14px;
            }
            QLineEdit:focus { border:1px solid #6cf; }
            """
        )
        self.editor.setFocusPolicy(Qt.NoFocus)
        self.editor.setReadOnly(True)

        bar.addWidget(fx)
        bar.addWidget(self.editor, 1)
        root.addLayout(bar)

        # =================================================
        # Log
        # =================================================
        self.log = QTextEdit()
        self.log.setReadOnly(True)
        self.log.setFixedHeight(170)
        self.log.setStyleSheet(
            """
            QTextEdit {
                background:#101010;
                color:#ccc;
                border-radius:6px;
                padding:6px;
            }
            """
        )
        self.log.setFocusPolicy(Qt.NoFocus)
        self.log.setTextInteractionFlags(Qt.NoTextInteraction)

        root.addWidget(self.log)

        self._log_add("Help: Ctrl+Shift+R = REC  /  Ctrl+Shift+S = SAVE", "#777")

        # =================================================
        # Key capture
        # =================================================
        self.installEventFilter(self)
        self.editor.installEventFilter(self)

        # =================================================
        # Poll Excel context
        # =================================================
        self._poll = QTimer(self)
        self._poll.timeout.connect(self._poll_context)
        self._poll.start(self.POLL_MS)

        QTimer.singleShot(0, self._focus_inspector)

        logger.info(f"[Inspector] InspectorPanel ready {focus_snapshot()}")

    # =================================================
    # tracing
    # =================================================
    def _next_trace_id(self) -> int:
        self._trace_seq += 1
        return self._trace_seq

    def _log_key_event(self, phase: str, obj: Any, e: QKeyEvent, trace_id: int):
        info = {
            "trace": trace_id,
            "phase": phase,
            "obj": obj_name(obj),
            "key": key_to_name(e.key()),
            "mod": mod_to_str(e.modifiers()),
            "auto": bool(e.isAutoRepeat()),
            "text": e.text(),
            "t_ms": now_ms(),
        }
        logger.info("[KEY] %s %s %s %s", kv_compact(info), safe_native_info(e), focus_snapshot(), "")

    def _log_exec(self, op: str, trace_id: int, kw: Dict[str, Any]):
        base = {
            "trace": trace_id,
            "op": op,
            "t_ms": now_ms(),
            "edit_mode": self._edit_mode,
        }
        logger.info("[EXEC] %s kw=%s %s", kv_compact(base), kw, focus_snapshot())

    # =================================================
    # Focus helpers
    # =================================================
    def _focus_inspector(self):
        try:
            self.setFocus(Qt.ActiveWindowFocusReason)
        except Exception as e:
            logger.error("[Inspector] focus failed: %s", e, exc_info=True)

    def showEvent(self, event):
        super().showEvent(event)
        try:
            self.raise_()
            self.activateWindow()
        except Exception as e:
            logger.error("[Inspector] showEvent activate failed: %s", e, exc_info=True)
        self._focus_inspector()

    def mousePressEvent(self, event):
        try:
            if not self._edit_mode:
                self._focus_inspector()
        except Exception as e:
            logger.error("[Inspector] mousePress focus failed: %s", e, exc_info=True)
        super().mousePressEvent(event)

    # =================================================
    # External bind
    # =================================================
    def set_tree(self, tree):
        self._tree = tree
        logger.info(f"[Inspector] set_tree tree={obj_name(tree)}")

    def set_current_cell(self, cell: str):
        self._active_cell = cell
        logger.info(f"[Inspector] set_current_cell cell={cell}")

    # =================================================
    # Event filter
    # =================================================
    def eventFilter(self, obj, event):

        # -----------------------------
        # KeyPress：記録のみ（EXECしない）
        # -----------------------------
        if event.type() == QEvent.KeyPress:
            e = event  # type: ignore[assignment]
            if not isinstance(e, QKeyEvent):
                return super().eventFilter(obj, event)

            trace_id = self._next_trace_id()
            self._log_key_event("press", obj, e, trace_id)

            # AutoRepeat は無視（既存仕様）
            if e.isAutoRepeat():
                return True

            # 編集中は editor に任せる（既存仕様）
            if self._edit_mode and obj is self.editor:
                logger.info("[KEY] trace=%s pass_to_editor=True", trace_id)
                return False

            # ★ ここでは EXEC しない
            return True

        # -----------------------------
        # KeyRelease：EXEC はここで
        # -----------------------------
        if event.type() == QEvent.KeyRelease:
            e = event  # type: ignore[assignment]
            if not isinstance(e, QKeyEvent):
                return super().eventFilter(obj, event)

            trace_id = self._next_trace_id()
            self._log_key_event("release", obj, e, trace_id)

            # 編集中は editor に任せる
            if self._edit_mode and obj is self.editor:
                logger.info("[KEY] trace=%s pass_to_editor=True", trace_id)
                return False

            self._handle_key(e, trace_id)
            return True

        return super().eventFilter(obj, event)

    # =================================================
    # Key logic
    # =================================================
    def _handle_key(self, e: QKeyEvent, trace_id: int):
        key = e.key()
        mod = e.modifiers()

        # ---- Macro ----
        if (mod & Qt.ControlModifier) and (mod & Qt.ShiftModifier):
            if key == Qt.Key_R:
                self._toggle_record()
                return
            if key == Qt.Key_S:
                self._save_macro_dialog()
                return

        # ---- Window ----
        if (mod & Qt.AltModifier) and key == Qt.Key_F4:
            self.close()
            return

        # ---- F2 ----
        if key == Qt.Key_F2:
            self._edit_mode = True
            self.editor.setReadOnly(False)
            self.editor.setFocusPolicy(Qt.StrongFocus)
            self.editor.setFocus(Qt.OtherFocusReason)
            self.editor.selectAll()
            self._log_add("Edit (F2)", "#7fd7ff")
            logger.info("[MODE] trace=%s edit_mode=True", trace_id)
            return

        # ---- Editing ----
        if self._edit_mode:
            if key in (Qt.Key_Return, Qt.Key_Enter):
                val = self.editor.text()
                self._exec("set_cell_value", trace_id, cell="*", value=val)
                self._log_add(f"Set = {val}", "#ffb347")
                self.editor.clear()

                self._edit_mode = False
                self.editor.setReadOnly(True)
                self.editor.setFocusPolicy(Qt.NoFocus)
                self._focus_inspector()
                logger.info("[MODE] trace=%s edit_mode=False commit=True", trace_id)
                return

            if key == Qt.Key_Escape:
                self.editor.clear()
                self._edit_mode = False
                self.editor.setReadOnly(True)
                self.editor.setFocusPolicy(Qt.NoFocus)
                self._log_add("Edit cancel (Esc)", "#aaa")
                self._focus_inspector()
                logger.info("[MODE] trace=%s edit_mode=False cancel=True", trace_id)
                return

            logger.info("[MODE] trace=%s editing_ignore key=%s", trace_id, key_to_name(key))
            return

        # ---- Ctrl+Shift (select edge) ----
        if (mod & Qt.ControlModifier) and (mod & Qt.ShiftModifier):
            if key in (Qt.Key_Up, Qt.Key_Down, Qt.Key_Left, Qt.Key_Right):
                self._exec("select_edge", trace_id, direction=self._dir(key))
                self._log_add("Select edge (Ctrl+Shift+Arrow)", "#7fd7ff")
                return

        # ---- Ctrl ----
        if mod & Qt.ControlModifier:
            if key == Qt.Key_A:
                self._exec_and_log("select_all", trace_id, "Select All (Ctrl+A)", "#6cf")
                return
            if key == Qt.Key_C:
                self._exec_and_log("copy", trace_id, "Copy (Ctrl+C)", "#6cf")
                return
            if key == Qt.Key_X:
                self._exec_and_log("cut", trace_id, "Cut (Ctrl+X)", "#6cf")
                return
            if key == Qt.Key_V:
                self._exec_and_log("paste", trace_id, "Paste (Ctrl+V)", "#6cf")
                return
            if key == Qt.Key_Z:
                self._exec_and_log("undo", trace_id, "Undo (Ctrl+Z)", "#6cf")
                return
            if key == Qt.Key_Y:
                self._exec_and_log("redo", trace_id, "Redo (Ctrl+Y)", "#6cf")
                return

            if key in (Qt.Key_Up, Qt.Key_Down, Qt.Key_Left, Qt.Key_Right):
                self._exec("move_edge", trace_id, direction=self._dir(key))
                self._log_add("Move edge (Ctrl+Arrow)", "#aaa")
                return

        # ---- Shift (select move) ----
        if mod & Qt.ShiftModifier:
            if key in (Qt.Key_Up, Qt.Key_Down, Qt.Key_Left, Qt.Key_Right):
                self._exec("select_move", trace_id, direction=self._dir(key))
                self._log_add("Select move (Shift+Arrow)", "#7fd7ff")
                return

        # ---- Arrow ----
        if key in (Qt.Key_Up, Qt.Key_Down, Qt.Key_Left, Qt.Key_Right):
            direction = self._dir(key)
            self._exec("move_cell", trace_id, direction=direction, step=1)
            self._log_add(f"Move {direction} (Arrow)", "#aaa")
            return

        logger.info("[KEY] trace=%s unhandled key=%s mod=%s", trace_id, key_to_name(key), mod_to_str(mod))

    # =================================================
    # Macro
    # =================================================
    def _toggle_record(self):
        if not self._macro.is_recording():
            self._macro.start()
            self.rec_label.setText("● REC")
            self._log_add("REC START (Ctrl+Shift+R)", "#ff4d4d")
            logger.info("[MACRO] start recording=True")
        else:
            self._macro.stop()
            self.rec_label.setText("")
            self._log_add("REC STOP (Ctrl+Shift+R)", "#ff4d4d")
            logger.info("[MACRO] stop recording=False")

    def _save_macro_dialog(self):
        try:
            cnt = self._macro.steps_count()
        except Exception:
            try:
                cnt = len(getattr(self._macro, "_steps", []))
            except Exception:
                cnt = 0

        logger.info("[MACRO] save_dialog steps=%s", cnt)

        if cnt == 0:
            self._log_add("No macro steps (Ctrl+Shift+S)", "#aaa")
            return

        path, _ = QFileDialog.getSaveFileName(self, "Save Macro", "", "Macro JSON (*.json)")
        if not path:
            self._log_add("Save canceled (Ctrl+Shift+S)", "#777")
            logger.info("[MACRO] save canceled")
            return

        try:
            self._macro.save_json(path)
            self._log_add(f"Saved macro ({cnt} steps) (Ctrl+Shift+S)", "#7fd7ff")
            logger.info("[MACRO] saved path=%s steps=%s", path, cnt)
        except Exception as e:
            logger.exception("macro save failed: %s", e)
            self._log_add("Macro save failed", "#f66")

    # =================================================
    # Poll
    # =================================================
    def _poll_context(self):
        if not self._tree:
            return

        try:
            ctx = self._tree._engine_exec("get_active_context")
        except Exception as e:
            logger.error(f"[CTX] get_active_context failed: {e}")
            return

        if not isinstance(ctx, dict):
            logger.info("[CTX] ignored ctx_type=%s", type(ctx))
            return

        addr = str(ctx.get("address", "")).replace("$", "")
        sheet = str(ctx.get("sheet", ""))
        label = f"{sheet}!{addr}" if sheet and addr else "—"

        if label != self._last_ctx:
            self._last_ctx = label
            self.addr_label.setText(label)
            logger.info(f"[CTX] update label={label} ctx={ctx}")

    # =================================================
    # Helpers
    # =================================================
    def _dir(self, key: int) -> str:
        return {
            Qt.Key_Up: "up",
            Qt.Key_Down: "down",
            Qt.Key_Left: "left",
            Qt.Key_Right: "right",
        }[key]

    def _exec(self, op: str, trace_id: int, **kw):
        now_ = now_ms()
        dt = now_ - self._last_exec_at_ms
        self._last_exec_at_ms = now_
        self._last_exec_trace = trace_id

        logger.warning(
            f"[CUT] INSPECTOR_EXEC seq={self._trace_seq} trace={trace_id} op={op} dt={dt}ms kw={kw}")

        if self._tree:
            payload = dict(kw)
            payload["_trace_id"] = trace_id
            payload["_dt_ms_from_prev_exec"] = dt
            self._tree._engine_exec(op, source="inspector", **payload)

    def _exec_and_log(self, op: str, trace_id: int, msg: str, color: str):
        self._exec(op, trace_id)
        self._log_add(msg, color)

    def _log_add(self, msg: str, color: str = "#ddd"):
        self._log_buf.appendleft(f'<span style="color:{color}">▸ {msg}</span>')
        self.log.setHtml("<br>".join(self._log_buf))


# =================================================
# Validation (dev only / same file)
#   ※ ネスト関数禁止のため top-level 化
# =================================================
class _FakeTree:
    def __init__(self):
        self.calls = []

    def _engine_exec(self, op: str, **kw):
        self.calls.append((op, dict(kw)))
        logger.info("[VALIDATE] engine_exec op=%s kw=%s", op, kw)
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

    logger.info("=== VALIDATE: key mapping ===")

    tree.clear()
    QTest.keyClick(panel, Qt.Key_Down)
    _assert_last(tree, "move_cell", direction="down", step=1, source="inspector")

    tree.clear()
    QTest.keyClick(panel, Qt.Key_Right, Qt.ShiftModifier)
    _assert_last(tree, "select_move", direction="right", source="inspector")

    tree.clear()
    QTest.keyClick(panel, Qt.Key_Left, Qt.ControlModifier)
    _assert_last(tree, "move_edge", direction="left", source="inspector")

    tree.clear()
    QTest.keyClick(panel, Qt.Key_Up, Qt.ControlModifier | Qt.ShiftModifier)
    _assert_last(tree, "select_edge", direction="up", source="inspector")

    logger.info("=== VALIDATE: key mapping OK ===")


def validate_edit(panel: InspectorPanel, tree: _FakeTree):
    from PySide6.QtTest import QTest

    logger.info("=== VALIDATE: edit mode ===")

    tree.clear()
    QTest.keyClick(panel, Qt.Key_F2)
    QTest.keyClicks(panel.editor, "abc")
    QTest.keyClick(panel.editor, Qt.Key_Return)

    _assert_last(
        tree,
        "set_cell_value",
        cell="*",
        value="abc",
        source="inspector",
    )

    assert panel._edit_mode is False
    assert panel.editor.isReadOnly() is True

    logger.info("=== VALIDATE: edit mode OK ===")


def validate_autorepeat(panel: InspectorPanel):
    from PySide6.QtTest import QTest
    from PySide6.QtGui import QKeyEvent
    from PySide6.QtCore import QEvent

    logger.info("=== VALIDATE: autorepeat ignore ===")

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
    QTest.qWait(10)

    assert len(tree.calls) == 0, "AutoRepeat で engine_exec が呼ばれています"

    logger.info("=== VALIDATE: autorepeat OK ===")


def validate_user_input(panel: InspectorPanel):
    from PySide6.QtTest import QTest

    logger.info("=== VALIDATE: user key input ===")
    logger.info("↓ 今から実キーを押してください ↓")
    logger.info("  Arrow / Shift+Arrow / Ctrl+Arrow / F2 → Enter")
    logger.info("  （10秒以内）")

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
    panel._poll.stop()  # ノイズ防止
    panel.show()

    QTest.qWaitForWindowExposed(panel)
    panel.activateWindow()
    panel.setFocus()

    validate_keys(panel, fake_tree)
    validate_edit(panel, fake_tree)
    validate_autorepeat(panel)
    validate_user_input(panel)

    logger.info("=== ALL VALIDATION PASSED ===")
