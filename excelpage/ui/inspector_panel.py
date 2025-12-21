from __future__ import annotations

from collections import deque
from typing import Optional

from PySide6.QtWidgets import (
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QTextEdit,
    QFileDialog,
    QSizePolicy,
)
from PySide6.QtCore import Qt, QEvent, QTimer

from logger import get_logger
from services.macro_recorder import get_macro_recorder

logger = get_logger("Inspector")


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
    """

    MAX_LOG = 10
    POLL_MS = 400

    # =================================================
    # Init
    # =================================================
    def __init__(self):
        super().__init__(None)

        self._tree = None
        self._macro = get_macro_recorder()

        self._log_buf = deque(maxlen=self.MAX_LOG)
        self._edit_mode = False
        self._last_ctx: Optional[str] = None
        self._active_cell: Optional[str] = None

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
        self.addr_label.setStyleSheet("""
            QLabel {
                background:#1b1b1b;
                color:#7fd7ff;
                font-size:14px;
                font-weight:700;
                padding:6px;
                border-radius:6px;
            }
        """)
        # ★ ラベルがフォーカスを奪わない
        self.addr_label.setFocusPolicy(Qt.NoFocus)
        self.addr_label.setTextInteractionFlags(Qt.NoTextInteraction)

        # ★ 幅固定の REC スロット（UIズレ防止）
        self.rec_label = QLabel("")
        self.rec_label.setFixedWidth(56)
        self.rec_label.setAlignment(Qt.AlignCenter)
        self.rec_label.setStyleSheet("""
            QLabel {
                color:#ff4d4d;
                font-weight:900;
            }
        """)
        self.rec_label.setFocusPolicy(Qt.NoFocus)
        self.rec_label.setTextInteractionFlags(Qt.NoTextInteraction)

        # ★ 操作ヒント（常時表示）
        self.hint_label = QLabel("Ctrl+Shift+R:REC   Ctrl+Shift+S:SAVE")
        self.hint_label.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        self.hint_label.setStyleSheet("""
            QLabel {
                color:#666;
                font-size:11px;
                padding-right:4px;
            }
        """)
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
        self.editor.setStyleSheet("""
            QLineEdit {
                background:#151515;
                color:#eee;
                border:1px solid #2a2a2a;
                border-radius:6px;
                padding:8px;
                font-size:14px;
            }
            QLineEdit:focus { border:1px solid #6cf; }
        """)

        # ★ 通常時は editor がフォーカスを奪わない（編集モード時だけ許可）
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
        self.log.setStyleSheet("""
            QTextEdit {
                background:#101010;
                color:#ccc;
                border-radius:6px;
                padding:6px;
            }
        """)

        # ★ ログ領域がフォーカス/キーを奪わない
        self.log.setFocusPolicy(Qt.NoFocus)
        self.log.setTextInteractionFlags(Qt.NoTextInteraction)

        root.addWidget(self.log)

        # 初回：操作案内
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

        logger.info("InspectorPanel ready")

    # =================================================
    # Focus helpers
    # =================================================
    def _focus_inspector(self):
        """
        Inspector が前面にいるときは、Panel 本体がキー司令塔になる。
        ラベル/ログ/editor へフォーカスが逃げない前提で、最後にここへ戻す。
        """
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
        # ★ どこをクリックしても Inspector がキー優先になる
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

    def set_current_cell(self, cell: str):
        self._active_cell = cell

    # =================================================
    # Event filter
    # =================================================
    def eventFilter(self, obj, event):
        if event.type() == QEvent.KeyPress:
            # ★ 編集中は editor に文字入力を任せる（それ以外は Inspector が握る）
            if self._edit_mode and obj is self.editor:
                return False
            self._handle_key(event)
            return True
        return super().eventFilter(obj, event)

    # =================================================
    # Key logic
    # =================================================
    def _handle_key(self, e):
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
            return

        # ---- Editing ----
        if self._edit_mode:
            if key in (Qt.Key_Return, Qt.Key_Enter):
                val = self.editor.text()
                self._exec("set_cell_value", cell="*", value=val)
                self._log_add(f"Set = {val}", "#ffb347")
                self.editor.clear()

                # ★ 編集終了：editor をフォーカス不能へ戻す
                self._edit_mode = False
                self.editor.setReadOnly(True)
                self.editor.setFocusPolicy(Qt.NoFocus)
                self._focus_inspector()
                return

            if key == Qt.Key_Escape:
                self.editor.clear()
                self._edit_mode = False
                self.editor.setReadOnly(True)
                self.editor.setFocusPolicy(Qt.NoFocus)
                self._log_add("Edit cancel (Esc)", "#aaa")
                self._focus_inspector()
                return

            return

        # ---- Ctrl ----
        if mod & Qt.ControlModifier:
            if key == Qt.Key_A:
                self._exec_and_log("select_all", "Select All (Ctrl+A)", "#6cf")
                return
            if key == Qt.Key_C:
                self._exec_and_log("copy", "Copy (Ctrl+C)", "#6cf")
                return
            if key == Qt.Key_X:
                self._exec_and_log("cut", "Cut (Ctrl+X)", "#6cf")
                return
            if key == Qt.Key_V:
                self._exec_and_log("paste", "Paste (Ctrl+V)", "#6cf")
                return
            if key == Qt.Key_Z:
                self._exec_and_log("undo", "Undo (Ctrl+Z)", "#6cf")
                return
            if key == Qt.Key_Y:
                self._exec_and_log("redo", "Redo (Ctrl+Y)", "#6cf")
                return

            if key in (Qt.Key_Up, Qt.Key_Down, Qt.Key_Left, Qt.Key_Right):
                self._exec("move_edge", direction=self._dir(key))
                self._log_add("Move edge (Ctrl+Arrow)", "#aaa")
                return

        # ---- Ctrl+Shift (select edge) ----
        if (mod & Qt.ControlModifier) and (mod & Qt.ShiftModifier):
            if key in (Qt.Key_Up, Qt.Key_Down, Qt.Key_Left, Qt.Key_Right):
                self._exec("select_edge", direction=self._dir(key))
                self._log_add("Select edge (Ctrl+Shift+Arrow)", "#7fd7ff")
                return

        # ---- Shift (select move) ----
        if mod & Qt.ShiftModifier:
            if key in (Qt.Key_Up, Qt.Key_Down, Qt.Key_Left, Qt.Key_Right):
                self._exec("select_move", direction=self._dir(key))
                self._log_add("Select move (Shift+Arrow)", "#7fd7ff")
                return

        # ---- Arrow ----
        if key in (Qt.Key_Up, Qt.Key_Down, Qt.Key_Left, Qt.Key_Right):
            direction = self._dir(key)
            self._exec("move_cell", direction=direction, step=1)
            self._log_add(f"Move {direction} (Arrow)", "#aaa")

    # =================================================
    # Macro
    # =================================================
    def _toggle_record(self):
        if not self._macro.is_recording():
            self._macro.start()
            self.rec_label.setText("● REC")
            self._log_add("REC START (Ctrl+Shift+R)", "#ff4d4d")
        else:
            self._macro.stop()
            self.rec_label.setText("")
            self._log_add("REC STOP (Ctrl+Shift+R)", "#ff4d4d")

    def _save_macro_dialog(self):
        # steps_count が無い実装でも死なないように最低限ガード
        try:
            cnt = self._macro.steps_count()
        except Exception:
            try:
                cnt = len(getattr(self._macro, "_steps", []))
            except Exception:
                cnt = 0

        if cnt == 0:
            self._log_add("No macro steps (Ctrl+Shift+S)", "#aaa")
            return

        path, _ = QFileDialog.getSaveFileName(
            self, "Save Macro", "", "Macro JSON (*.json)"
        )
        if not path:
            self._log_add("Save canceled (Ctrl+Shift+S)", "#777")
            return

        try:
            self._macro.save_json(path)
            self._log_add(
                f"Saved macro ({cnt} steps) (Ctrl+Shift+S)",
                "#7fd7ff",
            )
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
            logger.error("[CTX] get_active_context failed: %s", e, exc_info=True)
            return

        if not isinstance(ctx, dict):
            return

        addr = str(ctx.get("address", "")).replace("$", "")
        sheet = str(ctx.get("sheet", ""))
        label = f"{sheet}!{addr}" if sheet and addr else "—"

        if label != self._last_ctx:
            self._last_ctx = label
            self.addr_label.setText(label)

    # =================================================
    # Helpers
    # =================================================
    def _dir(self, key):
        return {
            Qt.Key_Up: "up",
            Qt.Key_Down: "down",
            Qt.Key_Left: "left",
            Qt.Key_Right: "right",
        }[key]

    def _exec(self, op: str, **kw):
        if self._tree:
            self._tree._engine_exec(op, source="inspector", **kw)

    def _exec_and_log(self, op: str, msg: str, color: str):
        self._exec(op)
        self._log_add(msg, color)

    def _log_add(self, msg: str, color="#ddd"):
        self._log_buf.appendleft(f'<span style="color:{color}">▸ {msg}</span>')
        self.log.setHtml("<br>".join(self._log_buf))
