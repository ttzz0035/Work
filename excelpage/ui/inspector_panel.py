from __future__ import annotations

from collections import deque
from typing import Optional, Dict, Any

from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel,
    QLineEdit, QTextEdit
)
from PySide6.QtCore import Qt, QEvent, QTimer

from logger import get_logger

logger = get_logger("Inspector")


class InspectorPanel(QWidget):
    """
    Excel Inspector (Excel-like, Keyboard-first)

    - F2 : edit cell
    - Enter : commit
    - Esc : cancel edit only
    - Ctrl / Shift / Arrow : Excel compatible（engine opに委譲）
    - Alt+F4 : close Inspector window
    - Show active sheet + selection (A1 or A1:C10)
    """

    MAX_LOG = 10
    POLL_MS = 400

    # =================================================
    # Init
    # =================================================
    def __init__(self):
        super().__init__(None)

        self._tree = None
        self._log_buf = deque(maxlen=self.MAX_LOG)

        self._edit_mode = False
        self._last_ctx: Optional[str] = None

        self._active_cell: Optional[str] = None  # 互換用（外部から入る）

        # ---- window
        self.setWindowTitle("Excel Inspector")
        self.setWindowFlags(Qt.Window | Qt.WindowStaysOnTopHint)
        self.resize(520, 360)

        # ---- background
        self.setStyleSheet("""
            QWidget { background-color:#0f0f0f; }
        """)

        root = QVBoxLayout(self)
        root.setContentsMargins(12, 12, 12, 12)
        root.setSpacing(10)

        # =================================================
        # Formula bar
        # =================================================
        bar = QHBoxLayout()
        bar.setSpacing(8)

        self.addr_label = QLabel("—")
        self.addr_label.setFixedWidth(180)
        self.addr_label.setAlignment(Qt.AlignCenter)
        self.addr_label.setStyleSheet("""
            QLabel {
                background:#1b1b1b;
                color:#7fd7ff;
                font-size:13px;
                font-weight:700;
                padding:6px;
                border-radius:6px;
            }
        """)

        fx = QLabel("fx")
        fx.setFixedWidth(26)
        fx.setAlignment(Qt.AlignCenter)
        fx.setStyleSheet("color:#6cf; font-weight:700;")

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

        bar.addWidget(self.addr_label)
        bar.addWidget(fx)
        bar.addWidget(self.editor)
        root.addLayout(bar)

        # =================================================
        # Log
        # =================================================
        self.log = QTextEdit()
        self.log.setReadOnly(True)
        self.log.setMaximumHeight(160)
        self.log.setStyleSheet("""
            QTextEdit {
                background:#101010;
                color:#ccc;
                border-radius:6px;
                padding:6px;
            }
        """)
        root.addWidget(self.log)

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

        QTimer.singleShot(0, self.setFocus)

        logger.info("InspectorPanel ready")

    # =================================================
    # External bind
    # =================================================
    def set_tree(self, tree):
        self._tree = tree

    # ★ 互換維持：tree_view.py がここに connect している
    def set_current_cell(self, cell: str):
        """
        互換用：ExcelWorker からの active_cell_changed を受ける。
        ここでは表示更新のヒントとして保持するだけ。
        """
        self._active_cell = cell
        logger.info("[CTX] active_cell_changed=%s", cell)

        # なるべく自然に反映（sheet名は不明なのでセルだけ）
        if cell:
            # 既存表示が "Sheet!A1:C10" なら上書きしない
            if "!" not in (self.addr_label.text() or ""):
                self.addr_label.setText(str(cell).replace("$", ""))

    # =================================================
    # Event filter
    # =================================================
    def eventFilter(self, obj, event):
        if event.type() == QEvent.KeyPress:
            # 編集中は editor に文字入力を任せる（握りつぶさない）
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

        # --- Alt+F4 : close Inspector ---
        if (mod & Qt.AltModifier) and key == Qt.Key_F4:
            self.close()
            return

        # --- F2 edit ---
        if key == Qt.Key_F2:
            self._edit_mode = True
            self.editor.setFocus(Qt.OtherFocusReason)
            self.editor.selectAll()
            self._log_add("Edit (F2)", "#7fd7ff")
            return

        # --- Editing ---
        if self._edit_mode:
            if key in (Qt.Key_Return, Qt.Key_Enter):
                val = self.editor.text()
                # ★ cell="*" を必ず付ける（tree_view engine互換）
                self._exec("set_cell_value", cell="*", value=val)
                self._log_add(f"Set = {val}", "#ffb347")
                self._edit_mode = False
                self.editor.clear()
                self.setFocus(Qt.OtherFocusReason)
                return

            if key == Qt.Key_Escape:
                self.editor.clear()
                self._edit_mode = False
                self.setFocus(Qt.OtherFocusReason)
                self._log_add("Edit cancel", "#aaa")
                return

            return  # let editor handle text

        # --- Ctrl shortcuts (engineへ委譲) ---
        if mod & Qt.ControlModifier:
            if key == Qt.Key_A:
                self._exec_and_log("select_all", "Select All", "#6cf"); return
            if key == Qt.Key_Z:
                self._exec_and_log("undo", "Undo", "#6cf"); return
            if key == Qt.Key_Y:
                self._exec_and_log("redo", "Redo", "#6cf"); return
            if key == Qt.Key_C:
                self._exec_and_log("copy", "Copy", "#6cf"); return
            if key == Qt.Key_V:
                self._exec_and_log("paste", "Paste", "#6cf"); return
            if key == Qt.Key_X:
                self._exec_and_log("cut", "Cut", "#6cf"); return
            if key == Qt.Key_S:
                self._exec_and_log("save", "Save", "#6cf"); return

            # Ctrl + Arrow
            if key in (Qt.Key_Up, Qt.Key_Down, Qt.Key_Left, Qt.Key_Right):
                direction = self._dir(key)
                self._exec("move_edge", direction=direction)
                self._log_add("Move edge", "#aaa")
                return

        # Ctrl+Shift + Arrow（select to edge）
        if (mod & Qt.ControlModifier) and (mod & Qt.ShiftModifier) and key in (
            Qt.Key_Up, Qt.Key_Down, Qt.Key_Left, Qt.Key_Right
        ):
            direction = self._dir(key)
            self._exec("select_edge", direction=direction)
            self._log_add("Select to edge", "#7fd7ff")
            return

        # Shift + Arrow（select move）
        if (mod & Qt.ShiftModifier) and key in (
            Qt.Key_Up, Qt.Key_Down, Qt.Key_Left, Qt.Key_Right
        ):
            direction = self._dir(key)
            self._exec("select_move", direction=direction)
            self._log_add("Select move", "#7fd7ff")
            return

        # Arrow（move）
        if key in (Qt.Key_Up, Qt.Key_Down, Qt.Key_Left, Qt.Key_Right):
            direction = self._dir(key)
            self._exec("move_cell", direction=direction, step=1)
            self._log_add(f"Move {direction}", "#aaa")
            return

        # 通常時の ESC は何もしない（Excel準拠・ノイズ抑止）

    # =================================================
    # Poll Excel context
    # =================================================
    def _poll_context(self):
        if not self._tree:
            return

        try:
            ctx = self._tree._engine_exec("get_active_context")
        except Exception as e:
            logger.error("[CTX] get_active_context failed: %s", e)
            return

        if not isinstance(ctx, dict):
            return

        addr = str(ctx.get("address", "")).replace("$", "")
        sheet = str(ctx.get("sheet", ""))
        book = str(ctx.get("workbook", ""))

        label = "—"
        if sheet and addr:
            label = f"{sheet}!{addr}"
        elif addr:
            label = addr
        elif self._active_cell:
            label = str(self._active_cell).replace("$", "")

        if label != self._last_ctx:
            self._last_ctx = label
            self.addr_label.setText(label)

            # 過剰ログは避ける（必要ならここをON）
            # self._log_add(f"Active {book}/{sheet}/{addr}", "#6cf")

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
            self._tree._engine_exec(op, **kw)

    def _exec_and_log(self, op: str, msg: str, color: str):
        self._exec(op)
        self._log_add(msg, color)

    def _log_add(self, msg: str, color="#ddd"):
        self._log_buf.appendleft(
            f'<span style="color:{color}">▸ {msg}</span>'
        )
        self.log.setHtml("<br>".join(self._log_buf))
