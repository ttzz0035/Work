from __future__ import annotations

import json
from pathlib import Path
from typing import Callable, Dict

from PySide6 import QtCore, QtGui, QtWidgets

ROOT = Path(__file__).resolve().parents[1]
CONFIG_FILE = ROOT / "config.json"

# ==================================================
# Defaults
# ==================================================
DEFAULT_KEYS: Dict[str, str] = {
    "capture":         "Space",
    "add_rect":        "Ctrl+A",
    "remove_selected": "Delete",
    "pick_new_color":  "Ctrl+C",
    "show_hotkeys":    "Ctrl+/",
    "exit_app":        "Ctrl+Q",
    "rec_start":       "Alt+1",
    "rec_stop":        "Alt+2",
    "rec_play":        "Alt+3",
}

DEFAULT_TOOLBAR = {
    "rect_color": "#FF3B30",
    "rect_stroke": 2,
}

DEFAULT_UI = {
    "toast_duration": 1.2,
}

DEFAULT_RECORD = {
    "last_dir": "",
}

# ==================================================
# Config
# ==================================================
class Config:
    """
    アプリ恒久設定
    - hotkeys
    - toolbar (new rect defaults)
    - ui (toast etc.)
    - record (dialog defaults)
    """

    def __init__(self):
        self.hotkeys: Dict[str, str] = dict(DEFAULT_KEYS)
        self.toolbar: Dict[str, object] = dict(DEFAULT_TOOLBAR)
        self.ui: Dict[str, object] = dict(DEFAULT_UI)
        self.record: Dict[str, object] = dict(DEFAULT_RECORD)

    # ------------------------------
    def load(self) -> None:
        if not CONFIG_FILE.exists():
            return

        try:
            data = json.loads(CONFIG_FILE.read_text(encoding="utf-8"))
        except Exception:
            return

        if not isinstance(data, dict):
            return

        # --- hotkeys ---
        hk = data.get("hotkeys")
        if isinstance(hk, dict):
            for k, v in hk.items():
                if k in DEFAULT_KEYS:
                    self.hotkeys[k] = str(v or "")

        # --- toolbar ---
        tb = data.get("toolbar")
        if isinstance(tb, dict):
            col = tb.get("rect_color")
            if isinstance(col, str) and col.startswith("#"):
                self.toolbar["rect_color"] = col

            stroke = tb.get("rect_stroke")
            try:
                stroke_i = int(stroke)
                if 1 <= stroke_i <= 20:
                    self.toolbar["rect_stroke"] = stroke_i
            except Exception:
                pass

        # --- ui ---
        ui = data.get("ui")
        if isinstance(ui, dict):
            td = ui.get("toast_duration")
            try:
                td_f = float(td)
                if 0.1 <= td_f <= 10.0:
                    self.ui["toast_duration"] = td_f
            except Exception:
                pass

        # --- record ---
        rec = data.get("record")
        if isinstance(rec, dict):
            ld = rec.get("last_dir")
            if isinstance(ld, str):
                self.record["last_dir"] = ld

    # ------------------------------
    def save(self) -> None:
        try:
            data: Dict[str, object] = {}
            if CONFIG_FILE.exists():
                try:
                    data = json.loads(CONFIG_FILE.read_text(encoding="utf-8"))
                except Exception:
                    data = {}

            data["hotkeys"] = self.hotkeys
            data["toolbar"] = self.toolbar
            data["ui"] = self.ui
            data["record"] = self.record

            CONFIG_FILE.write_text(
                json.dumps(data, ensure_ascii=False, indent=2),
                encoding="utf-8",
            )
        except Exception:
            pass


# ==================================================
# Hotkey manager
# ==================================================
class HotkeyManager(QtCore.QObject):
    def __init__(self, parent: QtWidgets.QWidget, actions: Dict[str, Callable], conf: Config):
        super().__init__(parent)
        self.parent = parent
        self.actions = actions
        self.conf = conf
        self._shortcuts: Dict[str, QtGui.QShortcut] = {}

    def clear(self):
        for sc in self._shortcuts.values():
            try:
                sc.disconnect()
            except Exception:
                pass
            sc.setParent(None)
        self._shortcuts.clear()

    def apply(self):
        self.clear()
        for key_name, seq in self.conf.hotkeys.items():
            if not seq:
                continue
            act = self.actions.get(key_name)
            if not callable(act):
                continue
            try:
                ks = QtGui.QKeySequence(seq)
                sc = QtGui.QShortcut(ks, self.parent)
                sc.activated.connect(act)
                self._shortcuts[key_name] = sc
            except Exception:
                continue


# ==================================================
# Last region state (作業状態)
# ==================================================
STATE_FILE = ROOT / "last_state.json"

def load_last_state() -> dict:
    if STATE_FILE.exists():
        try:
            return json.loads(STATE_FILE.read_text(encoding="utf-8"))
        except Exception:
            return {}
    return {}

def save_last_state(data: dict) -> None:
    try:
        STATE_FILE.write_text(
            json.dumps(data, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
    except Exception:
        pass
