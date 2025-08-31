from __future__ import annotations
from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable, Dict, List, Optional
import json
import logging

from PySide6 import QtCore, QtGui

# ---- ログ設定（既存ロガーに統合） ----
log = logging.getLogger("hotkeys")
if not log.handlers:
    h = logging.StreamHandler()
    fmt = logging.Formatter("%(asctime)s [%(levelname)s] [HOTKEY] %(message)s")
    h.setFormatter(fmt)
    log.addHandler(h)
log.setLevel(logging.INFO)

ROOT = Path(__file__).resolve().parents[1]
CONFIG_FILE = ROOT / "config.json"

# デフォルト割り当て
DEFAULT_KEYS = {
    "capture":         "Ctrl+Space",
    "add_rect":        "Ctrl+A",
    "remove_selected": "Delete",
    "pick_new_color":  "Ctrl+C",
    "show_hotkeys":    "Ctrl+/",
    "exit_app":        "Ctrl+Q",
    "rec_start":       "Alt+1",
    "rec_stop":        "Alt+2",
    "rec_play":        "Alt+3",
}

@dataclass
class Config:
    hotkeys: Dict[str, str] = field(default_factory=lambda: DEFAULT_KEYS.copy())

    def load(self):
        if CONFIG_FILE.exists():
            try:
                data = json.loads(CONFIG_FILE.read_text(encoding="utf-8"))
                hk = data.get("hotkeys")
                if isinstance(hk, dict):
                    # 既知キーのみ採用、未知キーは無視
                    loaded = {k: str(hk.get(k, DEFAULT_KEYS[k])) for k in DEFAULT_KEYS.keys()}
                    self.hotkeys.update(loaded)
                log.info(f"loaded keys: {self.hotkeys}")
            except Exception as e:
                log.warning(f"failed to load config: {e}")

    def save(self):
        try:
            obj = {"hotkeys": self.hotkeys}
            CONFIG_FILE.write_text(json.dumps(obj, ensure_ascii=False, indent=2), encoding="utf-8")
            log.info("config saved")
        except Exception as e:
            log.error(f"failed to save config: {e}")

# ---- Hotkey Manager ----
class HotkeyManager(QtCore.QObject):
    """
    - QtGui.QShortcut を使用
    - context = ApplicationShortcut
    - ショートカットを self._shortcuts に保持（GC対策）
    """
    def __init__(self,
                 target_widget: QtCore.QObject,
                 actions: Dict[str, Callable[[], None]],
                 config: Config):
        super().__init__(target_widget)
        self._target = target_widget   # QWidget 前提
        self._actions = actions        # key -> slot
        self._config = config
        self._shortcuts: List[QtGui.QShortcut] = []

    def clear(self):
        for sc in self._shortcuts:
            try:
                sc.activated.disconnect()
            except Exception:
                pass
            sc.setParent(None)
            sc.deleteLater()
        self._shortcuts.clear()

    def apply(self):
        """config に基づいてショートカット再登録"""
        self.clear()
        ok_count = 0
        for key, default_seq in DEFAULT_KEYS.items():
            seq_str = self._config.hotkeys.get(key, default_seq) or ""
            if not seq_str.strip():
                log.info(f"skip empty hotkey: {key}")
                continue
            if key not in self._actions:
                log.info(f"skip unknown action: {key}")
                continue

            seq = QtGui.QKeySequence(seq_str)
            if seq.isEmpty():
                log.warning(f"invalid key sequence -> '{seq_str}' for {key}; skipped")
                continue

            sc = QtGui.QShortcut(seq, self._target)  # ← QtGui.QShortcut !!
            sc.setContext(QtCore.Qt.ApplicationShortcut)
            # ラッパを噛ませてログ
            slot = self._wrap_action(key, self._actions[key])
            sc.activated.connect(slot)
            self._shortcuts.append(sc)
            ok_count += 1
            log.info(f"registered: {key} = {seq.toString()}")

        log.info(f"hotkeys applied: {ok_count} active, target={type(self._target).__name__}")

    def _wrap_action(self, key: str, func: Callable[[], None]):
        def _inner():
            try:
                log.info(f"activated: {key}")
                func()
            except Exception as e:
                log.error(f"action error ({key}): {e}")
        return _inner

    # 動的変更用（設定ダイアログから呼ばれる想定なら）
    def set_key(self, key: str, seq: str):
        if key in DEFAULT_KEYS:
            self._config.hotkeys[key] = seq
            self._config.save()
            self.apply()

# ---- 補助：最後の領域/矩形状態の保存（既存のまま） ----
STATE_FILE = ROOT / "last_state.json"

def load_last_state() -> dict:
    if STATE_FILE.exists():
        try:
            return json.loads(STATE_FILE.read_text(encoding="utf-8"))
        except Exception as e:
            log.warning(f"failed to load last_state: {e}")
    return {}

def save_last_state(data: dict):
    try:
        STATE_FILE.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    except Exception as e:
        log.error(f"failed to save last_state: {e}")
