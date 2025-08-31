# core/config.py
from __future__ import annotations
from pathlib import Path
import json
from typing import Callable, Dict, Optional
from PySide6 import QtCore, QtGui, QtWidgets

ROOT = Path(__file__).resolve().parents[1]
CONFIG_FILE = ROOT / "config.json"

# デフォルト割り当て
DEFAULT_KEYS: Dict[str, str] = {
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

class Config:
    """アプリ全体設定（ホットキー中心）"""
    def __init__(self):
        self.hotkeys: Dict[str, str] = dict(DEFAULT_KEYS)

    def load(self):
        if CONFIG_FILE.exists():
            try:
                data = json.loads(CONFIG_FILE.read_text(encoding="utf-8"))
                if isinstance(data, dict):
                    hk = data.get("hotkeys")
                    if isinstance(hk, dict):
                        # 既知キーのみ反映（未知キーは無視）
                        for k, v in hk.items():
                            if k in DEFAULT_KEYS:
                                self.hotkeys[k] = str(v or "")
            except Exception:
                pass

    def save(self):
        try:
            if CONFIG_FILE.exists():
                try:
                    data = json.loads(CONFIG_FILE.read_text(encoding="utf-8"))
                except Exception:
                    data = {}
            else:
                data = {}
            data["hotkeys"] = self.hotkeys
            CONFIG_FILE.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
        except Exception:
            pass


class HotkeyManager(QtCore.QObject):
    """
    QShortcut を用いてホットキーを登録する簡易ラッパ。
    - parent: ショートカットの親（通常は RegionWindow）
    - actions: key -> callable
    - conf: Config（hotkeys を参照）
    """
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
        """設定（conf.hotkeys）を読み取り直してショートカットを張りなおす"""
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
                # キーが不正などの場合は無視
                continue


# 領域/矩形の最終状態（座標/サイズ/色）を保存・読込
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
        STATE_FILE.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    except Exception:
        pass
