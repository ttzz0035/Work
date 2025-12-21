# services/macro_recorder.py
from __future__ import annotations

import json
from dataclasses import dataclass, asdict
from datetime import datetime
from typing import Any, Dict, List, Optional

from logger import get_logger

logger = get_logger("MacroRecorder")


@dataclass
class MacroStep:
    op: str
    args: Dict[str, Any]
    ts: str  # ISO8601


class MacroRecorder:
    """
    意味操作ベースのマクロレコーダー

    - UIイベントではなく Engine op を記録
    - Inspector からのみ記録される想定
    - JSON 保存 / 将来の再生用 payload を生成
    """

    def __init__(self):
        self._enabled: bool = False
        self._steps: List[MacroStep] = []
        self._name: str = "macro"

    # =================================================
    # State
    # =================================================
    def is_recording(self) -> bool:
        return self._enabled

    def steps_count(self) -> int:
        """★ Inspector 互換用（重要）"""
        return len(self._steps)

    # =================================================
    # Control
    # =================================================
    def start(self):
        self._steps.clear()
        self._enabled = True
        self._name = "macro"
        logger.info("[MACRO] start")

    def stop(self):
        self._enabled = False
        logger.info("[MACRO] stop steps=%s", len(self._steps))

    def clear(self):
        n = len(self._steps)
        self._steps.clear()
        logger.info("[MACRO] clear steps=%s", n)

    # =================================================
    # Record
    # =================================================
    def record(self, op: str, **kwargs):
        if not self._enabled:
            return

        step = MacroStep(
            op=op,
            args=dict(kwargs),
            ts=datetime.now().isoformat(timespec="seconds"),
        )
        self._steps.append(step)
        logger.info("[MACRO] record op=%s args=%s", op, kwargs)

    # =================================================
    # Export
    # =================================================
    def export_payload(self) -> Dict[str, Any]:
        return {
            "version": 1,
            "name": self._name,
            "created_at": datetime.now().isoformat(timespec="seconds"),
            "steps": [asdict(s) for s in self._steps],
        }

    def save_json(self, path: str):
        payload = self.export_payload()
        with open(path, "w", encoding="utf-8") as f:
            json.dump(payload, f, indent=2, ensure_ascii=False)
        logger.info(
            "[MACRO] saved json=%s steps=%s",
            path,
            len(self._steps),
        )


# =====================================================
# Singleton
# =====================================================
_recorder: Optional[MacroRecorder] = None


def get_macro_recorder() -> MacroRecorder:
    global _recorder
    if _recorder is None:
        _recorder = MacroRecorder()
        logger.info("[MACRO] singleton created")
    return _recorder
