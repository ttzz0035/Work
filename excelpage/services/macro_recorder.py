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
    ts: str


class MacroRecorder:
    """
    アプリ実行エンジン用マクロ記録
    （Excel直接操作ではなく、意味コマンドを保存）
    """

    def __init__(self):
        self._recording: bool = False
        self._steps: List[MacroStep] = []

    def start(self):
        self._steps.clear()
        self._recording = True
        logger.info("[MACRO] start")

    def stop(self):
        self._recording = False
        logger.info("[MACRO] stop steps=%d", len(self._steps))

    def is_recording(self) -> bool:
        return self._recording

    def record(self, op: str, **kwargs):
        if not self._recording:
            return
        step = MacroStep(
            op=op,
            args=dict(kwargs),
            ts=datetime.now().isoformat(timespec="seconds"),
        )
        self._steps.append(step)
        logger.info("[MACRO] record %s %s", op, kwargs)

    def save(self, path: str):
        payload = {
            "version": 1,
            "created_at": datetime.now().isoformat(timespec="seconds"),
            "steps": [asdict(s) for s in self._steps],
        }
        with open(path, "w", encoding="utf-8") as f:
            json.dump(payload, f, indent=2, ensure_ascii=False)
        logger.info("[MACRO] saved %s", path)


_recorder: Optional[MacroRecorder] = None


def get_macro_recorder() -> MacroRecorder:
    global _recorder
    if _recorder is None:
        _recorder = MacroRecorder()
    return _recorder
