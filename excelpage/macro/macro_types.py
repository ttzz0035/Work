from __future__ import annotations

from dataclasses import dataclass
from typing import Any, Dict, List
from datetime import datetime


@dataclass
class MacroStep:
    op: str
    args: Dict[str, Any]
    ts: datetime


@dataclass
class Macro:
    version: int
    name: str
    created_at: datetime
    steps: List[MacroStep]

    @staticmethod
    def from_dict(d: Dict[str, Any]) -> "Macro":
        steps = [
            MacroStep(
                op=s["op"],
                args=s.get("args", {}),
                ts=datetime.fromisoformat(s["ts"]),
            )
            for s in d.get("steps", [])
        ]

        return Macro(
            version=d.get("version", 1),
            name=d.get("name", ""),
            created_at=datetime.fromisoformat(d["created_at"]),
            steps=steps,
        )
