# core/calibration.py
from __future__ import annotations
from dataclasses import dataclass
from pathlib import Path
import json


@dataclass
class Calibration:
    scale: float = 1.0
    off_x: float = 0.0
    off_y: float = 0.0


def default_calibration_path(folder: Path) -> Path:
    return folder / "calibration.json"


# --------------------------------------------------
# low level
# --------------------------------------------------
def _load_all(path: Path) -> dict:
    if not path.exists():
        return {"tabs": {}, "active": None}
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return {"tabs": {}, "active": None}


def _save_all(path: Path, data: dict) -> None:
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


# --------------------------------------------------
# public API
# --------------------------------------------------
def load_calibration(path: Path) -> Calibration:
    """
    CaptureService 用。
    active タブがあればそれを返す。
    なければ default Calibration。
    """
    data = _load_all(path)
    key = data.get("active")
    tab = (data.get("tabs") or {}).get(key) if key else None
    if not tab:
        return Calibration()
    return Calibration(
        scale=float(tab.get("scale", 1.0)),
        off_x=float(tab.get("off_x", 0.0)),
        off_y=float(tab.get("off_y", 0.0)),
    )


def load_active(path: Path) -> Calibration:
    # calibration_panel 互換
    return load_calibration(path)


def set_tab(path: Path, name: str, cal: Calibration, *, activate: bool = True) -> None:
    data = _load_all(path)
    tabs = data.setdefault("tabs", {})
    tabs[name] = {
        "scale": float(cal.scale),
        "off_x": float(cal.off_x),
        "off_y": float(cal.off_y),
    }
    if activate:
        data["active"] = name
    _save_all(path, data)


def set_active(path: Path, name: str) -> None:
    data = _load_all(path)
    data["active"] = name
    _save_all(path, data)
