# core/model.py
from __future__ import annotations
from dataclasses import dataclass, field
from pathlib import Path
from typing import List, Optional, Dict, Any
import json

from core.render import render_annotated  # 既存の焼き込み関数を利用

@dataclass
class Anno:
    x: int
    y: int
    w: int
    h: int
    color: str
    stroke: int

@dataclass
class CaptureItem:
    title: str
    comment: str
    image_path: Path
    annos: List[Anno] = field(default_factory=list)
    device_pixel_ratio: float = 1.0
    meta_raw: Dict[str, Any] = field(default_factory=dict)

    def render_with_annos(self, folder: Path) -> Path:
        src = (folder / self.image_path) if not self.image_path.is_absolute() else self.image_path
        return render_annotated(src, self.meta_raw, folder)

@dataclass
class ExportBundle:
    title: str
    folder: Path
    items: List[CaptureItem]

def load_bundle_from_folder(folder: Path, *, title: Optional[str] = None) -> ExportBundle:
    items: List[CaptureItem] = []
    for jp in sorted(folder.glob("*.json")):
        try:
            meta = json.loads(jp.read_text(encoding="utf-8"))
        except Exception:
            continue
        comment = meta.get("comment", "") or ""
        img_name = meta.get("image_path") or jp.with_suffix(".png").name

        annos: List[Anno] = []
        for r in meta.get("rects", []) or []:
            try:
                annos.append(Anno(
                    x=int(r.get("x", 0)), y=int(r.get("y", 0)),
                    w=int(r.get("w", 0)), h=int(r.get("h", 0)),
                    color=str(r.get("color", "#FF3B30")),
                    stroke=int(r.get("stroke", 2)),
                ))
            except Exception:
                continue

        items.append(CaptureItem(
            title=jp.stem,
            comment=comment,
            image_path=Path(img_name),
            annos=annos,
            device_pixel_ratio=float(meta.get("region", {}).get("device_pixel_ratio", 1.0) or 1.0),
            meta_raw=meta,
        ))
    return ExportBundle(title=(title or "Captures Export").strip(), folder=folder, items=items)
