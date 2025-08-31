from __future__ import annotations
from pathlib import Path
from typing import Dict, Any, List
from PIL import Image, ImageDraw

def render_annotated(base_png: Path, meta: Dict[str, Any], out_dir: Path) -> Path:
    """
    PNGに矩形（枠線）を焼き込み、同じフォルダ(または out_dir)に _ann 付きで保存して返す。
    meta["rects"] = [{x,y,w,h,color,stroke}, ...]
    """
    out_dir.mkdir(parents=True, exist_ok=True)
    img = Image.open(base_png).convert("RGBA")
    draw = ImageDraw.Draw(img)

    rects: List[Dict[str, Any]] = meta.get("rects", []) or []
    for r in rects:
        x = int(r.get("x", 0)); y = int(r.get("y", 0))
        w = max(1, int(r.get("w", 1))); h = max(1, int(r.get("h", 1)))
        color = r.get("color", "#FF3B30")
        stroke = max(1, int(r.get("stroke", 2)))
        # 外接矩形（ImageDrawは右下を含む描画になるため -1）
        draw.rectangle([(x, y), (x + w - 1, y + h - 1)], outline=color, width=stroke)

    out_path = out_dir / (base_png.stem + "_ann.png")
    img.save(out_path)
    return out_path
