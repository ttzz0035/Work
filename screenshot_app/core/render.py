from __future__ import annotations

import logging
from pathlib import Path
from typing import Dict, Any, List
from PIL import Image, ImageDraw


# ==================================================
# Logger (no root pollution)
# ==================================================
def _get_logger() -> logging.Logger:
    logger = logging.getLogger("Screenshot.Render")
    if not logger.handlers:
        logger.setLevel(logging.DEBUG)
        h = logging.StreamHandler()
        h.setLevel(logging.DEBUG)
        fmt = logging.Formatter(
            "%(asctime)s [%(levelname)s] %(name)s: %(message)s"
        )
        h.setFormatter(fmt)
        logger.addHandler(h)
        logger.propagate = False
    return logger


logger = _get_logger()


# ==================================================
# Render
# ==================================================
def render_annotated(base_png: Path, meta: Dict[str, Any], out_dir: Path) -> Path:
    """
    前提（version >= 3）:
    - meta["rects_img_px"] は base_png 左上 (0,0) 基準の画像ピクセル
    - 本関数では座標補正・スケール・オフセット計算を一切行わない
    """
    logger.debug("=== render_annotated start ===")
    logger.debug("base_png=%s", base_png)
    logger.debug("out_dir=%s", out_dir)

    out_dir.mkdir(parents=True, exist_ok=True)

    img = Image.open(base_png).convert("RGBA")
    img_w, img_h = img.size
    logger.debug("image_size w=%d h=%d", img_w, img_h)

    draw = ImageDraw.Draw(img)

    rects: List[Dict[str, Any]] = meta.get("rects_img_px", []) or []
    logger.debug("rect_count(rects_img_px)=%d", len(rects))

    for idx, r in enumerate(rects):
        x = int(r.get("x", 0))
        y = int(r.get("y", 0))
        w = max(1, int(r.get("w", 1)))
        h = max(1, int(r.get("h", 1)))
        color = r.get("color", "#FF3B30")
        stroke = max(1, int(r.get("stroke", 2)))

        x2 = x + w - 1
        y2 = y + h - 1

        logger.debug(
            "[rect %d] img_px (%d,%d)-(%d,%d)",
            idx, x, y, x2, y2
        )

        draw.rectangle(
            [(x, y), (x2, y2)],
            outline=color,
            width=stroke,
        )

    out_path = out_dir / (base_png.stem + "_ann.png")
    img.save(out_path)

    logger.debug("saved=%s", out_path)
    logger.debug("=== render_annotated end ===")

    return out_path
