# core/capture_service.py
from __future__ import annotations

import json
import logging
import time
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, List, Optional

from mss import mss, tools

from core.calibration import (
    Calibration,
    default_calibration_path,
    load_calibration,
)

# ==================================================
# Logger (no root pollution)
# ==================================================
def _get_logger() -> logging.Logger:
    logger = logging.getLogger("core.CaptureService")
    if not logger.handlers:
        logger.setLevel(logging.DEBUG)
        h = logging.StreamHandler()
        h.setLevel(logging.DEBUG)
        fmt = logging.Formatter("%(asctime)s [%(levelname)s] %(name)s: %(message)s")
        h.setFormatter(fmt)
        logger.addHandler(h)
        logger.propagate = False
    return logger


logger = _get_logger()

# ==================================================
# DTO
# ==================================================
@dataclass
class CaptureAnnoRect:
    x: int
    y: int
    w: int
    h: int
    color: str
    stroke: int


@dataclass
class CaptureRegionLocal:
    x: int
    y: int
    w: int
    h: int


@dataclass
class CaptureGlobalTopLeft:
    x: int
    y: int


@dataclass
class CaptureRequest:
    save_dir: Path
    region_local: CaptureRegionLocal
    global_top_left: CaptureGlobalTopLeft
    device_pixel_ratio: float
    annos: List[CaptureAnnoRect]
    version: int = 3


@dataclass
class CaptureResult:
    ok: bool
    message: str
    png_path: Optional[Path] = None
    json_path: Optional[Path] = None
    meta: Optional[Dict[str, Any]] = None


# ==================================================
# Service
# ==================================================
class CaptureService:
    def __init__(self, logger_: Optional[logging.Logger] = None, calibration_path: Optional[Path] = None):
        self.logger = logger_ if logger_ is not None else logger
        self.calibration_path = calibration_path

    def capture(self, req: CaptureRequest) -> CaptureResult:
        self.logger.debug("=== CaptureService.capture start ===")
        try:
            req.save_dir.mkdir(exist_ok=True, parents=True)
            ts = time.strftime("%Y%m%d_%H%M%S")

            # ------------------------------
            # calibration
            # ------------------------------
            cal_path = self.calibration_path or default_calibration_path(req.save_dir)
            cal = load_calibration(cal_path)
            eff_scale = self._effective_scale(req.device_pixel_ratio, cal)

            self.logger.debug(
                "calibration path=%s scale=%.6f off=(%.3f,%.3f) dpr=%.6f eff_scale=%.6f",
                cal_path,
                cal.scale,
                cal.off_x,
                cal.off_y,
                float(req.device_pixel_ratio or 1.0),
                eff_scale,
            )

            # ------------------------------
            # bbox
            # ------------------------------
            bbox = self._build_bbox(req, cal, eff_scale)
            self.logger.debug("bbox=%s", bbox)

            png_path = req.save_dir / f"capture_{ts}.png"
            self._grab_to_png(bbox, png_path)

            # ------------------------------
            # rects logical -> image px
            # ------------------------------
            rects_img_px = self._convert_rects_to_image_px(
                region_local=req.region_local,
                annos=req.annos,
                cal=cal,
                eff_scale=eff_scale,
            )

            meta = self._build_meta(
                ts=ts,
                req=req,
                png_path=png_path,
                rects_img_px=rects_img_px,
                cal=cal,
                eff_scale=eff_scale,
                calibration_path=cal_path,
            )

            json_path = png_path.with_suffix(".json")
            json_path.write_text(json.dumps(meta, ensure_ascii=False, indent=2), encoding="utf-8")

            self.logger.debug("saved png=%s", png_path)
            self.logger.debug("saved json=%s", json_path)
            self.logger.debug("=== CaptureService.capture end (ok) ===")

            return CaptureResult(True, "ok", png_path, json_path, meta)

        except Exception as e:
            self.logger.exception("CaptureService.capture failed: %s", e)
            return CaptureResult(False, str(e))

    # --------------------------------------------------
    # helpers
    # --------------------------------------------------
    def _effective_scale(self, dpr: float, cal: Calibration) -> float:
        d = float(dpr or 1.0)
        s = float(cal.scale or 1.0)
        eff = d * s
        if eff == 0.0:
            eff = 1.0
        return eff

    def _build_bbox(self, req: CaptureRequest, cal: Calibration, eff_scale: float) -> Dict[str, int]:
        left = int(round(req.global_top_left.x * eff_scale + cal.off_x))
        top = int(round(req.global_top_left.y * eff_scale + cal.off_y))
        width = int(round(req.region_local.w * eff_scale))
        height = int(round(req.region_local.h * eff_scale))

        if width < 1:
            width = 1
        if height < 1:
            height = 1

        bbox = {
            "left": left,
            "top": top,
            "width": width,
            "height": height,
        }

        self.logger.debug(
            "build_bbox global=(%d,%d) local_wh=(%d,%d) eff_scale=%.6f off=(%.3f,%.3f)",
            req.global_top_left.x,
            req.global_top_left.y,
            req.region_local.w,
            req.region_local.h,
            eff_scale,
            cal.off_x,
            cal.off_y,
        )
        return bbox

    def _grab_to_png(self, bbox: Dict[str, int], png_path: Path) -> None:
        self.logger.debug("grab_to_png start png_path=%s bbox=%s", png_path, bbox)
        with mss() as sct:
            img = sct.grab(bbox)
            tools.to_png(img.rgb, img.size, output=str(png_path))
        self.logger.debug("grab_to_png end")

    def _convert_rects_to_image_px(
        self,
        *,
        region_local: CaptureRegionLocal,
        annos: List[CaptureAnnoRect],
        cal: Calibration,
        eff_scale: float,
    ) -> List[Dict[str, Any]]:
        rx0 = int(region_local.x)
        ry0 = int(region_local.y)

        out: List[Dict[str, Any]] = []
        for r in annos:
            x = int(round((r.x - rx0) * eff_scale))
            y = int(round((r.y - ry0) * eff_scale))
            w = max(1, int(round(r.w * eff_scale)))
            h = max(1, int(round(r.h * eff_scale)))
            out.append(
                {
                    "x": x,
                    "y": y,
                    "w": w,
                    "h": h,
                    "color": r.color,
                    "stroke": int(r.stroke),
                }
            )
        return out

    def _build_meta(
        self,
        *,
        ts: str,
        req: CaptureRequest,
        png_path: Path,
        rects_img_px: List[Dict[str, Any]],
        cal: Calibration,
        eff_scale: float,
        calibration_path: Path,
    ) -> Dict[str, Any]:
        meta: Dict[str, Any] = {
            "timestamp": ts,
            "region": {
                "left_global": int(req.global_top_left.x),
                "top_global": int(req.global_top_left.y),
                "width": int(req.region_local.w),
                "height": int(req.region_local.h),
                "device_pixel_ratio": float(req.device_pixel_ratio),
            },
            "calibration": {
                "path": str(calibration_path),
                "scale": float(cal.scale),
                "off_x": float(cal.off_x),
                "off_y": float(cal.off_y),
                "effective_scale": float(eff_scale),
            },
            "rects": [
                {
                    "x": int(r.x),
                    "y": int(r.y),
                    "w": int(r.w),
                    "h": int(r.h),
                    "color": str(r.color),
                    "stroke": int(r.stroke),
                }
                for r in req.annos
            ],
            "rects_img_px": rects_img_px,
            "image_path": png_path.name,
            "comment": "",
            "version": int(req.version),
        }
        return meta
