from __future__ import annotations
from pathlib import Path
import logging, tempfile, shutil, math

import xlwings as xw
from PIL import Image  # Pillow

from .base import BaseExporter, ExportOptions, DEFAULT_TITLE
from .registry import register
from core.model import ExportBundle
from core.calibration import Calibration

log = logging.getLogger("export.excel")


@register
class ExcelExporter(BaseExporter):
    """
    Excel エクスポータ（ratio-based 描画）
    - px → pt 変換は一切しない
    - picture を基準に「割合」で Shape を配置
    - calibration は px 空間で適用
    """
    name = "excel"
    ext = ".xlsx"

    MAX_IMG_W_PT = 720.0
    MAX_IMG_H_PT = 540.0
    PT_PER_COL = 48.0

    # --------------------------------------------------
    def export_bundle(self, bundle: ExportBundle, options: ExportOptions) -> Path:
        sheet_title = (options.title or bundle.title or DEFAULT_TITLE).strip()
        out = options.filename or (bundle.folder / f"captures_export{self.ext}")
        out.parent.mkdir(parents=True, exist_ok=True)

        app = xw.App(visible=False, add_book=False)
        prev_upd = app.screen_updating
        prev_disp = app.display_alerts
        try:
            app.screen_updating = False
            app.display_alerts = False

            if out.exists():
                book = app.books.open(str(out))
            else:
                book = app.books.add()
                book.save(str(out))

            try:
                self.remove_existing_by_title(
                    targets=book.sheets,
                    match_title=sheet_title,
                    get_name=lambda s: s.name,
                    delete=lambda s: s.delete(),
                )
                sht = book.sheets.add(
                    name=sheet_title,
                    after=book.sheets[-1] if book.sheets else None,
                )

                current_row = 1
                for it in bundle.items:
                    meta = getattr(it, "meta_raw", {}) or {}
                    title = self._resolve_title(it)
                    comment = it.comment or "(no comment)"

                    base_img_rel = meta.get("image_path")
                    if not base_img_rel:
                        sht.range((current_row, 1)).value = title
                        current_row += 3
                        continue

                    base_img_abs = (
                        bundle.folder / base_img_rel
                        if not Path(base_img_rel).is_absolute()
                        else Path(base_img_rel)
                    )
                    if not base_img_abs.exists():
                        sht.range((current_row, 1)).value = f"{title} (image not found)"
                        current_row += 3
                        continue

                    # --- image size (px) ---
                    with Image.open(str(base_img_abs)) as im:
                        img_w_px, img_h_px = im.size

                    # --- scale picture size (pt) ---
                    w_pt_raw = img_w_px * 0.75
                    h_pt_raw = img_h_px * 0.75
                    scale = min(
                        1.0,
                        self.MAX_IMG_W_PT / max(w_pt_raw, 1.0),
                        self.MAX_IMG_H_PT / max(h_pt_raw, 1.0),
                    )
                    pic_w_pt = w_pt_raw * scale
                    pic_h_pt = h_pt_raw * scale

                    sht.range((current_row, 1)).value = title
                    current_row += 1

                    tmp = self._copy_to_tmp(base_img_abs)
                    try:
                        left_pt = sht.range((current_row, 1)).left
                        top_pt = sht.range((current_row, 1)).top

                        pic = sht.pictures.add(str(tmp), left=left_pt, top=top_pt)
                        pic.width = pic_w_pt
                        pic.height = pic_h_pt

                        # ------------------------------
                        # calibration (px space)
                        # ------------------------------
                        cal_meta = meta.get("calibration", {}) or {}
                        cal = Calibration(
                            scale=float(cal_meta.get("scale", 1.0)),
                            off_x=float(cal_meta.get("off_x", 0.0)),
                            off_y=float(cal_meta.get("off_y", 0.0)),
                        )

                        rects = meta.get("rects_img_px", []) or []
                        log.debug(
                            "excel draw rects=%d img_px=%dx%d scale=%.6f off=(%.2f,%.2f)",
                            len(rects), img_w_px, img_h_px,
                            cal.scale, cal.off_x, cal.off_y
                        )

                        for i, r in enumerate(rects):
                            # --- apply calibration (px) ---
                            x_px = r["x"] * cal.scale + cal.off_x
                            y_px = r["y"] * cal.scale + cal.off_y
                            w_px = r["w"] * cal.scale
                            h_px = r["h"] * cal.scale

                            # --- px → ratio ---
                            rx = x_px / img_w_px
                            ry = y_px / img_h_px
                            rw = w_px / img_w_px
                            rh = h_px / img_h_px

                            # --- ratio → excel ---
                            left = pic.left + rx * pic.width
                            top = pic.top + ry * pic.height
                            width = rw * pic.width
                            height = rh * pic.height

                            shp = sht.api.Shapes.AddShape(
                                1, left, top, width, height
                            )
                            shp.Fill.Visible = False
                            shp.Line.Visible = True
                            shp.Line.ForeColor.RGB = self._rgb_from_hex(
                                r.get("color", "#FF3B30")
                            )
                            shp.Line.Weight = max(1.0, float(r.get("stroke", 2)))
                            shp.ZOrder(0)

                            log.debug(
                                "[EXCEL] rect%d px=(%.1f,%.1f,%.1f,%.1f) ratio=(%.4f,%.4f)",
                                i, x_px, y_px, w_px, h_px, rx, ry
                            )

                        used_cols = max(
                            3, math.ceil(pic.width / self.PT_PER_COL) + 1
                        )
                        comment_col = used_cols + 1
                        for col in range(1, comment_col + 3):
                            sht.range(1, col).column_width = 8.43
                        sht.range((current_row, comment_col)).value = comment

                        current_row += max(
                            15, int((pic.height or 240) / 12) + 3
                        )

                    finally:
                        tmp.unlink(missing_ok=True)

                book.save(str(out))
            finally:
                book.close()
        finally:
            app.display_alerts = prev_disp
            app.screen_updating = prev_upd
            app.quit()

        return out

    # --------------------------------------------------
    def _resolve_title(self, item) -> str:
        meta = getattr(item, "meta_raw", {}) or {}
        if meta.get("display_title"):
            return str(meta["display_title"])
        img = meta.get("image_path") or ""
        return Path(img).stem if img else str(getattr(item, "title", ""))

    def _copy_to_tmp(self, src: Path) -> Path:
        tmpdir = Path(tempfile.mkdtemp(prefix="capexp_"))
        dst = tmpdir / src.name
        shutil.copy2(src, dst)
        return dst

    def _rgb_from_hex(self, hexstr: str) -> int:
        s = hexstr.lstrip("#")
        if len(s) == 3:
            s = "".join(c * 2 for c in s)
        r = int(s[0:2], 16)
        g = int(s[2:4], 16)
        b = int(s[4:6], 16)
        return r + g * 256 + b * 65536
