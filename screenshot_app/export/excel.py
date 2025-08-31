# export/excel.py
from __future__ import annotations
from pathlib import Path
import logging, tempfile, shutil, math

import xlwings as xw
from PIL import Image  # Pillow

from .base import BaseExporter, ExportOptions, DEFAULT_TITLE
from .registry import register
from core.model import ExportBundle

log = logging.getLogger("export.excel")


@register
class ExcelExporter(BaseExporter):
    """
    Excel エクスポータ（ヘッドレス、画像と矩形は別描画、左=画像/右=コメント）
    - Excelは非表示(visible=False)で起動し、画面更新に依存しないヘッドレス動作
    - 画像は元PNGを貼付、矩形はShapeとして別描画（色/太さ反映）
    - レイアウトは画像実寸（px）→ポイント換算でスケールし、画像幅に応じて
      コメント開始列を自動決定（概算: 1列 ≒ 48pt）
    """
    name = "excel"
    ext = ".xlsx"

    # 画像サイズ→ポイント換算
    # 96dpi前提: 1px ≒ 0.75pt
    PX_TO_PT = 0.75
    # 画像の最大サイズ（ポイント）: 超える場合は等比縮小
    MAX_IMG_W_PT = 720.0   # お好みで（例: 10インチ）
    MAX_IMG_H_PT = 540.0   # お好みで（例: 7.5インチ）
    # コメントと画像の間隔（ポイント）
    GAP_PT = 18.0
    # 概算: 1列 ≒ 48pt（既定フォント/ズームの便宜的な近似）
    PT_PER_COL = 48.0

    def export_bundle(self, bundle: ExportBundle, options: ExportOptions) -> Path:
        sheet_title = (options.title or bundle.title or DEFAULT_TITLE).strip()
        out = options.filename or (bundle.folder / f"captures_export{self.ext}")
        out.parent.mkdir(parents=True, exist_ok=True)

        # ---- ヘッドレス・アプリ確保 ----
        app = xw.App(visible=False, add_book=False)
        # ちらつき/ダイアログ抑止
        prev_upd = app.screen_updating
        prev_disp = app.display_alerts
        try:
            app.screen_updating = False
            app.display_alerts = False

            # ブック確保
            if out.exists():
                book = app.books.open(str(out))
            else:
                book = app.books.add()
                book.save(str(out))

            try:
                # 既存シート削除 → 再作成
                self.remove_existing_by_title(
                    targets=book.sheets,
                    match_title=sheet_title,
                    get_name=lambda s: s.name,
                    delete=lambda s: s.delete(),
                )
                sht = book.sheets.add(name=sheet_title, after=book.sheets[-1] if book.sheets else None)

                current_row = 1
                for it in bundle.items:
                    meta = getattr(it, "meta_raw", {}) or {}
                    title = self._resolve_title(it)
                    comment = it.comment or "(no comment)"

                    # 画像パス（元PNG）
                    base_img_rel = meta.get("image_path")
                    if not base_img_rel:
                        # タイトルのみ配置してスキップ
                        sht.range((current_row, 1)).value = title
                        current_row += 3
                        continue
                    base_img_abs = (bundle.folder / base_img_rel) if not Path(base_img_rel).is_absolute() else Path(base_img_rel)
                    if not base_img_abs.exists():
                        sht.range((current_row, 1)).value = f"{title} (image not found)"
                        current_row += 3
                        continue

                    # ---- 画像の実寸(px)を取得 → ポイントに換算 → 最大値でクリップ ----
                    with Image.open(str(base_img_abs)) as im:
                        w_px, h_px = im.size
                    w_pt_raw = w_px * self.PX_TO_PT
                    h_pt_raw = h_px * self.PX_TO_PT
                    scale = min(1.0, self.MAX_IMG_W_PT / max(w_pt_raw, 1.0), self.MAX_IMG_H_PT / max(h_pt_raw, 1.0))
                    w_pt = w_pt_raw * scale
                    h_pt = h_pt_raw * scale

                    # ---- タイトル（画像の1行上） ----
                    sht.range((current_row, 1)).value = title
                    current_row += 1

                    # ---- 画像貼付（一時ファイル経由） ----
                    tmp = self._copy_to_tmp(base_img_abs)
                    try:
                        # A列のセル左上座標（ポイント）
                        left_pt = sht.range((current_row, 1)).left
                        top_pt = sht.range((current_row, 1)).top

                        pic = sht.pictures.add(str(tmp), left=left_pt, top=top_pt)
                        # 実寸ベースへ縮尺
                        if pic.width:
                            pic.width = w_pt
                        if pic.height:
                            pic.height = h_pt

                        # ---- 矩形を別描画（領域座標 → 画像スケールへ） ----
                        region = meta.get("region", {}) or {}
                        r_w = float(region.get("width") or 1.0)
                        r_h = float(region.get("height") or 1.0)
                        sx = float(pic.width) / r_w
                        sy = float(pic.height) / r_h
                        img_left = pic.left
                        img_top = pic.top

                        rects = meta.get("rects", []) or []
                        for r in rects:
                            try:
                                rx = float(r.get("x", 0.0)); ry = float(r.get("y", 0.0))
                                rw = float(r.get("w", 0.0)); rh = float(r.get("h", 0.0))
                                col_hex = str(r.get("color", "#FF3B30"))
                                stroke = max(1.0, float(r.get("stroke", 2)))

                                left = img_left + rx * sx
                                top  = img_top  + ry * sy
                                width  = rw * sx
                                height = rh * sy

                                shp = sht.api.Shapes.AddShape(1, left, top, width, height)  # msoShapeRectangle=1
                                shp.Fill.Visible = False
                                shp.Line.Visible = True
                                shp.Line.ForeColor.RGB = self._rgb_from_hex(col_hex)
                                shp.Line.Weight = stroke
                                shp.ZOrder(0)  # 前面
                            except Exception as ex:
                                log.warning("Rect draw failed: %s", ex)

                        # ---- コメントを画像の“右側”に配置 ----
                        # コメント開始列 = 画像幅ポイント / 列ポイント ≒ 列数（切り上げ） + 余白列
                        used_cols = max(3, math.ceil(pic.width / self.PT_PER_COL) + 1)
                        comment_col = used_cols + 1  # 1列スペーサのさらに右
                        # ざっくり列幅を整える（全列を8.43にしておくと単位が安定）
                        for col in range(1, comment_col + 3):
                            sht.range(1, col).column_width = 8.43
                        sht.range((current_row, comment_col)).value = comment

                        # ---- 次の行へ：画像高さに応じて行送り（ポイント→行換算ざっくり）----
                        # 標準行高 ≒ 15pt 前後 → 行数 ≒ (高さ/12)+α
                        current_row += max(15, int((pic.height or 240) / 12) + 3)

                    finally:
                        try:
                            tmp.unlink(missing_ok=True)
                        except Exception:
                            pass

                book.save(str(out))
            finally:
                book.close()
        finally:
            app.display_alerts = prev_disp
            app.screen_updating = prev_upd
            app.quit()

        return out

    # ---- helpers ----
    def _resolve_title(self, item) -> str:
        meta = getattr(item, "meta_raw", {}) or {}
        disp = meta.get("display_title")
        if disp:
            return str(disp)
        img = meta.get("image_path") or ""
        if img:
            try:
                return Path(img).stem
            except Exception:
                pass
        return str(getattr(item, "title", ""))

    def _copy_to_tmp(self, src: Path) -> Path:
        tmpdir = Path(tempfile.mkdtemp(prefix="capexp_"))
        dst = tmpdir / src.name
        shutil.copy2(src, dst)
        return dst

    def _rgb_from_hex(self, hexstr: str) -> int:
        s = hexstr.strip()
        if s.startswith("#"):
            s = s[1:]
        if len(s) == 3:
            s = "".join(c * 2 for c in s)
        r = int(s[0:2], 16)
        g = int(s[2:4], 16)
        b = int(s[4:6], 16)
        # VBAのRGB順（B*65536 + G*256 + R）ではなく、xlwingsのShape.Line.ForeColor.RGBは
        # VBAと同じ整数。VBAのRGB(r,g,b)は r + g*256 + b*65536 を返す。
        return r + g * 256 + b * 65536
