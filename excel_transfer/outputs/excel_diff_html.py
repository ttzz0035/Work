from __future__ import annotations

import json
import html
from pathlib import Path
from typing import Dict, Any, List
import logging


# =========================================================
# Logger
# =========================================================
def get_logger(name: str) -> logging.Logger:
    logger = logging.getLogger(name)
    if logger.handlers:
        return logger

    logger.setLevel(logging.INFO)
    handler = logging.StreamHandler()
    formatter = logging.Formatter(
        "%(asctime)s - %(levelname)s - [%(name)s] %(message)s"
    )
    handler.setFormatter(formatter)
    logger.addHandler(handler)
    logger.propagate = False
    return logger


logger = get_logger("excel_diff_html")


# =========================================================
# HTML Builder
# =========================================================
class ExcelDiffHtmlReport:
    def __init__(self, diff_data: Dict[str, Any]):
        self.data = diff_data

    # -----------------------------------------------------
    # Public
    # -----------------------------------------------------
    def build_html(self) -> str:
        logger.info("[HTML] build start")
        parts: List[str] = []
        parts.append(self._html_header())
        parts.append(self._build_meta())
        parts.append(self._build_summary())
        parts.append(self._build_cell_diff())
        parts.append(self._build_shape_diff())
        parts.append(self._html_footer())
        logger.info("[HTML] build end")
        return "\n".join(parts)

    # -----------------------------------------------------
    # Sections
    # -----------------------------------------------------
    def _build_meta(self) -> str:
        meta = self.data.get("meta", {})
        return f"""
<h2>Compare Meta</h2>
<table class="meta">
<tr><th>File A</th><td>{html.escape(meta.get("file_a",""))}</td></tr>
<tr><th>File B</th><td>{html.escape(meta.get("file_b",""))}</td></tr>
<tr><th>Range A</th><td>{meta.get("range_a","")}</td></tr>
<tr><th>Range B</th><td>{meta.get("range_b","")}</td></tr>
<tr><th>Base File</th><td>{meta.get("base_file","")}</td></tr>
<tr><th>Compare Formula</th><td>{meta.get("compare_formula")}</td></tr>
<tr><th>Compare Shapes</th><td>{meta.get("compare_shapes")}</td></tr>
</table>
"""

    def _build_summary(self) -> str:
        s = self.data.get("summary", {})
        return f"""
<h2>Summary</h2>
<ul>
<li>Cell Modified: {s.get("cell_mod_count",0)}</li>
<li>Shape Diff: {s.get("shape_diff_count",0)}</li>
<li>Base File: {s.get("base_file","")}</li>
</ul>
"""

    # -----------------------------------------------------
    # Cell Diff
    # -----------------------------------------------------
    def _build_cell_diff(self) -> str:
        rows: List[str] = []

        for idx, d in enumerate(self.data.get("diff_cells", []), start=1):
            try:
                sheet = html.escape(str(d.get("sheet", "")))
                r = d["row"]
                c = d["col"]
                t = d.get("type", "")
                base = d.get("base", "")
                va = html.escape(d.get("value_a", ""))
                vb = html.escape(d.get("value_b", ""))
            except KeyError as e:
                logger.error(f"[SKIP] invalid diff_cell structure: {d} err={e}")
                continue

            rows.append(
                f"<tr class='diff-{t.lower()}'>"
                f"<td>{idx}</td>"
                f"<td>{sheet}</td>"
                f"<td>{r}</td>"
                f"<td>{c}</td>"
                f"<td>{t}</td>"
                f"<td>{va}</td>"
                f"<td>{vb}</td>"
                f"<td>{base}</td>"
                f"</tr>"
            )

        body = "\n".join(rows) if rows else "<tr><td colspan='8'>No diff</td></tr>"

        return f"""
<h2>Cell Differences</h2>
<table class="diff">
<tr>
  <th>No</th>
  <th>Sheet</th>
  <th>Row</th>
  <th>Col</th>
  <th>Type</th>
  <th>Value A</th>
  <th>Value B</th>
  <th>Base</th>
</tr>
{body}
</table>
"""

    # -----------------------------------------------------
    # Shape Diff（ADD / DEL / GEOM / TEXT 対応）
    # -----------------------------------------------------
    def _build_shape_diff(self) -> str:
        rows: List[str] = []

        for idx, d in enumerate(self.data.get("diff_shapes", []), start=1):
            t = d.get("type", "")
            sheet = html.escape(str(d.get("sheet", "")))
            name = html.escape(str(d.get("name", "")))

            detail_a = ""
            detail_b = ""

            if t == "SHAPE_GEOM":
                detail_a = html.escape(json.dumps(d.get("a", {}), ensure_ascii=False))
                detail_b = html.escape(json.dumps(d.get("b", {}), ensure_ascii=False))

            elif t == "SHAPE_TEXT":
                detail_a = html.escape(d.get("text_a", ""))
                detail_b = html.escape(d.get("text_b", ""))

            elif t in ("SHAPE_ADD", "SHAPE_DEL"):
                detail_a = "-"
                detail_b = "-"

            rows.append(
                f"<tr class='diff-shape diff-{t.lower()}'>"
                f"<td>{idx}</td>"
                f"<td>{sheet}</td>"
                f"<td>{name}</td>"
                f"<td>{t}</td>"
                f"<td>{detail_a}</td>"
                f"<td>{detail_b}</td>"
                f"</tr>"
            )

        body = "\n".join(rows) if rows else "<tr><td colspan='6'>No shape diff</td></tr>"

        return f"""
<h2>Shape Differences</h2>
<table class="diff">
<tr>
  <th>No</th>
  <th>Sheet</th>
  <th>Name</th>
  <th>Type</th>
  <th>Detail A</th>
  <th>Detail B</th>
</tr>
{body}
</table>
"""

    # -----------------------------------------------------
    # HTML Frame
    # -----------------------------------------------------
    def _html_header(self) -> str:
        return """
<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8"/>
<title>Excel Diff Report</title>
<style>
body { font-family: Consolas, monospace; font-size: 13px; }
table { border-collapse: collapse; margin-bottom: 20px; width: 100%; }
th, td { border: 1px solid #aaa; padding: 4px 8px; vertical-align: top; }
th { background: #eee; }

.diff-mod { background: #fff3cd; }
.diff-add { background: #d4edda; }
.diff-del { background: #f8d7da; }

.diff-shape { background: #e2d6f3; }
.diff-shape.diff-shape_geom { background: #f6d6d6; }
.diff-shape.diff-shape_text { background: #fff0b3; }

.meta th { text-align: left; width: 160px; }
</style>
</head>
<body>
<h1>Excel Diff Report</h1>
"""

    def _html_footer(self) -> str:
        return """
</body>
</html>
"""


# =========================================================
# Entry
# =========================================================
def generate_html_report(json_path: Path, out_path: Path) -> None:
    logger.info(f"[LOAD] {json_path}")
    with json_path.open("r", encoding="utf-8") as f:
        data = json.load(f)

    report = ExcelDiffHtmlReport(data)
    html_text = report.build_html()

    logger.info(f"[WRITE] {out_path}")
    out_path.write_text(html_text, encoding="utf-8")
