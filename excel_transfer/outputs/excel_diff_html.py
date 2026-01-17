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

    def _build_cell_diff(self) -> str:
        rows = []
        for d in self.data.get("diff_cells", []):
            r = d["mark"]["row"]
            c = d["mark"]["col"]
            t = d["type"]
            base = d["mark"]["base"]
            rows.append(
                f"<tr class='diff-{t.lower()}'>"
                f"<td>R{r}C{c}</td>"
                f"<td>{t}</td>"
                f"<td>{base}</td>"
                f"</tr>"
            )

        body = "\n".join(rows) if rows else "<tr><td colspan='3'>No diff</td></tr>"

        return f"""
<h2>Cell Differences</h2>
<table class="diff">
<tr><th>Cell</th><th>Type</th><th>Base</th></tr>
{body}
</table>
"""

    def _build_shape_diff(self) -> str:
        rows = []
        for d in self.data.get("diff_shapes", []):
            name = html.escape(d.get("name", ""))
            t = d.get("type", "")
            a = d.get("a", {})
            b = d.get("b", {})
            rows.append(f"""
<tr class="diff-shape">
<td>{name}</td>
<td>{t}</td>
<td>{html.escape(str(a))}</td>
<td>{html.escape(str(b))}</td>
</tr>
""")

        body = "\n".join(rows) if rows else "<tr><td colspan='4'>No shape diff</td></tr>"

        return f"""
<h2>Shape Differences</h2>
<table class="diff">
<tr><th>Name</th><th>Type</th><th>A</th><th>B</th></tr>
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
table { border-collapse: collapse; margin-bottom: 20px; }
th, td { border: 1px solid #aaa; padding: 4px 8px; }
th { background: #eee; }

.diff-mod { background: #fff3cd; }
.diff-add { background: #d4edda; }
.diff-del { background: #f8d7da; }
.diff-shape { background: #e2d6f3; }

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


if __name__ == "__main__":
    import tkinter as tk
    from tkinter import filedialog

    root = tk.Tk()
    root.withdraw()  # 画面非表示

    logger.info("[UI] select diff json")

    json_path = filedialog.askopenfilename(
        title="Select diff JSON",
        filetypes=[("JSON files", "*.json")],
    )

    if not json_path:
        logger.info("[UI] canceled (json)")
        raise SystemExit(0)

    logger.info("[UI] select output html")

    out_path = filedialog.asksaveasfilename(
        title="Save HTML report",
        defaultextension=".html",
        filetypes=[("HTML files", "*.html")],
    )

    if not out_path:
        logger.info("[UI] canceled (html)")
        raise SystemExit(0)

    generate_html_report(Path(json_path), Path(out_path))
