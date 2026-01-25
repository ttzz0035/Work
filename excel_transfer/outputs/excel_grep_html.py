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


logger = get_logger("excel_grep_html")


# =========================================================
# HTML Builder
# =========================================================
class ExcelGrepHtmlReport:
    def __init__(self, grep_data: Dict[str, Any], labels: Dict[str, str]):
        self.data = grep_data
        self.labels = labels

    # -----------------------------------------------------
    # Public
    # -----------------------------------------------------
    def build_html(self) -> str:
        logger.info("[HTML] build start")
        parts: List[str] = []
        parts.append(self._html_header())
        parts.append(self._build_meta())
        parts.append(self._build_summary())
        parts.append(self._build_results())
        parts.append(self._html_footer())
        logger.info("[HTML] build end")
        return "\n".join(parts)

    # -----------------------------------------------------
    # Sections
    # -----------------------------------------------------
    def _build_meta(self) -> str:
        meta = self.data.get("meta", {}) or {}
        L = self.labels

        def esc(v: Any) -> str:
            return html.escape("" if v is None else str(v))

        off = meta.get("offset", [0, 0])
        off_r = off[0] if len(off) > 0 else 0
        off_c = off[1] if len(off) > 1 else 0

        return f"""
<h2>{L["grep_html_meta_title"]}</h2>
<table class="meta">
<tr><th>{L["grep_html_meta_keyword"]}</th><td>{esc(meta.get("search_pattern"))}</td></tr>
<tr><th>{L["grep_html_meta_use_regex"]}</th><td>{esc(meta.get("use_regex"))}</td></tr>
<tr><th>{L["grep_html_meta_file_regex"]}</th><td>{esc(meta.get("file_name_regex"))}</td></tr>
<tr><th>{L["grep_html_meta_sheet_regex"]}</th><td>{esc(meta.get("sheet_name_regex"))}</td></tr>
<tr><th>{L["grep_html_meta_offset_row"]}</th><td>{off_r}</td></tr>
<tr><th>{L["grep_html_meta_offset_col"]}</th><td>{off_c}</td></tr>
<tr><th>{L["grep_html_meta_replace_pattern"]}</th><td>{esc(meta.get("replace_pattern"))}</td></tr>
<tr><th>{L["grep_html_meta_replace_mode"]}</th><td>{esc(meta.get("replace_mode"))}</td></tr>
</table>
"""

    def _build_summary(self) -> str:
        files = self.data.get("files", []) or []
        L = self.labels

        file_cnt = 0
        sheet_cnt = 0
        hit_cnt = 0
        checked_cnt = 0
        diff_cnt = 0

        for f in files:
            file_cnt += 1
            for s in f.get("sheets", []) or []:
                sheet_cnt += 1
                for it in s.get("items", []) or []:
                    hit_cnt += 1
                    if it.get("checked"):
                        checked_cnt += 1
                    if it.get("before") != it.get("after"):
                        diff_cnt += 1

        return f"""
<h2>{L["grep_html_summary_title"]}</h2>
<ul>
<li>{L["grep_html_summary_files"]}: {file_cnt}</li>
<li>{L["grep_html_summary_sheets"]}: {sheet_cnt}</li>
<li>{L["grep_html_summary_hits"]}: {hit_cnt}</li>
<li>{L["grep_html_summary_checked"]}: {checked_cnt}</li>
<li>{L["grep_html_summary_diff"]}: {diff_cnt}</li>
</ul>
"""

    # -----------------------------------------------------
    # Results
    # -----------------------------------------------------
    def _build_results(self) -> str:
        rows: List[str] = []
        idx = 0
        L = self.labels

        for f in self.data.get("files", []) or []:
            fpath = html.escape(str(f.get("path", "")))
            for s in f.get("sheets", []) or []:
                sname = html.escape(str(s.get("name", "")))
                for it in s.get("items", []) or []:
                    idx += 1

                    hit = it.get("hit", {})
                    tgt = it.get("target", {})

                    hr = hit.get("row", "")
                    hc = hit.get("col", "")
                    tr = tgt.get("row", "")
                    tc = tgt.get("col", "")

                    before = "" if it.get("before") is None else str(it.get("before"))
                    after = "" if it.get("after") is None else str(it.get("after"))

                    esc_before = html.escape(before)
                    esc_after = html.escape(after)

                    checked = bool(it.get("checked", False))
                    checked_mark = "✔" if checked else ""

                    diff_cls = "diff-yes" if before != after else "diff-no"

                    rows.append(
                        f"<tr class='{diff_cls}'>"
                        f"<td>{idx}</td>"
                        f"<td class='mono'>{fpath}</td>"
                        f"<td>{sname}</td>"
                        f"<td class='mono'>R{hr}C{hc}</td>"
                        f"<td class='mono'>R{tr}C{tc}</td>"
                        f"<td class='pre'>{esc_before}</td>"
                        f"<td class='pre'>{esc_after}</td>"
                        f"<td class='center'>{checked_mark}</td>"
                        f"</tr>"
                    )

        body = "\n".join(rows) if rows else f"<tr><td colspan='8'>{L['grep_html_no_result']}</td></tr>"

        return f"""
<h2>{L["grep_html_results_title"]}</h2>
<table class="diff">
<tr>
  <th>{L["grep_html_col_no"]}</th>
  <th>{L["grep_html_col_file"]}</th>
  <th>{L["grep_html_col_sheet"]}</th>
  <th>{L["grep_html_col_hit"]}</th>
  <th>{L["grep_html_col_target"]}</th>
  <th>{L["grep_html_col_before"]}</th>
  <th>{L["grep_html_col_after"]}</th>
  <th>{L["grep_html_col_checked"]}</th>
</tr>
{body}
</table>
"""

    # -----------------------------------------------------
    # HTML Frame
    # -----------------------------------------------------
    def _html_header(self) -> str:
        title = self.labels["grep_html_title"]
        return f"""
<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8"/>
<title>{title}</title>
<style>
body {{ font-family: Consolas, monospace; font-size: 13px; }}
table {{ border-collapse: collapse; margin-bottom: 20px; width: 100%; }}
th, td {{ border: 1px solid #aaa; padding: 4px 8px; vertical-align: top; }}
th {{ background: #eee; }}

.meta th {{ text-align: left; width: 180px; }}

.mono {{ white-space: nowrap; }}
.pre {{ white-space: pre-wrap; word-break: break-word; }}
.center {{ text-align: center; }}

.diff-yes {{ background: #fff3cd; }}
.diff-no  {{ background: #f8f9fa; }}
</style>
</head>
<body>
<h1>{title}</h1>
"""

    def _html_footer(self) -> str:
        return """
</body>
</html>
"""


# =========================================================
# Entry
# =========================================================
def generate_grep_html_report(json_path: Path, out_path: Path, labels: Dict[str, str]) -> None:
    logger.info(f"[LOAD] {json_path}")
    with json_path.open("r", encoding="utf-8") as f:
        data = json.load(f)

    report = ExcelGrepHtmlReport(data, labels)
    html_text = report.build_html()

    logger.info(f"[WRITE] {out_path}")
    out_path.write_text(html_text, encoding="utf-8")
