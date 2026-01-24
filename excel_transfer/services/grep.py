# ============================================
# services/grep.py
# ============================================
from __future__ import annotations

import os
import re
import json
import xlwings as xw
from typing import Tuple, List, Dict, Any
from collections import defaultdict
import shutil
from datetime import datetime

from models.dto import GrepRequest, LogFn
from utils.search_utils import compile_matcher

EXCEL_EXTS = (".xlsx", ".xlsm", ".xlsb", ".xls")


# =================================================
# replace core
# =================================================
def apply_replace(items: List[Dict[str, Any]], logger) -> None:
    import shutil
    from datetime import datetime

    logger.info(f"[GREP][REPLACE] start items={len(items)}")

    grouped = defaultdict(lambda: defaultdict(list))
    for it in items:
        if not it.get("checked"):
            continue
        grouped[it["path"]][it["sheet"]].append(it)

    for file_path, sheets in grouped.items():
        app = None
        book = None

        # ==========================
        # backup（★ 追加）
        # ==========================
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        base, ext = os.path.splitext(file_path)
        backup_path = f"{base}.bak_{ts}{ext}"

        try:
            shutil.copy2(file_path, backup_path)
            logger.info(f"[GREP][BACKUP] created: {backup_path}")
        except Exception as e:
            logger.error(f"[GREP][BACKUP] failed: {file_path} -> {backup_path} ({e})")
            raise

        try:
            logger.info(f"[GREP][REPLACE] open file={file_path}")
            app = xw.App(visible=False, add_book=False)
            book = app.books.open(file_path, read_only=False)

            for sheet_name, sheet_items in sheets.items():
                sht = book.sheets[sheet_name]
                logger.info(
                    f"[GREP][REPLACE] sheet={sheet_name} items={len(sheet_items)}"
                )

                for it in sheet_items:
                    r = it["target"]["row"]
                    c = it["target"]["col"]
                    before = it["before"]
                    after = it["after"]

                    logger.info(
                        f"[GREP][REPLACE] APPLY {sheet_name}!R{r}C{c} "
                        f"'{before}' -> '{after}'"
                    )

                    sht.cells(r, c).value = after

            logger.info(f"[GREP][REPLACE] save file={file_path}")
            book.save()

            # 保存後確認
            for sheet_name, sheet_items in sheets.items():
                sht = book.sheets[sheet_name]
                for it in sheet_items:
                    r = it["target"]["row"]
                    c = it["target"]["col"]
                    actual = sht.cells(r, c).value
                    logger.info(
                        f"[GREP][REPLACE] VERIFY {sheet_name}!R{r}C{c} value='{actual}'"
                    )

        finally:
            try:
                if book:
                    book.close()
                    logger.info(f"[GREP][REPLACE] close book={file_path}")
            except Exception as e:
                logger.error(f"[GREP][REPLACE] close error: {e}")

            try:
                if app:
                    app.kill()
                    logger.info("[GREP][REPLACE] excel process killed")
            except Exception as e:
                logger.error(f"[GREP][REPLACE] app kill error: {e}")

    logger.info("[GREP][REPLACE] done")

# =================================================
# helpers
# =================================================
def _find_excel_files(root: str, file_name_pattern: re.Pattern | None) -> List[str]:
    hits: List[str] = []
    for dp, _, fns in os.walk(root):
        for fn in fns:
            if not fn.lower().endswith(EXCEL_EXTS):
                continue
            if file_name_pattern and not file_name_pattern.search(fn):
                continue
            hits.append(os.path.join(dp, fn))
    return hits


def _to_2d(vals: Any) -> List[List[Any]]:
    if vals is None:
        return []
    if not isinstance(vals, list):
        return [[vals]]
    if len(vals) == 0:
        return []
    if isinstance(vals[0], list):
        return vals
    return [vals]


def _build_after(before: str, req: GrepRequest) -> str:
    if not req.replace_enabled:
        return before

    if not req.use_regex:
        return req.replace_pattern

    try:
        return re.sub(req.keyword, req.replace_pattern, before)
    except re.error as e:
        raise ValueError(f"置換正規表現エラー: {e}")


# =================================================
# run_grep
# =================================================
def run_grep(
    req: GrepRequest,
    ctx,
    logger,
    append_log: LogFn,
) -> Tuple[str, int]:
    append_log("=== Grep開始 ===")

    matcher = compile_matcher(req.keyword, req.use_regex, req.ignore_case)

    total_hits = 0
    items: List[Dict[str, Any]] = []
    result = {"files": []}

    for path in _find_excel_files(req.root_dir, None):
        app = book = None
        try:
            app = xw.App(visible=False, add_book=False)
            book = app.books.open(path, read_only=True)

            file_entry = {"path": path, "sheets": []}

            for sht in book.sheets:
                vr = sht.used_range
                vals = _to_2d(vr.value)
                sheet_entry = {"name": sht.name, "items": []}

                for r0, row in enumerate(vals):
                    for c0, v in enumerate(row if isinstance(row, list) else [row]):
                        if not matcher(v):
                            continue

                        match_r = vr.row + r0
                        match_c = vr.column + c0

                        hit_r = match_r - req.offset_col
                        hit_c = match_c + req.offset_row
                        if hit_r < 1 or hit_c < 1:
                            continue

                        cell = sht.cells(hit_r, hit_c)
                        before = "" if cell.value is None else str(cell.value)
                        after = (
                            re.sub(req.keyword, req.replace_pattern, before)
                            if req.replace_enabled and req.use_regex
                            else req.replace_pattern if req.replace_enabled
                            else before
                        )

                        item = {
                            "id": total_hits,
                            "path": path,
                            "sheet": sht.name,
                            "match": {"row": match_r, "col": match_c},
                            "hit": {"row": hit_r, "col": hit_c},
                            "target": {"row": hit_r, "col": hit_c},
                            "before": before,
                            "after": after,
                            "checked": True,
                        }

                        sheet_entry["items"].append(item)
                        items.append(item)
                        total_hits += 1

                if sheet_entry["items"]:
                    file_entry["sheets"].append(sheet_entry)

            if file_entry["sheets"]:
                result["files"].append(file_entry)

        finally:
            if book:
                book.close()
            if app:
                app.kill()

    # JSON保存
    os.makedirs(ctx.output_dir, exist_ok=True)
    json_path = os.path.join(ctx.output_dir, "grep_result.json")
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(result, f, ensure_ascii=False, indent=2)

    append_log(f"[GREP] json saved: {json_path}")

    # ---- 置換フェーズ ----
    append_log(
        f"[GREP][REPLACE] phase enter "
        f"enabled={req.replace_enabled} "
        f"mode={getattr(req, 'replace_mode', None)} "
        f"items={len(items)}"
    )

    if req.replace_enabled:
        if req.replace_mode == "auto":
            append_log("[GREP][REPLACE][AUTO] applying replace")
            apply_replace(items, logger)

        elif req.replace_mode == "preview":
            accepted = getattr(req, "preview_accepted", False)
            append_log(
                f"[GREP][REPLACE][PREVIEW] accepted={accepted}"
            )
            if accepted:
                append_log("[GREP][REPLACE][PREVIEW] applying replace")
                apply_replace(items, logger)
            else:
                append_log("[GREP][REPLACE][PREVIEW] canceled")

        else:
            append_log(
                f"[GREP][REPLACE][ERROR] invalid replace_mode={req.replace_mode}"
            )
            raise ValueError(f"invalid replace_mode: {req.replace_mode}")

    append_log("=== Grep終了 ===")
    return json_path, total_hits
