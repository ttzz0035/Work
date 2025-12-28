from __future__ import annotations

import re
from pathlib import Path
from typing import Optional, Tuple, List

from wiki_module.bookstack_api import BookStackApi
from wiki_module.fs_repo import LocalWikiRepo
from wiki_module.models import LocalPage

# =========================================================
# Regex
# =========================================================

_METADATA_BLOCK_RE = re.compile(
    r"""
    ^---\s*\n
    \{\s*
    (?:.|\n)*?
    \}\s*\n
    ---\s*$
    """,
    re.MULTILINE | re.VERBOSE,
)

_HTML_IMG_RE = re.compile(
    r"<img\s+[^>]*src\s*=\s*(['\"]?)(?P<src>[^'\"\s>]+)\1[^>]*>",
    re.IGNORECASE,
)

_MD_IMG_RE = re.compile(
    r"!\[(?P<alt>[^\]]*)\]\((?P<src>[^)]+)\)"
)

_TABLE_ROW_RE = re.compile(r"^\s*\|.*\|\s*$")
_TABLE_SEP_RE = re.compile(r"^\s*\|[\s:\-]+\|\s*$")

# =========================================================
# Path helpers（UI import 契約）
# =========================================================

def repo_root(repo: LocalWikiRepo) -> Path:
    root = getattr(repo, "repo_root", None)
    if not root:
        raise AttributeError("LocalWikiRepo.repo_root not found")
    return Path(root).resolve()


def wiki_root(repo: LocalWikiRepo) -> Path:
    return repo_root(repo) / "wiki"


def infer_book_chapter_from_path(path: str) -> Tuple[str, str]:
    p = Path(path).as_posix()
    parts = p.split("/")
    try:
        idx = parts.index("books")
    except ValueError:
        return "", ""
    book = parts[idx + 1] if len(parts) > idx + 1 else ""
    chapter = parts[idx + 2] if len(parts) > idx + 2 else ""
    return book, chapter

# =========================================================
# Preview helpers
# =========================================================

def strip_metadata_blocks_for_preview(md: str, logger) -> str:
    if not md:
        return md
    before = len(md)
    md2 = _METADATA_BLOCK_RE.sub("", md)
    after = len(md2)
    if before != after:
        logger.info("[PREVIEW][META] removed metadata block (%s -> %s)", before, after)
    return md2


def _normalize_src(src: str) -> str:
    return (src or "").replace("\\", "/").lstrip("/")


def _resolve_image_for_preview(
    src: str,
    repo: LocalWikiRepo,
    logger,
) -> Optional[str]:
    s = _normalize_src(src)

    if not s.lower().endswith(".png"):
        s = f"{s}.png"

    if not s.startswith("wiki/"):
        s = f"wiki/{s}"

    disk = repo_root(repo) / s[len("wiki/"):]

    logger.info("[PREVIEW][RESOLVE] path=%s", s)
    logger.info("[PREVIEW][RESOLVE] check disk=%s", disk)

    if disk.exists():
        logger.info("[PREVIEW][RESOLVE] FOUND -> %s", s)
        return s

    logger.info("[PREVIEW][RESOLVE] NOT FOUND")
    return None

# =========================================================
# Image replacers
# =========================================================

class HtmlImgReplacer:
    def __init__(self, repo: LocalWikiRepo, logger):
        self.repo = repo
        self.logger = logger

    def __call__(self, match: re.Match) -> str:
        src = match.group("src") or ""
        new_src = _resolve_image_for_preview(src, self.repo, self.logger)
        if new_src:
            self.logger.info("[PREVIEW][HTML] rewrite %s -> %s", src, new_src)
            return f"![]({new_src})"
        return match.group(0)


class MdImgReplacer:
    def __init__(self, repo: LocalWikiRepo, logger):
        self.repo = repo
        self.logger = logger

    def __call__(self, match: re.Match) -> str:
        alt = match.group("alt") or ""
        src = match.group("src") or ""
        new_src = _resolve_image_for_preview(src, self.repo, self.logger)
        if new_src:
            self.logger.info("[PREVIEW][MD] rewrite %s -> %s", src, new_src)
            return f"![{alt}]({new_src})"
        return match.group(0)

# =========================================================
# Table normalize（Flet Markdown 対応・決定版）
# =========================================================

def normalize_tables(md: str, logger) -> str:
    if not md:
        return md

    lines = md.splitlines()
    out: List[str] = []

    table_lines: List[str] = []
    in_table = False

    def flush_table():
        if not table_lines:
            return
        # テーブル内の空行を全削除（これが決定打）
        cleaned = [l for l in table_lines if l.strip() != ""]
        out.extend(cleaned)
        table_lines.clear()

    for line in lines:
        if _TABLE_ROW_RE.match(line):
            if not in_table:
                # テーブル開始前に必ず空行
                if out and out[-1].strip() != "":
                    out.append("")
                in_table = True
            table_lines.append(line.rstrip())
            continue

        if in_table:
            # 空行は一旦無視（後で削除）
            if line.strip() == "":
                table_lines.append("")
                continue

            # テーブル終了
            flush_table()
            in_table = False

        out.append(line.rstrip())

    if in_table:
        flush_table()

    logger.info("[PREVIEW][TABLE] normalized (markdown-safe)")
    return "\n".join(out)

# =========================================================
# Main entry（互換シグネチャ）
# =========================================================

def render_markdown_for_flet(
    lp: LocalPage,
    repo: LocalWikiRepo,
    api: BookStackApi,
    logger,
) -> str:
    logger.info("========== [PREVIEW] START ==========")

    md = lp.content or ""
    logger.info("[PREVIEW][INPUT] len=%s", len(md))

    md = strip_metadata_blocks_for_preview(md, logger)

    md = _HTML_IMG_RE.sub(HtmlImgReplacer(repo, logger), md)
    md = _MD_IMG_RE.sub(MdImgReplacer(repo, logger), md)

    md = normalize_tables(md, logger)

    logger.info("[PREVIEW][OUTPUT] len=%s", len(md))
    logger.info("========== [PREVIEW] END ==========")
    return md
