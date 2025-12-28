# =========================================================
# fs_repo.py（V1FIX・Attachment=ID.png 固定・完全互換）
# =========================================================
from __future__ import annotations

import hashlib
import json
import re
from dataclasses import asdict
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict, List, Optional, Union

from wiki_module.bookstack_api import BookStackApi
from wiki_module.logger import AppLogger
from wiki_module.models import LocalConfig, LocalPageMeta, LocalPage


# =========================================================
# Utilities
# =========================================================
def compute_hash(text: str) -> str:
    h = hashlib.sha256()
    h.update((text or "").encode("utf-8", errors="ignore"))
    return h.hexdigest()


_INVALID_WIN_CHARS = r'<>:"/\\|?*\x00-\x1f'
_INVALID_WIN_RE = re.compile(f"[{_INVALID_WIN_CHARS}]")


def _now_iso() -> str:
    return datetime.now(timezone.utc).isoformat().replace("+00:00", "Z")


def _sanitize_segment(v: Any) -> str:
    s = str(v or "").strip()
    s = _INVALID_WIN_RE.sub("_", s)
    s = s.replace("\u200b", "")
    s = s.strip(" .")
    return s or "_"


# =========================================================
# Repo
# =========================================================
class LocalWikiRepo:
    """
    Local Wiki Repository（V1FIX最終・Attachment=.png固定）

    - Page仕様: 確定（変更禁止）
    - Attachment:
        * SyncEngine 互換: save_attachment(page_id, filename, data)
        * filename が数値なら attachment_id
        * 数値でなければ page_id を attachment_id とする
        * 保存名は {attachment_id}.png 固定
    - Repo は FS 管理のみ
    """

    def __init__(
        self,
        cfg_or_root: Union[LocalConfig, str],
        logger_name: str = "LocalWikiRepo",
    ) -> None:
        if isinstance(cfg_or_root, LocalConfig):
            self.repo_root = Path(cfg_or_root.repo_root)
            self.assets_root = cfg_or_root.get_assets_dir()
        else:
            self.repo_root = Path(cfg_or_root)
            self.assets_root = self.repo_root / "assets"

        self.books_root = self.repo_root / "books"
        self.attachments_root = self.assets_root / "attachments"
        self.gallery_root = self.assets_root / "gallery"

        self.logger = AppLogger(logger_name, "logs/local_repo.log", "INFO").get()

        self.books_root.mkdir(parents=True, exist_ok=True)
        self.attachments_root.mkdir(parents=True, exist_ok=True)
        self.gallery_root.mkdir(parents=True, exist_ok=True)

        # 互換必須
        self._api: Optional[BookStackApi] = None

    # =====================================================
    # API Injection
    # =====================================================
    def set_bookstack_api(self, api: BookStackApi) -> None:
        self._api = api
        self.logger.info("[REPO] BookStackApi injected")

    # =====================================================
    # Tree
    # =====================================================
    def build_tree_index(self) -> Dict[str, Dict[str, List[LocalPage]]]:
        tree: Dict[str, Dict[str, List[LocalPage]]] = {}

        if not self.books_root.exists():
            return tree

        for book_dir in self.books_root.iterdir():
            if not book_dir.is_dir():
                continue
            tree.setdefault(book_dir.name, {})

            for chap_dir in book_dir.iterdir():
                if not chap_dir.is_dir():
                    continue
                pages: List[LocalPage] = []
                for md in chap_dir.glob("*.md"):
                    try:
                        pages.append(self.load_page(str(md)))
                    except Exception as e:
                        self.logger.error("[TREE] load failed %s err=%s", md, e)
                tree[book_dir.name][chap_dir.name] = pages

        return tree

    # =====================================================
    # Page IO（確定）
    # =====================================================
    def load_page(self, path: str) -> LocalPage:
        p = Path(path)
        txt = p.read_text(encoding="utf-8")
        meta = LocalPageMeta(
            page_id=None,
            book="",
            chapter="",
            title=p.stem,
            updated_at=_now_iso(),
            remote_updated_at=None,
            synced_at=None,
            content_hash=compute_hash(txt),
        )
        return LocalPage(path=str(p), meta=meta, content=txt)

    def save_page(self, lp: LocalPage) -> None:
        p = Path(lp.path)
        p.parent.mkdir(parents=True, exist_ok=True)
        fm = json.dumps(asdict(lp.meta), ensure_ascii=False, indent=2)
        body = f"---\n{fm}\n---\n\n{lp.content or ''}\n"
        p.write_text(body, encoding="utf-8")
        self.logger.info("[REPO] page saved %s", p)

    def make_new_local_page(
        self,
        book: str,
        chapter: str,
        title: str,
        content: str,
    ) -> LocalPage:
        book_s = _sanitize_segment(book or "default")
        chapter_s = _sanitize_segment(chapter or "_no_chapter")
        title_s = _sanitize_segment(title or "Untitled")

        out = self.books_root / book_s / chapter_s / f"{title_s}.md"

        meta = LocalPageMeta(
            page_id=None,
            book=book,
            chapter=chapter,
            title=title,
            updated_at=_now_iso(),
            remote_updated_at=None,
            synced_at=None,
            content_hash=compute_hash(content),
        )

        lp = LocalPage(path=str(out), meta=meta, content=content)
        self.save_page(lp)
        return lp

    def save_page_from_remote(self, page_detail: Dict[str, Any]) -> LocalPage:
        page_id = int(page_detail["id"])
        page_name = _sanitize_segment(page_detail.get("name") or f"page_{page_id}")

        book_id = int(page_detail["book_id"])
        if self._api:
            try:
                book_name = _sanitize_segment(self._api.get_book_name(book_id))
            except Exception as e:
                self.logger.error("[BOOK] resolve failed id=%s err=%s", book_id, e)
                book_name = f"book_{book_id}"
        else:
            book_name = f"book_{book_id}"

        chapter_id = page_detail.get("chapter_id")
        if chapter_id:
            if self._api:
                try:
                    chapter_name = _sanitize_segment(
                        self._api.get_chapter_name(int(chapter_id))
                    )
                except Exception as e:
                    self.logger.error(
                        "[CHAPTER] resolve failed id=%s err=%s",
                        chapter_id,
                        e,
                    )
                    chapter_name = f"chapter_{chapter_id}"
            else:
                chapter_name = f"chapter_{chapter_id}"
        else:
            chapter_name = "_no_chapter"

        out = self.books_root / book_name / chapter_name / f"{page_name}.md"

        meta = LocalPageMeta(
            page_id=page_id,
            book=str(book_id),
            chapter=str(chapter_id) if chapter_id else "",
            title=page_detail.get("name") or page_name,
            updated_at=_now_iso(),
            remote_updated_at=page_detail.get("updated_at"),
            synced_at=_now_iso(),
            content_hash=compute_hash(page_detail.get("markdown") or ""),
        )

        lp = LocalPage(
            path=str(out),
            meta=meta,
            content=page_detail.get("markdown") or "",
        )
        self.save_page(lp)
        return lp

    # =====================================================
    # Attachment（★ID.png 固定）
    # =====================================================
    def save_attachment(
        self,
        page_id: int,
        filename: Any,
        data: Optional[bytes] = None,
        **_,
    ) -> Path:
        """
        filename:
          - 数値文字列 → attachment_id
          - それ以外 → page_id を attachment_id とする
        保存名:
          - {attachment_id}.png
        """
        try:
            attachment_id = int(filename)
        except Exception:
            attachment_id = int(page_id)

        out = self.attachments_root / f"{attachment_id}.png"

        if data is None:
            self.logger.warning(
                "[REPO] attachment skipped (no data) page_id=%s id=%s",
                page_id,
                attachment_id,
            )
            return out

        if out.exists():
            self.logger.info(
                "[REPO] attachment exists -> skip id=%s",
                attachment_id,
            )
            return out

        out.write_bytes(data)
        self.logger.info(
            "[REPO] attachment saved id=%s file=%s size=%s",
            attachment_id,
            out.name,
            len(data),
        )
        return out

    # =====================================================
    # Gallery（変更なし）
    # =====================================================
    def save_gallery_image(
        self,
        filename: Any,
        data: Optional[bytes] = None,
        **_,
    ) -> Path:
        fn = _sanitize_segment(filename)
        out = self.gallery_root / fn

        if data is None:
            self.logger.warning(
                "[REPO] gallery skipped (no data) file=%s",
                out,
            )
            return out

        if out.exists():
            self.logger.info(
                "[REPO] gallery exists -> skip file=%s",
                out,
            )
            return out

        out.write_bytes(data)
        self.logger.info(
            "[REPO] gallery saved file=%s size=%s",
            out,
            len(data),
        )
        return out
