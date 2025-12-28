# =========================================================
# fs_repo.py（V1FIX・Attachment=ID.png 固定・完全互換 + DIFF）
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

    ★ DIFF 追加仕様（互換破壊なし）
      - page 保存時に created / updated / unchanged を判定
      - ログに必ず差分を出力
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
    # Page IO
    # =====================================================
    def load_page(self, path: str) -> LocalPage:
        p = Path(path)
        txt = p.read_text(encoding="utf-8", errors="ignore")

        # 互換: 既存は front-matter を解析せず content に全体を入れる
        # UI 側は strip_metadata_blocks_for_preview() で除去する設計
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

    # =====================================================
    # Local New Page（UI の Save(Local) が呼ぶ）
    # =====================================================
    def make_new_local_page(
        self,
        book: str,
        chapter: str,
        title: str,
        content: str,
    ) -> LocalPage:
        """
        New Page 作成（Local only）
        - page_id は None
        - UI の on_save_local() 互換のため必須
        """
        book_s = _sanitize_segment(book or "default")
        chapter_s = _sanitize_segment(chapter) if (chapter or "").strip() else "_no_chapter"
        title_s = _sanitize_segment(title or "Untitled")

        out = self.books_root / book_s / chapter_s / f"{title_s}.md"

        meta = LocalPageMeta(
            page_id=None,
            book=book_s,
            chapter=chapter_s if chapter_s != "_no_chapter" else "",
            title=title or title_s,
            updated_at=_now_iso(),
            remote_updated_at=None,
            synced_at=None,
            content_hash=compute_hash(content or ""),
        )

        lp = LocalPage(
            path=str(out),
            meta=meta,
            content=content or "",
        )

        self.logger.info(
            "[REPO][NEW][LOCAL] book=%s chapter=%s title=%s path=%s",
            book_s,
            chapter_s,
            title,
            out,
        )

        self.save_page(lp)
        return lp

    # =====================================================
    # Local Delete（UI の Delete(Local) が呼ぶ）
    # =====================================================
    def delete_page_local(self, path: str) -> bool:
        p = Path(path)
        self.logger.info("[DIFF][DELETE][LOCAL] start path=%s", p)

        if not p.exists():
            self.logger.warning("[DIFF][DELETE][LOCAL] not found path=%s", p)
            return False

        ok = True
        try:
            p.unlink()
            self.logger.info("[DIFF][DELETE][LOCAL] removed path=%s", p)
        except Exception as e:
            ok = False
            self.logger.error("[DIFF][DELETE][LOCAL] failed path=%s err=%s", p, e)

        # 空ディレクトリ整理（chapter / book）※安全側
        try:
            parent = p.parent
            if parent.exists() and not any(parent.iterdir()):
                parent.rmdir()
                self.logger.info("[DIFF][DELETE][LOCAL] removed empty dir=%s", parent)

            book_dir = parent.parent
            if book_dir.exists() and book_dir.is_dir() and not any(book_dir.iterdir()):
                book_dir.rmdir()
                self.logger.info("[DIFF][DELETE][LOCAL] removed empty dir=%s", book_dir)
        except Exception as e:
            ok = False
            self.logger.error("[DIFF][DELETE][LOCAL] cleanup failed path=%s err=%s", p, e)

        self.logger.info("[DIFF][DELETE][LOCAL] end path=%s ok=%s", p, ok)
        return ok

    # =====================================================
    # Remote Page Save + DIFF
    # =====================================================
    def save_page_from_remote(self, page_detail: Dict[str, Any]) -> Dict[str, Any]:
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

        new_content = page_detail.get("markdown") or ""
        new_hash = compute_hash(new_content)

        # -----------------------------
        # DIFF 判定
        # -----------------------------
        if not out.exists():
            action = "created"
            old_hash = None
        else:
            try:
                old_text = out.read_text(encoding="utf-8", errors="ignore")
                old_hash = compute_hash(old_text)
            except Exception:
                old_hash = None

            if old_hash == new_hash:
                action = "unchanged"
            else:
                action = "updated"

        self.logger.info(
            "[DIFF][PAGE] %s path=%s old=%s new=%s",
            action,
            out,
            old_hash,
            new_hash,
        )

        # -----------------------------
        # 保存（unchanged でも互換のため書く）
        # -----------------------------
        meta = LocalPageMeta(
            page_id=page_id,
            book=str(book_id),
            chapter=str(chapter_id) if chapter_id else "",
            title=page_detail.get("name") or page_name,
            updated_at=_now_iso(),
            remote_updated_at=page_detail.get("updated_at"),
            synced_at=_now_iso(),
            content_hash=new_hash,
        )

        lp = LocalPage(
            path=str(out),
            meta=meta,
            content=new_content,
        )
        self.save_page(lp)

        return {
            "action": action,
            "path": str(out),
            "hash_before": old_hash,
            "hash_after": new_hash,
        }

    # =====================================================
    # Attachment（ID.png 固定）
    # =====================================================
    def save_attachment(
        self,
        page_id: int,
        filename: Any,
        data: Optional[bytes] = None,
        **_,
    ) -> Path:
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
                "[DIFF][ATTACHMENT] unchanged id=%s",
                attachment_id,
            )
            return out

        out.write_bytes(data)
        self.logger.info(
            "[DIFF][ATTACHMENT] created id=%s file=%s size=%s",
            attachment_id,
            out.name,
            len(data),
        )
        return out

    # =====================================================
    # Gallery
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
                "[DIFF][GALLERY] unchanged file=%s",
                out,
            )
            return out

        out.write_bytes(data)
        self.logger.info(
            "[DIFF][GALLERY] created file=%s size=%s",
            out,
            len(data),
        )
        return out

    def delete_page_by_meta(
        self,
        book: str,
        chapter: str,
        title: str,
    ) -> bool:
        """
        Book / Chapter / Title からローカルページを削除
        """
        book_s = _sanitize_segment(book or "default")
        chapter_s = _sanitize_segment(chapter) if (chapter or "").strip() else "_no_chapter"
        title_s = _sanitize_segment(title or "")

        if not title_s:
            self.logger.warning("[DELETE][LOCAL] empty title")
            return False

        p = self.books_root / book_s / chapter_s / f"{title_s}.md"

        self.logger.info(
            "[DIFF][DELETE][LOCAL] start book=%s chapter=%s title=%s path=%s",
            book_s,
            chapter_s,
            title_s,
            p,
        )

        if not p.exists():
            self.logger.warning(
                "[DIFF][DELETE][LOCAL] not found path=%s",
                p,
            )
            return False

        ok = True

        try:
            p.unlink()
            self.logger.info(
                "[DIFF][DELETE][LOCAL] removed page=%s",
                p,
            )
        except Exception as e:
            self.logger.error(
                "[DIFF][DELETE][LOCAL] unlink failed path=%s err=%s",
                p,
                e,
            )
            return False

        # ---- 空ディレクトリ整理（chapter / book）----
        try:
            chap_dir = p.parent
            if chap_dir.exists() and not any(chap_dir.iterdir()):
                chap_dir.rmdir()
                self.logger.info(
                    "[DIFF][DELETE][LOCAL] removed empty chapter=%s",
                    chap_dir,
                )

            book_dir = chap_dir.parent
            if book_dir.exists() and not any(book_dir.iterdir()):
                book_dir.rmdir()
                self.logger.info(
                    "[DIFF][DELETE][LOCAL] removed empty book=%s",
                    book_dir,
                )
        except Exception as e:
            ok = False
            self.logger.error(
                "[DIFF][DELETE][LOCAL] cleanup failed err=%s",
                e,
            )

        self.logger.info(
            "[DIFF][DELETE][LOCAL] end ok=%s",
            ok,
        )
        return ok

    def save_page_by_meta(
        self,
        book: str,
        chapter: str,
        title: str,
        content: str,
    ) -> LocalPage:
        """
        Book / Chapter / Title 基準で保存
        - 同一 title のみ上書き
        - 別 title は必ず別ファイル
        """
        book_s = _sanitize_segment(book or "default")
        chapter_s = _sanitize_segment(chapter) if (chapter or "").strip() else "_no_chapter"
        title_s = _sanitize_segment(title or "Untitled")

        out = self.books_root / book_s / chapter_s / f"{title_s}.md"

        exists = out.exists()
        action = "updated" if exists else "created"

        meta = LocalPageMeta(
            page_id=None,
            book=book_s,
            chapter=chapter_s if chapter_s != "_no_chapter" else "",
            title=title,
            updated_at=_now_iso(),
            remote_updated_at=None,
            synced_at=None,
            content_hash=compute_hash(content or ""),
        )

        lp = LocalPage(
            path=str(out),
            meta=meta,
            content=content or "",
        )

        self.save_page(lp)

        self.logger.info(
            "[DIFF][SAVE][LOCAL] %s book=%s chapter=%s title=%s path=%s",
            action,
            book_s,
            chapter_s,
            title,
            out,
        )

        return lp
