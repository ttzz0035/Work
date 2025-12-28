# sync_engine.py
from __future__ import annotations

import time
from typing import Dict, Any, Set

from wiki_module.logger import AppLogger
from wiki_module.bookstack_api import BookStackApi
from wiki_module.fs_repo import LocalWikiRepo
from wiki_module.models import SyncReport, SyncConfig


class BidirectionalSyncEngine:
    """
    BookStack -> LocalWiki 同期エンジン（PULL専用・確定版）

    【最終確定・互換維持ルール】
    - 既存の関数名・キー名は一切変更しない
    - Page一覧(/api/pages)で得られる book_id / chapter_id を「核ID」とする
    - Book名・Chapter名は **必ず API (/api/books/{id}, /api/chapters/{id}) で取得**
    - page detail の book / chapter 情報は信用しない（欠損するため）
    """

    def __init__(
        self,
        api: BookStackApi,
        repo: LocalWikiRepo,
        sync_cfg: SyncConfig,
        logger_name: str = "SyncEngine",
    ) -> None:
        self.api = api
        self.repo = repo
        self.cfg = sync_cfg
        self.logger = AppLogger(logger_name, "logs/sync_engine.log", "INFO").get()

    # =====================================================
    # PULL
    # =====================================================
    def pull(self) -> SyncReport:
        report = SyncReport(ok=True, message="ok")
        seen_page_ids: Set[int] = set()

        try:
            books = self.api.list_books()
            self.logger.info("[PULL] books=%s", len(books))

            # =================================================
            # Pages
            # =================================================
            for book in books:
                offset = 0

                while True:
                    pages = self.api.list_pages(
                        book_id=book.id,
                        per_page=self.cfg.per_page,
                        offset=offset,
                    )
                    if not pages:
                        break

                    for p in pages:
                        if p.id in seen_page_ids:
                            continue
                        seen_page_ids.add(p.id)

                        # -----------------------------------------
                        # ★ 核IDから Book / Chapter 名を API で取得
                        # -----------------------------------------
                        # Book
                        try:
                            book_slug = self.api.get_book_name(p.book_id)
                        except Exception:
                            book_slug = "default"

                        # Chapter
                        if p.chapter_id in (None, 0):
                            chapter_name = ""
                        else:
                            try:
                                chapter_name = self.api.get_chapter_name(p.chapter_id)
                            except Exception:
                                chapter_name = ""

                        page_title = p.name
                        page_slug = p.slug

                        self.logger.info(
                            "[PULL] page id=%s book=%s chapter=%s title=%s",
                            p.id,
                            book_slug,
                            chapter_name or "_no_chapter",
                            page_title,
                        )

                        # -----------------------------------------
                        # 本文取得
                        # -----------------------------------------
                        detail: Dict[str, Any] = self.api.get_page_detail(p.id)

                        # -----------------------------------------
                        # fs_repo 互換キー（変更禁止）
                        # -----------------------------------------
                        detail["_local_book_slug"] = book_slug
                        detail["_local_chapter_name"] = chapter_name
                        detail["_local_page_slug"] = page_slug
                        detail["_local_page_title"] = page_title

                        self.repo.save_page_from_remote(detail)
                        report.pulled += 1

                        time.sleep(0.05)

                    offset += len(pages)

            # =================================================
            # Attachments（ページ非依存・重複排除）
            # =================================================
            self.logger.info("[PULL] attachments (global)")

            seen_attachment_ids: Set[int] = set()

            for page_id in seen_page_ids:
                try:
                    atts = self.api.list_page_attachments(page_id)
                except Exception as e:
                    report.errors += 1
                    self.logger.error(
                        "[PULL] attachment list failed page_id=%s err=%s",
                        page_id,
                        e,
                    )
                    continue

                for a in atts:
                    att_id = a.get("id")
                    if not att_id:
                        continue
                    att_id = int(att_id)

                    if att_id in seen_attachment_ids:
                        continue
                    seen_attachment_ids.add(att_id)

                    try:
                        data = self.api.download_attachment(att_id)
                        name = a.get("name") or f"attachment_{att_id}"
                        self.repo.save_attachment(att_id, name, data)
                        report.assets += 1
                        time.sleep(self.cfg.asset_interval_sec)
                    except Exception as e:
                        report.errors += 1
                        self.logger.error(
                            "[PULL] attachment failed id=%s err=%s",
                            att_id,
                            e,
                        )

            # =================================================
            # Image Gallery（uploaded_to 単位で最新のみ）
            # =================================================
            self.logger.info("[PULL] image-gallery (global latest only)")

            latest_by_uploaded_to: Dict[int, Dict[str, Any]] = {}

            for page_id in seen_page_ids:
                try:
                    items = self.api.list_image_gallery(page_id)
                except Exception as e:
                    report.errors += 1
                    self.logger.error(
                        "[PULL] gallery list failed page_id=%s err=%s",
                        page_id,
                        e,
                    )
                    continue

                for it in items:
                    uploaded_to = it.get("uploaded_to")
                    if not uploaded_to:
                        continue
                    uploaded_to = int(uploaded_to)

                    prev = latest_by_uploaded_to.get(uploaded_to)
                    if not prev or it.get("updated_at", "") > prev.get(
                        "updated_at", ""
                    ):
                        latest_by_uploaded_to[uploaded_to] = it

            for it in latest_by_uploaded_to.values():
                gid = int(it["id"])
                filename = it.get("name") or f"gallery_{gid}.png"

                try:
                    data = self.api.download_gallery_image(gid)
                    self.repo.save_gallery_image(filename, data)
                    report.assets += 1
                    time.sleep(self.cfg.asset_interval_sec)
                except Exception as e:
                    report.errors += 1
                    self.logger.error(
                        "[PULL] gallery failed id=%s err=%s",
                        gid,
                        e,
                    )

        except Exception as e:
            report.ok = False
            report.message = str(e)
            self.logger.error("[PULL] fatal err=%s", e)

        return report

    # =====================================================
    # PUSH（互換用 stub）
    # =====================================================
    def push(self) -> SyncReport:
        self.logger.info("[PUSH] not implemented")
        return SyncReport(ok=False, message="push not implemented")
