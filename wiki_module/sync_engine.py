# wiki_module/sync_engine.py
from __future__ import annotations

import shutil
import time
from pathlib import Path
from typing import Any, Dict, Optional, Set, Tuple

from wiki_module.bookstack_api import BookStackApi
from wiki_module.fs_repo import LocalWikiRepo
from wiki_module.logger import AppLogger
from wiki_module.models import SyncConfig, SyncReport


class BidirectionalSyncEngine:
    """
    BookStack -> LocalWiki 同期エンジン（PULL専用・確定版）

    【最終確定・互換維持ルール】
    - 既存の関数名・キー名は一切変更しない
    - Page一覧(/api/pages)で得られる book_id / chapter_id を「核ID」とする
    - Book名・Chapter名は **必ず API (/api/books/{id}, /api/chapters/{id}) で取得**
    - page detail の book / chapter 情報は信用しない（欠損するため）
    - ★ MDの参照置換は PULL 時のみ実施（Previewでは保存行為禁止）
    - ★ 置換の真実は YAML（repo.app_cfg）を使用する（api.base_url等は使用しない）
    - ★ server.type == "filesystem" の場合は、フォルダ構成をそのまま PULL/PUSH 可能
      - PULL: server.base_path -> local.repo_root
      - PUSH: local.repo_root -> server.base_path
      - 差分判定は mtime、新しい方のみコピー、削除はしない（安全優先）
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

        try:
            app_cfg = getattr(self.repo, "app_cfg", None)
            if not app_cfg:
                raise AttributeError("LocalWikiRepo.app_cfg not found (YAML config must be attached to repo)")

            server_type = self._get_server_type(app_cfg)

            if server_type == "filesystem":
                return self._pull_filesystem(app_cfg)

            return self._pull_bookstack(app_cfg)

        except Exception as e:
            report.ok = False
            report.message = str(e)
            self.logger.error("[PULL] fatal err=%s", e)

        return report

    # =====================================================
    # PUSH
    # =====================================================
    def push(self) -> SyncReport:
        report = SyncReport(ok=True, message="ok")

        try:
            app_cfg = getattr(self.repo, "app_cfg", None)
            if not app_cfg:
                raise AttributeError("LocalWikiRepo.app_cfg not found (YAML config must be attached to repo)")

            server_type = self._get_server_type(app_cfg)

            if server_type == "filesystem":
                return self._push_filesystem(app_cfg)

            # BookStack PUSH は互換維持（現状未実装）
            self.logger.info("[PUSH] not implemented (server.type=bookstack)")
            return SyncReport(ok=False, message="push not implemented")

        except Exception as e:
            report.ok = False
            report.message = str(e)
            self.logger.error("[PUSH] fatal err=%s", e)

        return report

    # =====================================================
    # Internal: server.type 判定
    # =====================================================
    def _get_server_type(self, app_cfg: Any) -> str:
        server_cfg = getattr(app_cfg, "server", None)
        if server_cfg is None:
            # 互換: server セクションが無い場合は bookstack 扱い
            self.logger.info("[CFG] server not found -> treat as bookstack")
            return "bookstack"

        v = str(getattr(server_cfg, "type", "") or "").strip().lower()
        if not v:
            self.logger.info("[CFG] server.type empty -> treat as bookstack")
            return "bookstack"
        return v

    # =====================================================
    # Internal: BookStack PULL（既存ロジック）
    # =====================================================
    def _pull_bookstack(self, app_cfg: Any) -> SyncReport:
        report = SyncReport(ok=True, message="ok")
        seen_page_ids: Set[int] = set()

        try:
            # -------------------------------------------------
            # YAML(=repo.app_cfg) を唯一の真実として利用
            # -------------------------------------------------
            sync_extra = getattr(app_cfg, "sync_extra", {})
            if not isinstance(sync_extra, dict):
                sync_extra = {}

            url_rewrite_cfg = sync_extra.get("url_rewrite", {})
            if not isinstance(url_rewrite_cfg, dict):
                url_rewrite_cfg = {}

            url_rewrite_enabled = bool(url_rewrite_cfg.get("enabled", False))
            server_base = str(getattr(app_cfg.bookstack, "base_url", "") or "").rstrip("/")
            local_assets = str(getattr(app_cfg.local, "assets_dir", "") or "").rstrip("/")

            self.logger.info(
                "[PULL][CFG] server.type=bookstack url_rewrite=%s server_base=%s local_assets=%s",
                url_rewrite_enabled,
                server_base,
                local_assets,
            )

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
                        try:
                            book_slug = self.api.get_book_name(p.book_id)
                        except Exception as e:
                            self.logger.error("[PULL] get_book_name failed book_id=%s err=%s", p.book_id, e)
                            book_slug = "default"

                        if p.chapter_id in (None, 0):
                            chapter_name = ""
                        else:
                            try:
                                chapter_name = self.api.get_chapter_name(p.chapter_id)
                            except Exception as e:
                                self.logger.error("[PULL] get_chapter_name failed chapter_id=%s err=%s", p.chapter_id, e)
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
                        # ★ PULL 時点で参照を YAML 基準で正規化（ここだけが保存前の変換点）
                        #   - Preview/OPEN時は一切置換・保存しない
                        # -----------------------------------------
                        try:
                            if url_rewrite_enabled and server_base and local_assets:
                                md = detail.get("markdown")
                                if isinstance(md, str) and md:
                                    md2 = self._rewrite_server_to_local(md, server_base, local_assets)
                                    detail["markdown"] = md2
                                    if md2 != md:
                                        self.logger.info("[PULL][MD][REWRITE] page_id=%s rewritten", p.id)
                                else:
                                    self.logger.info("[PULL][MD][REWRITE] page_id=%s skip (no markdown)", p.id)
                            else:
                                self.logger.info("[PULL][MD][REWRITE] disabled or invalid config")
                        except Exception as e:
                            report.errors += 1
                            self.logger.error("[PULL][MD][REWRITE] failed page_id=%s err=%s", p.id, e)

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
            asset_interval = float(getattr(self.cfg, "asset_interval_sec", 0.05) or 0.05)

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
                    try:
                        att_id = int(att_id)
                    except Exception:
                        continue

                    if att_id in seen_attachment_ids:
                        continue
                    seen_attachment_ids.add(att_id)

                    try:
                        data = self.api.download_attachment(att_id)
                        name = a.get("name") or f"attachment_{att_id}"
                        self.repo.save_attachment(att_id, name, data)
                        report.assets += 1
                        self.logger.info("[PULL][ASSET][ATTACHMENT] saved id=%s name=%s size=%s", att_id, name, len(data))
                        time.sleep(asset_interval)
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
                    try:
                        uploaded_to_i = int(uploaded_to)
                    except Exception:
                        continue

                    prev = latest_by_uploaded_to.get(uploaded_to_i)
                    if not prev or it.get("updated_at", "") > prev.get("updated_at", ""):
                        latest_by_uploaded_to[uploaded_to_i] = it

            for it in latest_by_uploaded_to.values():
                gid = int(it["id"])
                filename = it.get("name") or f"gallery_{gid}.png"

                try:
                    data = self.api.download_gallery_image(gid)
                    self.repo.save_gallery_image(filename, data)
                    report.assets += 1
                    self.logger.info("[PULL][ASSET][GALLERY] saved id=%s name=%s size=%s", gid, filename, len(data))
                    time.sleep(asset_interval)
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
    # Internal: FileSystem PULL/PUSH
    # =====================================================
    def _pull_filesystem(self, app_cfg: Any) -> SyncReport:
        report = SyncReport(ok=True, message="ok")

        try:
            server_root, local_root = self._get_filesystem_roots(app_cfg)

            self.logger.info("[PULL][CFG] server.type=filesystem server_root=%s local_root=%s", server_root, local_root)

            copied, skipped, errors = self._fs_mirror(src=server_root, dst=local_root)
            report.pulled += copied
            report.errors += errors

            self.logger.info("[PULL][FS] copied=%s skipped=%s errors=%s", copied, skipped, errors)

        except Exception as e:
            report.ok = False
            report.message = str(e)
            self.logger.error("[PULL][FS] fatal err=%s", e)

        return report

    def _push_filesystem(self, app_cfg: Any) -> SyncReport:
        report = SyncReport(ok=True, message="ok")

        try:
            server_root, local_root = self._get_filesystem_roots(app_cfg)

            self.logger.info("[PUSH][CFG] server.type=filesystem server_root=%s local_root=%s", server_root, local_root)

            copied, skipped, errors = self._fs_mirror(src=local_root, dst=server_root)
            # SyncReport に pushed が無い前提の互換維持: pulled に加算しない
            report.errors += errors

            self.logger.info("[PUSH][FS] copied=%s skipped=%s errors=%s", copied, skipped, errors)

        except Exception as e:
            report.ok = False
            report.message = str(e)
            self.logger.error("[PUSH][FS] fatal err=%s", e)

        return report

    def _get_filesystem_roots(self, app_cfg: Any) -> Tuple[Path, Path]:
        server_cfg = getattr(app_cfg, "server", None)
        if server_cfg is None:
            raise AttributeError("app_cfg.server not found (server.type=filesystem requires server section)")

        base_path = str(getattr(server_cfg, "base_path", "") or "").strip()
        if not base_path:
            raise ValueError("server.base_path is empty (server.type=filesystem requires base_path)")

        repo_root = str(getattr(app_cfg.local, "repo_root", "") or "").strip()
        if not repo_root:
            raise ValueError("local.repo_root is empty")

        server_root = Path(base_path).resolve()
        local_root = Path(repo_root).resolve()

        return server_root, local_root

    # =====================================================
    # Internal: FileSystem Mirror（mtime 差分コピー / 削除なし）
    # =====================================================
    def _fs_mirror(self, src: Path, dst: Path) -> Tuple[int, int, int]:
        """
        src -> dst にコピー（フォルダ構成完全維持）
        - mtime が新しい場合のみ上書き
        - dst にしか無いファイルは削除しない（安全）
        戻り値: (copied, skipped, errors)
        """
        copied = 0
        skipped = 0
        errors = 0

        if not src.exists():
            self.logger.error("[FS][MIRROR][FAIL] src not found: %s", src)
            return (0, 0, 1)

        dst.mkdir(parents=True, exist_ok=True)

        for src_path in src.rglob("*"):
            rel = src_path.relative_to(src)
            dst_path = dst / rel

            try:
                if src_path.is_dir():
                    dst_path.mkdir(parents=True, exist_ok=True)
                    continue

                if self._fs_copy_if_needed(src_path, dst_path):
                    copied += 1
                else:
                    skipped += 1

            except Exception as e:
                errors += 1
                self.logger.error("[FS][MIRROR][ERROR] path=%s err=%s", rel, e)

        return (copied, skipped, errors)

    def _fs_copy_if_needed(self, src: Path, dst: Path) -> bool:
        """
        True: コピーした
        False: スキップした
        """
        if not dst.exists():
            self._fs_copy(src, dst, reason="new")
            return True

        try:
            src_mtime = src.stat().st_mtime
            dst_mtime = dst.stat().st_mtime
        except Exception as e:
            self.logger.error("[FS][STAT][ERROR] src=%s dst=%s err=%s", src, dst, e)
            # stat 失敗時は安全側でコピーを試みる
            self._fs_copy(src, dst, reason="stat_fail_copy")
            return True

        if src_mtime > dst_mtime:
            self._fs_copy(src, dst, reason="update")
            return True

        self.logger.debug("[FS][SKIP] %s", src)
        return False

    def _fs_copy(self, src: Path, dst: Path, reason: str) -> None:
        dst.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(src, dst)
        self.logger.info("[FS][COPY][%s] %s -> %s", reason, src, dst)

    # =====================================================
    # Internal: MD rewrite（Import 失敗回避のため、このファイル内に実装）
    # =====================================================
    def _rewrite_server_to_local(self, markdown: str, server_base: str, local_assets: str) -> str:
        """
        BookStack の URL 参照をローカル assets 参照に寄せる最小実装。
        - 本関数は PULL の保存前変換点でのみ使用される
        - 既存実装が別モジュールにある場合でも、Import 失敗で落ちないようにここで完結させる
        """
        if not markdown:
            return markdown

        sb = (server_base or "").rstrip("/")
        la = (local_assets or "").rstrip("/")

        if not sb or not la:
            return markdown

        # 典型例:
        #   ![](/uploads/images/gallery/xxx.png)
        #   ![](http://host:port/uploads/images/gallery/xxx.png)
        #   <img src="http://host:port/uploads/images/gallery/xxx.png">
        # を local assets のパスへ（粗いが安全側の置換）
        out = markdown

        # 絶対URL
        out = out.replace(sb + "/uploads", la + "/uploads")
        out = out.replace(sb + "/attachments", la + "/attachments")
        out = out.replace(sb + "/api/attachments", la + "/attachments")
        out = out.replace(sb + "/api/image-gallery", la + "/uploads/images/gallery")

        # 先頭スラッシュ相対
        out = out.replace("](/uploads", "](" + la + "/uploads")
        out = out.replace("](/attachments", "](" + la + "/attachments")
        out = out.replace('src="/uploads', 'src="' + la + "/uploads")
        out = out.replace('src="/attachments', 'src="' + la + "/attachments")

        return out
