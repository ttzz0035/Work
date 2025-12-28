# =========================================================
# filesystem_backend.py
# FileSystem / OneDrive / NAS 用 Wiki Backend
# PULL / PUSH 両対応（完全コード）
# =========================================================

from __future__ import annotations

import shutil
from pathlib import Path
from typing import Optional

from wiki_module.logger import AppLogger


# =========================================================
# Logger
# =========================================================

logger = AppLogger("FileSystemBackend")


# =========================================================
# Backend
# =========================================================

class FileSystemBackend:
    """
    ファイルシステム同期 Backend

    - フォルダ構成をそのまま同期
    - PULL / PUSH 両対応
    - mtime による差分コピー
    - 削除は行わない（安全）
    """

    def __init__(
        self,
        server_root: Path,
        local_root: Path,
    ):
        self.server_root = Path(server_root).resolve()
        self.local_root = Path(local_root).resolve()

        logger.info(
            "[INIT] server_root=%s local_root=%s",
            self.server_root,
            self.local_root,
        )

    # -----------------------------------------------------
    # Public API
    # -----------------------------------------------------

    def pull(self) -> None:
        """
        サーバー（OneDrive / NAS） → ローカル
        """
        logger.info("[PULL][START]")
        self._mirror(src=self.server_root, dst=self.local_root)
        logger.info("[PULL][END]")

    def push(self) -> None:
        """
        ローカル → サーバー（OneDrive / NAS）
        """
        logger.info("[PUSH][START]")
        self._mirror(src=self.local_root, dst=self.server_root)
        logger.info("[PUSH][END]")

    # -----------------------------------------------------
    # Core
    # -----------------------------------------------------

    def _mirror(self, src: Path, dst: Path) -> None:
        """
        src → dst にミラーコピー
        - フォルダ構成保持
        - 新しいファイルのみコピー
        - 削除はしない
        """

        if not src.exists():
            logger.error("[MIRROR][FAIL] src not found: %s", src)
            return

        dst.mkdir(parents=True, exist_ok=True)

        for src_path in src.rglob("*"):
            rel = src_path.relative_to(src)
            dst_path = dst / rel

            try:
                if src_path.is_dir():
                    dst_path.mkdir(parents=True, exist_ok=True)
                    continue

                self._copy_if_needed(src_path, dst_path)

            except Exception as e:
                logger.error(
                    "[MIRROR][ERROR] path=%s err=%s",
                    rel,
                    e,
                )

    # -----------------------------------------------------
    # Utils
    # -----------------------------------------------------

    def _copy_if_needed(self, src: Path, dst: Path) -> None:
        """
        mtime 比較して必要な場合のみコピー
        """

        if not dst.exists():
            self._copy(src, dst, reason="new")
            return

        src_mtime = src.stat().st_mtime
        dst_mtime = dst.stat().st_mtime

        if src_mtime > dst_mtime:
            self._copy(src, dst, reason="update")
        else:
            logger.debug(
                "[SKIP] %s (dst newer or same)",
                src.relative_to(src.parents[0]),
            )

    def _copy(self, src: Path, dst: Path, reason: str) -> None:
        dst.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(src, dst)

        logger.info(
            "[COPY][%s] %s",
            reason,
            src.relative_to(src.parents[0]),
        )
