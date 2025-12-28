# localwiki/models.py
from __future__ import annotations

from dataclasses import dataclass
from typing import Optional


# =========================================================
# App Config
# =========================================================
@dataclass
class AppConfig:
    bookstack: "BsConfig"
    local: "LocalConfig"
    sync: "SyncConfig"


# =========================================================
# BookStack Config
# =========================================================
@dataclass
class BsConfig:
    base_url: str
    token_id: str
    token_secret: str
    timeout_sec: int = 30
    verify_ssl: bool = True


# =========================================================
# Local Config
# =========================================================
@dataclass
class LocalConfig:
    repo_root: str
    assets_dir: Optional[str] = None

    def get_assets_dir(self) -> str:
        if self.assets_dir:
            return self.assets_dir
        return f"{self.repo_root}/assets"


# =========================================================
# Sync Config
# =========================================================
@dataclass
class SyncConfig:
    per_page: int = 50
    stop_on_conflict: bool = False

    # rate-limit 対策（asset / gallery 共通）
    asset_interval_sec: float = 0.1

    create_missing_books: bool = False
    create_missing_chapters: bool = False


# =========================================================
# Sync Report（★欠落していた定義を追加）
# =========================================================
@dataclass
class SyncReport:
    ok: bool = True
    message: str = ""

    pulled: int = 0
    pushed: int = 0
    assets: int = 0
    errors: int = 0


# =========================================================
# BookStack Entities
# =========================================================
@dataclass
class BsBook:
    id: int
    name: str
    slug: str


@dataclass
class BsChapter:
    id: int
    name: str
    slug: str
    book_id: int


@dataclass
class BsPage:
    id: int
    name: str
    slug: str
    book_id: int
    chapter_id: Optional[int]
    updated_at: str
    book_slug: Optional[str] = None


# =========================================================
# Local Page
# =========================================================
@dataclass
class LocalPageMeta:
    page_id: Optional[int]
    book: str
    chapter: str
    title: str
    updated_at: str
    remote_updated_at: Optional[str] = None
    synced_at: Optional[str] = None
    content_hash: str = ""


# backward compatibility
PageMeta = LocalPageMeta


@dataclass
class LocalPage:
    path: str
    meta: LocalPageMeta
    content: str
