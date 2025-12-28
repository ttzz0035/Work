# localwiki/bookstack_api.py
from __future__ import annotations

import base64
import time
from typing import Dict, List, Optional
from urllib.parse import urlparse

import requests

from wiki_module.logger import AppLogger
from wiki_module.models import BsConfig, BsBook, BsChapter, BsPage


class BookStackApi:
    """
    BookStack REST API wrapper（Book / Chapter ID 正規解決対応）
    """

    def __init__(self, cfg: BsConfig, logger_name: str = "BookStackApi") -> None:
        self.cfg = cfg
        self.base = (cfg.base_url or "").rstrip("/")
        self.logger = AppLogger(logger_name, "logs/bookstack_api.log", "INFO").get()

        self._headers = {
            "Authorization": f"Token {cfg.token_id}:{cfg.token_secret}",
        }

        # ★ ID → 名前キャッシュ
        self._book_name_cache: Dict[int, str] = {}
        self._chapter_name_cache: Dict[int, str] = {}

    # =====================================================
    # HTTP helper（既存）
    # =====================================================
    def _request(self, method: str, path: str, params=None, json=None) -> dict:
        url = self.base + path

        for i in range(1, 6):
            self.logger.info("[HTTP] %s %s params=%s try=%s", method, url, params, i)
            r = requests.request(
                method,
                url,
                headers={**self._headers, "Content-Type": "application/json"},
                params=params,
                json=json,
                timeout=self.cfg.timeout_sec,
                verify=self.cfg.verify_ssl,
            )

            if r.status_code == 429:
                time.sleep(float(i))
                continue

            if r.status_code >= 400:
                raise RuntimeError(f"HTTP {r.status_code}: {r.text}")

            return r.json()

        raise RuntimeError("HTTP retry exceeded")

    # =====================================================
    # ★ Book / Chapter 名取得（今回の本題）
    # =====================================================
    def get_book_name(self, book_id: int) -> str:
        bid = int(book_id)
        if bid not in self._book_name_cache:
            r = self._request("GET", f"/api/books/{bid}")
            name = r.get("name")
            if not name:
                raise RuntimeError(f"book name missing id={bid}")
            self._book_name_cache[bid] = name
            self.logger.info("[CACHE] book id=%s name=%s", bid, name)
        return self._book_name_cache[bid]

    def get_chapter_name(self, chapter_id: Optional[int]) -> str:
        if not chapter_id or int(chapter_id) == 0:
            return ""

        cid = int(chapter_id)
        if cid not in self._chapter_name_cache:
            r = self._request("GET", f"/api/chapters/{cid}")
            name = r.get("name") or "-"
            self._chapter_name_cache[cid] = name
            self.logger.info("[CACHE] chapter id=%s name=%s", cid, name)
        return self._chapter_name_cache[cid]

    # =====================================================
    # Pages（既存）
    # =====================================================
    def list_books(self, per_page: int = 100) -> List[BsBook]:
        res = self._request("GET", "/api/books", params={"count": per_page})
        return [
            BsBook(id=b["id"], name=b["name"], slug=b["slug"])
            for b in res.get("data", [])
        ]

    def list_pages(self, book_id: int, per_page: int = 100, offset: int = 0) -> List[BsPage]:
        res = self._request(
            "GET",
            "/api/pages",
            params={"book_id": book_id, "count": per_page, "offset": offset},
        )
        out: List[BsPage] = []
        for p in res.get("data", []):
            out.append(
                BsPage(
                    id=p["id"],
                    name=p["name"],
                    slug=p["slug"],
                    book_id=p["book_id"],
                    chapter_id=p.get("chapter_id"),
                    updated_at=str(p.get("updated_at") or ""),
                    book_slug=p.get("book_slug"),
                )
            )
        return out

    def get_page_detail(self, page_id: int) -> dict:
        return self._request("GET", f"/api/pages/{int(page_id)}")

    # =====================================================
    # Attachments（既存・変更なし）
    # =====================================================
    def list_page_attachments(self, page_id: int) -> List[dict]:
        res = self._request("GET", "/api/attachments", params={"page_id": page_id})
        return res.get("data") or []

    def download_attachment(self, attachment_id: int) -> bytes:
        r = self._request("GET", f"/api/attachments/{attachment_id}")
        content = r.get("content")
        if not content:
            raise RuntimeError(f"attachment content missing id={attachment_id}")
        data = base64.b64decode(content)
        self.logger.info("[ATTACHMENT] decoded id=%s size=%s", attachment_id, len(data))
        return data

    # =====================================================
    # Image Gallery（既存・変更禁止）
    # =====================================================
    def list_image_gallery(self, page_id: int) -> List[dict]:
        res = self._request("GET", "/api/image-gallery", params={"page_id": page_id})
        return res.get("data") or []

    def get_gallery_item(self, gallery_id: int) -> dict:
        return self._request("GET", f"/api/image-gallery/{gallery_id}")

    def download_gallery_image(self, gallery_id: int) -> bytes:
        item = self.get_gallery_item(gallery_id)
        path = item.get("path")
        if not path:
            raise RuntimeError(f"gallery item has no path id={gallery_id}")
        return self.get_binary(path, auth=False)

    def get_binary(self, path_or_url: str, auth: bool) -> bytes:
        if path_or_url.startswith("http"):
            u = urlparse(path_or_url)
            url = f"{u.scheme}://{u.netloc}{u.path}"
        else:
            url = self.base + (path_or_url if path_or_url.startswith("/") else "/" + path_or_url)

        self.logger.info("[HTTP] GET(binary) %s auth=%s", url, auth)
        r = requests.get(
            url,
            headers=self._headers if auth else {},
            timeout=self.cfg.timeout_sec,
            verify=self.cfg.verify_ssl,
        )
        if r.status_code >= 400:
            raise RuntimeError(f"HTTP {r.status_code}: {r.text}")
        return r.content
