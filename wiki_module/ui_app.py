from __future__ import annotations

import re
import threading
from pathlib import Path
from typing import Dict, List, Optional

import flet as ft
import yaml

from wiki_module.bookstack_api import BookStackApi
from wiki_module.fs_repo import LocalWikiRepo, compute_hash
from wiki_module.logger import AppLogger
from wiki_module.models import AppConfig, BsConfig, LocalConfig, SyncConfig, LocalPage
from wiki_module.sync_engine import BidirectionalSyncEngine


# =========================================================
# Preview helpers
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

# HTML <img>
_HTML_IMG_RE = re.compile(
    r"<img\s+[^>]*src\s*=\s*(['\"]?)(?P<src>[^'\"\s>]+)\1[^>]*>",
    re.IGNORECASE,
)

_HTML_ALT_RE = re.compile(
    r"alt\s*=\s*(['\"])(?P<alt>[^'\"]*)\1",
    re.IGNORECASE,
)

# Markdown image ![alt](src)
_MD_IMG_RE = re.compile(r"!\[(?P<alt>[^\]]*)\]\((?P<src>[^)]+)\)")

# BookStack URL (attachment/gallery)
_ATTACH_URL_RE = re.compile(
    r"https?://[^/]+/attachments/(?P<id>\d+)",
    re.IGNORECASE,
)

_GALLERY_URL_RE = re.compile(
    r"https?://[^/]+/image-gallery/(?P<id>\d+)",
    re.IGNORECASE,
)

# Local assets path (id only) - now supports:
#   attachments/61.PNG
#   assets/attachments/61.PNG
#   wiki/assets/attachments/61.PNG
_LOCAL_ATTACH_ID_RE = re.compile(
    r"(?:^|/)(?:wiki/)?(?:assets/)?attachments/(?P<id>\d+)(?:\.[A-Za-z0-9]+)?$",
    re.IGNORECASE,
)

_LOCAL_GALLERY_ID_RE = re.compile(
    r"(?:^|/)(?:wiki/)?(?:assets/)?gallery/(?P<id>\d+)(?:\.[A-Za-z0-9]+)?$",
    re.IGNORECASE,
)


def extract_img_alt_from_html(img_tag: str) -> str:
    m = _HTML_ALT_RE.search(img_tag)
    return (m.group("alt") or "").strip() if m else ""


def strip_metadata_blocks_for_preview(md: str, logger) -> str:
    if not md:
        return md
    before = len(md)
    md2 = _METADATA_BLOCK_RE.sub("", md)
    after = len(md2)
    if before != after:
        logger.info("[PREVIEW][META] removed metadata block (%s -> %s)", before, after)
    return md2


# =========================================================
# Repo / assets helpers（要求: wiki/assets/attachments/61.PNG で参照）
# =========================================================

def _repo_root(repo: LocalWikiRepo) -> Path:
    root = getattr(repo, "repo_root", None)
    if not root:
        raise AttributeError("LocalWikiRepo.repo_root not found")
    return Path(root).resolve()


def _wiki_root(repo: LocalWikiRepo) -> Path:
    return (_repo_root(repo) / "wiki").resolve()


def _attachment_md_path(attachment_id: int) -> str:
    return f"wiki/assets/attachments/{int(attachment_id)}.PNG"


def _gallery_md_path(gallery_id: int) -> str:
    return f"wiki/assets/gallery/{int(gallery_id)}.PNG"


def _attachment_disk_path(repo: LocalWikiRepo, attachment_id: int) -> Path:
    return (_repo_root(repo) / _attachment_md_path(attachment_id)).resolve()


def _gallery_disk_path(repo: LocalWikiRepo, gallery_id: int) -> Path:
    return (_repo_root(repo) / _gallery_md_path(gallery_id)).resolve()


def _ensure_attachment_file(
    repo: LocalWikiRepo,
    api: BookStackApi,
    attachment_id: int,
    logger,
) -> str:
    out_path = _attachment_disk_path(repo, attachment_id)
    out_path.parent.mkdir(parents=True, exist_ok=True)

    md_path = _attachment_md_path(attachment_id)

    if out_path.exists():
        logger.info("[ATTACHMENT][CACHE] id=%s path=%s", attachment_id, out_path)
        return md_path

    data = api.download_attachment(attachment_id)
    out_path.write_bytes(data)
    logger.info(
        "[ATTACHMENT][SAVE] id=%s path=%s size=%s",
        attachment_id,
        out_path,
        len(data),
    )
    return md_path


def _ensure_gallery_file(
    repo: LocalWikiRepo,
    api: BookStackApi,
    gallery_id: int,
    logger,
) -> str:
    out_path = _gallery_disk_path(repo, gallery_id)
    out_path.parent.mkdir(parents=True, exist_ok=True)

    md_path = _gallery_md_path(gallery_id)

    if out_path.exists():
        logger.info("[GALLERY][CACHE] id=%s path=%s", gallery_id, out_path)
        return md_path

    data = api.download_gallery_image(gallery_id)
    out_path.write_bytes(data)
    logger.info(
        "[GALLERY][SAVE] id=%s path=%s size=%s",
        gallery_id,
        out_path,
        len(data),
    )
    return md_path


def _normalize_slashes(s: str) -> str:
    return (s or "").replace("\\", "/").strip()


def _resolve_attachment_or_gallery_src(
    src: str,
    repo: LocalWikiRepo,
    api: BookStackApi,
    logger,
) -> Optional[str]:
    s = _normalize_slashes(src)
    logger.info("[IMG][RESOLVE] src=%s", s)

    m = _ATTACH_URL_RE.search(s)
    if m:
        attach_id = int(m.group("id"))
        md_path = _ensure_attachment_file(repo, api, attach_id, logger)
        logger.info("[IMG][ATTACHMENT][URL] %s -> %s", s, md_path)
        return md_path

    m = _GALLERY_URL_RE.search(s)
    if m:
        gallery_id = int(m.group("id"))
        md_path = _ensure_gallery_file(repo, api, gallery_id, logger)
        logger.info("[IMG][GALLERY][URL] %s -> %s", s, md_path)
        return md_path

    m = _LOCAL_ATTACH_ID_RE.search(s)
    if m:
        attach_id = int(m.group("id"))
        md_path = _ensure_attachment_file(repo, api, attach_id, logger)
        logger.info("[IMG][ATTACHMENT][LOCAL] %s -> %s", s, md_path)
        return md_path

    m = _LOCAL_GALLERY_ID_RE.search(s)
    if m:
        gallery_id = int(m.group("id"))
        md_path = _ensure_gallery_file(repo, api, gallery_id, logger)
        logger.info("[IMG][GALLERY][LOCAL] %s -> %s", s, md_path)
        return md_path

    logger.info("[IMG][SKIP] not attachment/gallery src=%s", s)
    return None


def infer_book_chapter_from_path(path: str) -> tuple[str, str]:
    """
    wiki/books/<book>/<chapter>/<title>.md
    """
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
# Image conversion（HTML / Markdown 両対応 + DEBUG）
# ※ ネスト関数禁止対応: replacer はクラス化
# =========================================================

class HtmlImgReplacer:
    def __init__(self, repo: LocalWikiRepo, api: BookStackApi, logger):
        self.repo = repo
        self.api = api
        self.logger = logger

    def __call__(self, match: re.Match) -> str:
        full = match.group(0)
        src = (match.group("src") or "").strip()
        alt = extract_img_alt_from_html(full)
        self.logger.info("[IMG][HTML] found alt=%s src=%s", alt, src)

        new_src = _resolve_attachment_or_gallery_src(src, self.repo, self.api, self.logger)
        if new_src:
            replaced = f"![{alt}]({new_src})"
            self.logger.info("[IMG][HTML] replaced -> %s", replaced)
            return replaced

        self.logger.info("[IMG][HTML] unchanged")
        return full


class MdImgReplacer:
    def __init__(self, repo: LocalWikiRepo, api: BookStackApi, logger):
        self.repo = repo
        self.api = api
        self.logger = logger

    def __call__(self, match: re.Match) -> str:
        full = match.group(0)
        alt = match.group("alt") or ""
        src = (match.group("src") or "").strip()
        self.logger.info("[IMG][MD] found alt=%s src=%s", alt, src)

        new_src = _resolve_attachment_or_gallery_src(src, self.repo, self.api, self.logger)
        if new_src:
            replaced = f"![{alt}]({new_src})"
            self.logger.info("[IMG][MD] replaced -> %s", replaced)
            return replaced

        self.logger.info("[IMG][MD] unchanged")
        return full


def convert_images_to_markdown(
    md: str,
    repo: LocalWikiRepo,
    api: BookStackApi,
    logger,
) -> str:
    if not md:
        return ""

    logger.info("[IMG][CONVERT][BEFORE]\n%s", md)

    md2 = _HTML_IMG_RE.sub(HtmlImgReplacer(repo, api, logger), md)
    md3 = _MD_IMG_RE.sub(MdImgReplacer(repo, api, logger), md2)

    logger.info("[IMG][CONVERT][AFTER]\n%s", md3)
    return md3


# =========================================================
# Markdown normalize
# =========================================================

def is_table_row_line(line: str) -> bool:
    s = (line or "").strip()
    return bool(s and "|" in s)


def is_table_sep_line(line: str) -> bool:
    s = (line or "").strip()
    if not s or "|" not in s:
        return False
    t = s.replace("|", "").replace(":", "").replace("-", "").strip()
    return t == "" and "-" in s


def normalize_tables(md: str) -> str:
    if not md:
        return ""

    lines = md.splitlines()
    out: List[str] = []
    in_table = False
    saw_header = False
    saw_sep = False

    for i, line in enumerate(lines):
        s = line.rstrip("\n")

        if not in_table:
            out.append(s)
            if is_table_row_line(s):
                j = i + 1
                while j < len(lines) and lines[j].strip() == "":
                    j += 1
                if j < len(lines) and is_table_sep_line(lines[j]):
                    in_table = True
                    saw_header = True
                    saw_sep = False
            continue

        if s.strip() == "":
            continue

        if saw_header and not saw_sep:
            if is_table_sep_line(s):
                saw_sep = True
                out.append(s)
                continue
            in_table = False
            saw_header = False
            saw_sep = False
            out.append(s)
            continue

        if is_table_row_line(s):
            out.append(s)
            continue

        in_table = False
        saw_header = False
        saw_sep = False
        out.append(s)

    return "\n".join(out)


def render_markdown_for_flet(
    lp: LocalPage,
    repo: LocalWikiRepo,
    api: BookStackApi,
    logger,
) -> str:
    md = lp.content or ""
    md = strip_metadata_blocks_for_preview(md, logger)
    md = convert_images_to_markdown(md, repo, api, logger)
    md = normalize_tables(md)
    return md


# =========================================================
# App
# =========================================================
class LocalWikiApp:
    TREE_MIN = 240
    TREE_MAX = 520
    TREE_INIT = 320
    SPLITTER_W = 6

    def __init__(self, page: ft.Page) -> None:
        self.page = page
        self.logger = AppLogger("LocalWikiApp", "logs/localwiki_ui.log", "INFO").get()

        self.cfg = self._load_config("config.yaml")
        self.repo = LocalWikiRepo(self.cfg.local.repo_root, "LocalWikiRepo")

        self.api = BookStackApi(self.cfg.bookstack, "BookStackApi")
        self.sync_engine = BidirectionalSyncEngine(self.api, self.repo, self.cfg.sync)

        self.tree_data: Dict[str, Dict[str, List[LocalPage]]] = {}
        self.selected_page: Optional[LocalPage] = None
        self.selected_path: Optional[str] = None
        self.is_busy = False
        self.current_tab = "preview"

        self.tree_scroll_y = 0
        self.preview_scroll_y = 0
        self.tree_width = self.TREE_INIT

        self._built = False

        self.status_text = ft.Text("ready")
        self.log_area = ft.TextField(
            value="",
            multiline=True,
            read_only=True,
            min_lines=4,
            max_lines=4,
        )

        self.tree_view = ft.Column(
            expand=True,
            scroll=ft.ScrollMode.AUTO,
            on_scroll=self._on_tree_scroll,
        )

        self.preview_view = ft.Column(
            expand=True,
            scroll=ft.ScrollMode.AUTO,
            on_scroll=self._on_preview_scroll,
            visible=True,
        )

        self.editor = ft.TextField(multiline=True, expand=True, visible=False)

        self.book_field = ft.TextField(label="Book", expand=2)
        self.chapter_field = ft.TextField(label="Chapter", expand=2)
        self.title_field = ft.TextField(label="Title", expand=4)

        self.new_btn = ft.ElevatedButton("New", on_click=self.on_new_page)
        self.pull_btn = ft.ElevatedButton("Pull", on_click=self.on_pull)
        self.push_btn = ft.ElevatedButton("Push", on_click=self.on_push)
        self.reload_btn = ft.OutlinedButton("Reload", on_click=self.on_reload_tree)
        self.save_btn = ft.ElevatedButton("Save (Local)", on_click=self.on_save_local)
        self.delete_local_btn = ft.ElevatedButton(
            "Delete (Local)",
            icon=ft.icons.DELETE,
            on_click=self.on_delete_local,
        )
        self._delete_confirm_dialog: Optional[ft.AlertDialog] = None

        self.tab_preview_btn = ft.TextButton(
            "Preview", on_click=TabHandler(self, "preview").handle
        )
        self.tab_edit_btn = ft.TextButton(
            "Edit", on_click=TabHandler(self, "edit").handle
        )

        self.left_container: Optional[ft.Container] = None
        self.right_container: Optional[ft.Container] = None
        self.splitter: Optional[ft.GestureDetector] = None

    def build(self) -> None:
        self.page.title = "LocalWiki (BookStack Sync)"
        self.page.theme_mode = ft.ThemeMode.DARK
        self.page.padding = 10

        self.page.assets_dir = str(_repo_root(self.repo))
        self.logger.info("[FLET][ASSETS] assets_dir=%s", self.page.assets_dir)
        self.logger.info(
            "[FLET][ASSETS] wiki_assets=%s",
            str((_wiki_root(self.repo) / "assets").resolve()),
        )

        left = ft.Column(
            [
                ft.Row(
                    [self.new_btn, self.pull_btn, self.push_btn, self.reload_btn],
                    spacing=6,
                    wrap=True,
                ),
                ft.Divider(),
                ft.Text("Pages", weight=ft.FontWeight.BOLD),
                self.tree_view,
                ft.Divider(),
                ft.Text("Status / Log", weight=ft.FontWeight.BOLD),
                self.status_text,
                self.log_area,
            ],
            expand=True,
        )

        self.left_container = ft.Container(
            left,
            width=self.tree_width,
            padding=8,
            border=ft.border.all(1, ft.colors.OUTLINE),
        )

        self.splitter = ft.GestureDetector(
            content=ft.Container(
                width=self.SPLITTER_W,
                expand=True,
                bgcolor=ft.colors.OUTLINE_VARIANT,
            ),
            on_pan_update=self._on_splitter_drag,
            mouse_cursor=ft.MouseCursor.RESIZE_LEFT_RIGHT,
        )

        right = ft.Column(
            [
                ft.Text("Editor", weight=ft.FontWeight.BOLD),
                ft.Row(
                    [self.book_field, self.chapter_field, self.title_field],
                    spacing=8,
                ),
                ft.Row([self.save_btn, self.delete_local_btn]),
                ft.Divider(),
                ft.Row([self.tab_preview_btn, self.tab_edit_btn], spacing=10),
                self.preview_view,
                self.editor,
            ],
            expand=True,
        )

        self.right_container = ft.Container(
            right,
            expand=True,
            padding=8,
            border=ft.border.all(1, ft.colors.OUTLINE),
        )

        self.page.add(
            ft.Row(
                [self.left_container, self.splitter, self.right_container],
                expand=True,
            )
        )

        self._built = True
        self._reload_tree()
        self._switch_tab("preview", update=False)
        self.page.update()

    def _apply_page_to_ui(self, lp: LocalPage) -> None:
        # ★ meta は信用しない。常に path から判定
        book, chapter = infer_book_chapter_from_path(lp.path)
        title = lp.meta.title or Path(lp.path).stem

        self.logger.info(
            "[OPEN][APPLY][PATH] path=%s book=%s chapter=%s title=%s",
            lp.path,
            book,
            chapter,
            title,
        )

        # ---- UI 反映 ----
        self.book_field.value = book
        self.chapter_field.value = chapter
        self.title_field.value = title
        self.editor.value = lp.content or ""

        # TextField は明示 update
        self.book_field.update()
        self.chapter_field.update()
        self.title_field.update()
        self.editor.update()

        # ---- Preview ----
        md = render_markdown_for_flet(lp, self.repo, self.api, self.logger)
        self.preview_view.controls = [
            ft.Markdown(
                value=md,
                selectable=True,
                extension_set=ft.MarkdownExtensionSet.GITHUB_WEB,
            )
        ]
        self.preview_view.update()

        self._switch_tab("preview", update=False)

    def _on_tree_scroll(self, e: ft.OnScrollEvent) -> None:
        self.tree_scroll_y = e.pixels

    def _on_preview_scroll(self, e: ft.OnScrollEvent) -> None:
        self.preview_scroll_y = e.pixels

    def _on_splitter_drag(self, e) -> None:
        dx = getattr(e, "delta_x", 0) or 0
        new_w = max(self.TREE_MIN, min(self.TREE_MAX, self.tree_width + int(dx)))
        if new_w != self.tree_width:
            self.tree_width = new_w
            if self.left_container:
                self.left_container.width = new_w
            self.page.update()

    def _switch_tab(self, tab: str, update: bool = True) -> None:
        if not self._built:
            return
        self.current_tab = tab
        self.preview_view.visible = tab == "preview"
        self.editor.visible = tab == "edit"
        self.tab_preview_btn.disabled = tab == "preview"
        self.tab_edit_btn.disabled = tab == "edit"
        if update:
            self.page.update()

    def on_pull(self, e) -> None:
        self.repo.set_bookstack_api(self.api)
        self._run_async(self.sync_engine.pull, "PULL")

    def on_push(self, e) -> None:
        self.repo.set_bookstack_api(self.api)
        self._run_async(self.sync_engine.push, "PUSH")

    def _run_async(self, fn, label: str) -> None:
        if self.is_busy:
            return
        self.is_busy = True
        self.status_text.value = f"{label} running..."
        self.page.update()
        threading.Thread(
            target=WorkerRunner(self, fn, label).run,
            daemon=True,
        ).start()

    def _on_worker_done(
        self,
        label: str,
        ok: bool,
        err: Optional[Exception],
    ) -> None:
        if ok:
            self.status_text.value = f"{label} done"
            self._reload_tree()
        else:
            self.status_text.value = f"{label} failed"
            if err:
                self.logger.error("[%s] %s", label, err)
        self.is_busy = False
        self.page.update()

    def on_new_page(self, e) -> None:
        self.book_field.value = "default"
        self.chapter_field.value = ""
        self.title_field.value = "New Page"
        self.editor.value = "# New Page\n\nwrite here..."
        self.preview_view.controls = []
        self._switch_tab("edit")
        self.page.update()

    def on_save_local(self, e) -> None:
        book = self.book_field.value or "default"
        chapter = self.chapter_field.value or ""
        title = self.title_field.value or "Untitled"
        content = self.editor.value or ""

        self.logger.info(
            "[UI][SAVE][LOCAL] book=%s chapter=%s title=%s",
            book,
            chapter,
            title,
        )

        lp = self.repo.save_page_by_meta(book, chapter, title, content)

        # ---- 選択状態更新（新規でも既存でも）----
        self.selected_page = lp
        self.selected_path = lp.path

        md = render_markdown_for_flet(lp, self.repo, self.api, self.logger)
        self.preview_view.controls = [
            ft.Markdown(
                value=md,
                selectable=True,
                extension_set=ft.MarkdownExtensionSet.GITHUB_WEB,
            )
        ]

        self._reload_tree()
        self._switch_tab("preview")
        self.page.update()

    def on_delete_local(self, e) -> None:
        if not (self.title_field.value or "").strip():
            self.logger.info("[UI][DELETE][LOCAL] no title")
            return

        self.logger.info("[UI][DELETE][LOCAL] open confirm dialog")
        self._open_delete_confirm_dialog()

    def on_reload_tree(self, e=None) -> None:
        self._reload_tree()

    def _reload_tree(self) -> None:
        self.tree_view.controls.clear()
        self.tree_data = self.repo.build_tree_index()

        for book, chapters in self.tree_data.items():
            self.tree_view.controls.append(
                ft.Text(f"📚 {book}", weight=ft.FontWeight.BOLD)
            )
            for ch, pages in chapters.items():
                self.tree_view.controls.append(ft.Text(f"  📁 {ch or '(no chapter)'}"))
                for lp in pages:
                    self.tree_view.controls.append(
                        ft.TextButton(
                            content=ft.Text(f"    📝 {lp.meta.title}"),
                            on_click=OpenHandler(self, lp.path).handle,
                        )
                    )

        self.tree_view.scroll_to(self.tree_scroll_y)
        self.page.update()

    def open_page(self, path: str) -> None:
        try:
            self.logger.info("[OPEN] requested path=%s", path)

            lp = self.repo.load_page(path)
            self.selected_page = lp
            self.selected_path = path

            self._apply_page_to_ui(lp)

            self.page.update()

            self.logger.info("[OPEN] done path=%s", path)
        except Exception as ex:
            self.logger.error("[OPEN][FAIL] path=%s err=%s", path, ex)
            self.status_text.value = f"open failed: {ex}"
            self.page.update()

    def _load_config(self, path: str) -> AppConfig:
        raw = yaml.safe_load(Path(path).read_text(encoding="utf-8"))
        return AppConfig(
            bookstack=BsConfig(**raw["bookstack"]),
            local=LocalConfig(**raw["local"]),
            sync=SyncConfig(**raw["sync"]),
        )

    def _open_delete_confirm_dialog(self) -> None:
        book = self.book_field.value or ""
        chapter = self.chapter_field.value or ""
        title = self.title_field.value or ""

        self.logger.info(
            "[UI][DELETE][CONFIRM] open book=%s chapter=%s title=%s",
            book,
            chapter,
            title,
        )

        def _cancel(e):
            self.logger.info("[UI][DELETE][CONFIRM] canceled")
            if self.page.dialog:
                self.page.dialog.open = False
            self.page.update()

        def _confirm(e):
            self.logger.info("[UI][DELETE][CONFIRM] confirmed")
            if self.page.dialog:
                self.page.dialog.open = False
            self.page.update()
            self._delete_local_confirmed()

        self._delete_confirm_dialog = ft.AlertDialog(
            modal=True,
            title=ft.Text("Delete page (Local)"),
            content=ft.Text(
                f"本当に削除しますか？\n\n"
                f"Book: {book}\n"
                f"Chapter: {chapter or '(no chapter)'}\n"
                f"Title: {title}"
            ),
            actions=[
                ft.TextButton("Cancel", on_click=_cancel),
                ft.ElevatedButton("Delete", on_click=_confirm),
            ],
            actions_alignment=ft.MainAxisAlignment.END,
        )

        self.page.dialog = self._delete_confirm_dialog
        self.page.dialog.open = True
        self.page.update()

    def _delete_local_confirmed(self) -> None:
        book = self.book_field.value or ""
        chapter = self.chapter_field.value or ""
        title = self.title_field.value or ""

        self.logger.info(
            "[UI][DELETE][LOCAL] confirmed book=%s chapter=%s title=%s",
            book,
            chapter,
            title,
        )

        ok = self.repo.delete_page_by_meta(book, chapter, title)

        if not ok:
            self.status_text.value = "delete failed"
            self.page.update()
            return

        # ---- UI 状態クリア ----
        self.selected_page = None
        self.selected_path = None

        self.book_field.value = ""
        self.chapter_field.value = ""
        self.title_field.value = ""
        self.editor.value = ""
        self.preview_view.controls = []

        self.book_field.update()
        self.chapter_field.update()
        self.title_field.update()
        self.editor.update()
        self.preview_view.update()

        self._reload_tree()

        self.status_text.value = "deleted (local)"
        self.page.update()


class WorkerRunner:
    def __init__(self, app: LocalWikiApp, fn, label: str):
        self.app = app
        self.fn = fn
        self.label = label

    def run(self) -> None:
        try:
            self.fn()
            self.app._on_worker_done(self.label, True, None)
        except Exception as ex:
            self.app._on_worker_done(self.label, False, ex)


class OpenHandler:
    def __init__(self, app: LocalWikiApp, path: str):
        self.app = app
        self.path = path

    def handle(self, e):
        # 「呼ばれてない」を確実に潰す
        self.app.logger.info("[CLICK][OPEN] path=%s", self.path)
        self.app.open_page(self.path)


class TabHandler:
    def __init__(self, app: LocalWikiApp, tab: str):
        self.app = app
        self.tab = tab

    def handle(self, e):
        self.app._switch_tab(self.tab)
