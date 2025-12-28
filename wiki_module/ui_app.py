from __future__ import annotations

import threading
from pathlib import Path
from typing import Dict, List, Optional

import flet as ft
import yaml

from wiki_module.bookstack_api import BookStackApi
from wiki_module.fs_repo import LocalWikiRepo
from wiki_module.logger import AppLogger
from wiki_module.models import AppConfig, BsConfig, LocalConfig, SyncConfig, LocalPage
from wiki_module.sync_engine import BidirectionalSyncEngine

# ★ Preview / Path helpers は外部モジュールのみ
from wiki_module.preview_utils import (
    render_markdown_for_flet,
    infer_book_chapter_from_path,
    repo_root,
    wiki_root,
)


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
        self.repo.app_cfg = self.cfg

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

        self.page.assets_dir = str(repo_root(self.repo))
        self.logger.info("[FLET][ASSETS] assets_dir=%s", self.page.assets_dir)
        self.logger.info(
            "[FLET][ASSETS] wiki_assets=%s",
            str((wiki_root(self.repo) / "assets").resolve()),
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
        book, chapter = infer_book_chapter_from_path(lp.path)
        title = lp.meta.title or Path(lp.path).stem

        self.logger.info(
            "[OPEN][APPLY][PATH] path=%s book=%s chapter=%s title=%s",
            lp.path,
            book,
            chapter,
            title,
        )

        self.book_field.value = book
        self.chapter_field.value = chapter
        self.title_field.value = title
        self.editor.value = lp.content or ""

        self.book_field.update()
        self.chapter_field.update()
        self.title_field.update()
        self.editor.update()

        md = render_markdown_for_flet(lp, self.repo, self.api, self.logger)

        # ★ 画像は Markdown 内でのみ描画（上部に分離表示しない）
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

        lp = self.repo.save_page_by_meta(book, chapter, title, content)
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
            return
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
        server_cfg = dict(raw.get("server", {}))
        server_cfg.pop("type", None)
        sync_cfg = dict(raw.get("sync", {}))
        sync_cfg.pop("url_rewrite", None)

        return AppConfig(
            bookstack=BsConfig(**server_cfg),
            local=LocalConfig(**raw["local"]),
            sync=SyncConfig(**sync_cfg),
        )

    def _open_delete_confirm_dialog(self) -> None:
        book = self.book_field.value or ""
        chapter = self.chapter_field.value or ""
        title = self.title_field.value or ""

        def _cancel(e):
            self.page.dialog.open = False
            self.page.update()

        def _confirm(e):
            self.page.dialog.open = False
            self.page.update()
            self._delete_local_confirmed()

        self.page.dialog = ft.AlertDialog(
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
        self.page.dialog.open = True
        self.page.update()

    def _delete_local_confirmed(self) -> None:
        book = self.book_field.value or ""
        chapter = self.chapter_field.value or ""
        title = self.title_field.value or ""

        if not self.repo.delete_page_by_meta(book, chapter, title):
            self.status_text.value = "delete failed"
            self.page.update()
            return

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
        self.app.logger.info("[CLICK][OPEN] path=%s", self.path)
        self.app.open_page(self.path)


class TabHandler:
    def __init__(self, app: LocalWikiApp, tab: str):
        self.app = app
        self.tab = tab

    def handle(self, e):
        self.app._switch_tab(self.tab)
