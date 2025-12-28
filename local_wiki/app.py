from __future__ import annotations

import flet as ft

from wiki_module.ui_app import LocalWikiApp


def main(page: ft.Page) -> None:
    app = LocalWikiApp(page)
    app.build()


if __name__ == "__main__":
    ft.app(target=main)
