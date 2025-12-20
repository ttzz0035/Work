# infra/excel_runtime.py
from __future__ import annotations

import os
from typing import Optional, Any


_xw = None


def _get_xw():
    """
    xlwings の import 地雷をここに閉じ込める。
    PyInstaller 環境で numpy が二重初期化されるのを避けるため、遅延 import。
    """
    global _xw
    if _xw is None:
        import xlwings as xw  # ★ xlwings import はこの1箇所だけ
        _xw = xw
    return _xw


# =====================================================
# App / Book
# =====================================================
def get_app(*, visible: bool = True, add_book: bool = False):
    xw = _get_xw()
    try:
        if xw.apps and len(xw.apps) > 0:
            app = xw.apps.active
            if app is not None:
                return app
    except Exception:
        pass
    return xw.App(visible=visible, add_book=add_book)


def find_open_book(app, file_path: str):
    target = os.path.abspath(file_path)
    try:
        for b in app.books:
            try:
                if os.path.abspath(b.fullname) == target:
                    return b
            except Exception:
                continue
    except Exception:
        return None
    return None


def open_book(app, file_path: str, *, read_only: bool = False):
    return app.books.open(file_path, read_only=read_only)


# =====================================================
# Sheet / Range
# =====================================================
def get_sheet(book, sheet_name: str):
    if sheet_name:
        return book.sheets[sheet_name]
    return book.sheets[0]


def get_used_values(sheet) -> Any:
    # used_range.value の型ゆれは呼び出し側で正規化する
    return sheet.used_range.value


def get_cell_value(sheet, row: int, col: int, *, formula: bool = False) -> Any:
    cell = sheet.range((row, col))
    return cell.formula if formula else cell.value


def activate_cell(book, sheet_name: str, row: int, col: int):
    sht = get_sheet(book, sheet_name)

    # 前面化（環境差に備える）
    try:
        try:
            book.app.activate(steal_focus=True)
        except Exception:
            book.app.activate()
    except Exception:
        pass

    sht.activate()
    sht.range((int(row), int(col))).select()


def safe_close_book(book):
    try:
        if book:
            book.close()
    except Exception:
        pass


def safe_kill_app(app):
    try:
        if app:
            app.kill()
    except Exception:
        pass
