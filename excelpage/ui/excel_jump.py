# excel_transfer/ui/excel_jump.py
from __future__ import annotations

import os
from typing import Optional

import xlwings as xw


def _get_or_create_app(logger) -> xw.App:
    # 既存Excelを優先して使う（新規乱立を避ける）
    try:
        if xw.apps and len(xw.apps) > 0:
            app = xw.apps.active
            if app is not None:
                logger.debug("[JUMP] reuse active excel app")
                return app
    except Exception as ex:
        logger.warning("[JUMP] xw.apps.active failed: %s", ex)

    logger.info("[JUMP] create new excel app")
    return xw.App(visible=True, add_book=False)


def _find_open_book(app: xw.App, file_path: str, logger) -> Optional[xw.Book]:
    target = os.path.abspath(file_path)
    try:
        for b in app.books:
            try:
                if os.path.abspath(b.fullname) == target:
                    return b
            except Exception:
                continue
    except Exception as ex:
        logger.warning("[JUMP] iter books failed: %s", ex)
    return None


def jump_to_cell(
    file_path: str,
    sheet_name: str,
    row: int,
    col: int,
    logger,
) -> None:
    """
    指定ファイル/シート/セルへジャンプ（Excelを前面化）
    """
    logger.info("[JUMP] request file=%s sheet=%s r=%s c=%s", file_path, sheet_name, row, col)

    if not file_path or not os.path.exists(file_path):
        logger.error("[JUMP] file not found: %s", file_path)
        raise ValueError(f"file not found: {file_path}")

    if row is None or col is None:
        logger.error("[JUMP] invalid cell: r=%s c=%s", row, col)
        raise ValueError("invalid cell position")

    app = _get_or_create_app(logger)

    book = _find_open_book(app, file_path, logger)
    if book is None:
        logger.info("[JUMP] open book: %s", file_path)
        book = app.books.open(file_path, read_only=False)

    try:
        sht = book.sheets[sheet_name] if sheet_name else book.sheets[0]
    except Exception as ex:
        logger.error("[JUMP] sheet not found sheet=%s ex=%s", sheet_name, ex)
        raise

    try:
        # 前面化
        try:
            book.app.activate(steal_focus=True)
        except Exception:
            # steal_focus が無い/効かない環境もある
            book.app.activate()

        sht.activate()
        rng = sht.range((row, col))
        rng.select()

        logger.info("[JUMP] done")
    except Exception as ex:
        logger.error("[JUMP] failed ex=%s", ex)
        raise
