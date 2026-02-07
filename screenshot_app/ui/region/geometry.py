# ui/region/geometry.py
from __future__ import annotations

import logging
from typing import Dict

from PySide6 import QtCore


# ==================================================
# Logger (no root pollution)
# ==================================================
def _get_logger() -> logging.Logger:
    logger = logging.getLogger("UI.Region.Geometry")
    if not logger.handlers:
        logger.setLevel(logging.DEBUG)
        h = logging.StreamHandler()
        h.setLevel(logging.DEBUG)
        fmt = logging.Formatter(
            "%(asctime)s [%(levelname)s] %(name)s: %(message)s"
        )
        h.setFormatter(fmt)
        logger.addHandler(h)
        logger.propagate = False
    return logger


logger = _get_logger()


# ==================================================
# Geometry helpers (PURE FUNCTIONS)
# ==================================================
def handle_rects(rect: QtCore.QRect, handle_size: int = 10) -> Dict[str, QtCore.QRect]:
    """
    Resize handle rectangles for a given rect.

    IMPORTANT:
    - rect is treated as READ-ONLY
    - returned rects are NEW instances
    - safe for paint / hit-test
    """
    hs = int(handle_size)
    cx = rect.x() + rect.width() // 2
    cy = rect.y() + rect.height() // 2

    logger.debug(
        "handle_rects input rect=(%d,%d,%d,%d) hs=%d",
        rect.x(), rect.y(), rect.width(), rect.height(), hs
    )

    return {
        "nw": QtCore.QRect(rect.x() - hs // 2, rect.y() - hs // 2, hs, hs),
        "n":  QtCore.QRect(cx - hs // 2, rect.y() - hs // 2, hs, hs),
        "ne": QtCore.QRect(rect.right() - hs // 2, rect.y() - hs // 2, hs, hs),
        "e":  QtCore.QRect(rect.right() - hs // 2, cy - hs // 2, hs, hs),
        "se": QtCore.QRect(rect.right() - hs // 2, rect.bottom() - hs // 2, hs, hs),
        "s":  QtCore.QRect(cx - hs // 2, rect.bottom() - hs // 2, hs, hs),
        "sw": QtCore.QRect(rect.x() - hs // 2, rect.bottom() - hs // 2, hs, hs),
        "w":  QtCore.QRect(rect.x() - hs // 2, cy - hs // 2, hs, hs),
    }


def clamp_inside(
    rect: QtCore.QRect,
    bounds: QtCore.QRect,
    min_w: int = 12,
    min_h: int = 12,
) -> QtCore.QRect:
    """
    Clamp rect inside bounds with minimum size.

    IMPORTANT:
    - rect is NOT modified
    - a NEW QRect is returned
    - safe for paint / drag / hit-test
    - positional args ARE ALLOWED (for backward compatibility)
    """
    src = QtCore.QRect(rect)  # input snapshot
    r = QtCore.QRect(rect)    # working copy

    if r.width() < min_w:
        r.setWidth(min_w)
    if r.height() < min_h:
        r.setHeight(min_h)

    if r.left() < bounds.left():
        r.moveLeft(bounds.left())
    if r.top() < bounds.top():
        r.moveTop(bounds.top())
    if r.right() > bounds.right():
        r.moveRight(bounds.right())
    if r.bottom() > bounds.bottom():
        r.moveBottom(bounds.bottom())

    logger.debug(
        "clamp_inside in=(%d,%d,%d,%d) out=(%d,%d,%d,%d) bounds=(%d,%d,%d,%d)",
        src.x(), src.y(), src.width(), src.height(),
        r.x(), r.y(), r.width(), r.height(),
        bounds.x(), bounds.y(), bounds.width(), bounds.height(),
    )

    return r


def rect_close_rect(rect: QtCore.QRect, close_size: int = 16) -> QtCore.QRect:
    """
    Close button rectangle (top-right inside rect)

    IMPORTANT:
    - rect is READ-ONLY
    - returned QRect is NEW
    """
    cs = int(close_size)

    rc = QtCore.QRect(
        rect.right() - cs + 1,
        rect.top(),
        cs,
        cs,
    )

    logger.debug(
        "rect_close_rect rect=(%d,%d,%d,%d) close=(%d,%d,%d,%d)",
        rect.x(), rect.y(), rect.width(), rect.height(),
        rc.x(), rc.y(), rc.width(), rc.height(),
    )

    return rc
