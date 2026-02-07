# core/annotation.py
from __future__ import annotations

import logging
from dataclasses import dataclass
from typing import List, Optional, Tuple

from PySide6 import QtCore


# ==================================================
# Logger (no root pollution)
# ==================================================
def _get_logger() -> logging.Logger:
    logger = logging.getLogger("Core.Annotation")
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
# Data
# ==================================================
@dataclass
class AnnoRect:
    x: int
    y: int
    w: int
    h: int
    color: str
    stroke: int
    ui_color: Optional[str] = None
    
# ==================================================
# Annotation Manager
# ==================================================
class AnnotationManager(QtCore.QObject):
    changed = QtCore.Signal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self.annos: List[AnnoRect] = []
        self.selected: Optional[int] = None

    # ------------------------------
    # basic ops
    # ------------------------------
    def add(
        self,
        x: int = 16,
        y: int = 16,
        w: int = 160,
        h: int = 90,
        *,
        color: str,
        stroke: int,
    ):
        self.annos.append(AnnoRect(x, y, w, h, color, stroke))
        self.selected = len(self.annos) - 1
        logger.debug(
            "add rect idx=%d (%d,%d,%d,%d)",
            self.selected, x, y, w, h
        )
        self.changed.emit()

    def remove_selected(self):
        if self.selected is None:
            return
        if 0 <= self.selected < len(self.annos):
            logger.debug("remove selected idx=%d", self.selected)
            del self.annos[self.selected]
            self.selected = None
            self.changed.emit()

    def remove_at(self, idx: int):
        if 0 <= idx < len(self.annos):
            logger.debug("remove rect idx=%d", idx)
            del self.annos[idx]
            if self.selected == idx:
                self.selected = None
            elif self.selected is not None and self.selected > idx:
                self.selected -= 1
            self.changed.emit()

    # ------------------------------
    # geometry
    # ------------------------------
    def qrect(self, idx: int) -> QtCore.QRect:
        a = self.annos[idx]
        return QtCore.QRect(a.x, a.y, a.w, a.h)

    def move_to(self, idx: int, rect: QtCore.QRect):
        a = self.annos[idx]
        a.x = rect.x()
        a.y = rect.y()
        a.w = rect.width()
        a.h = rect.height()
        logger.debug(
            "move rect idx=%d -> (%d,%d,%d,%d)",
            idx, a.x, a.y, a.w, a.h
        )
        self.changed.emit()

    # ------------------------------
    # hit test
    # ------------------------------
    def hit_handle(
        self,
        pos: QtCore.QPoint,
        handle_rects_fn,
    ) -> Tuple[Optional[int], Optional[str]]:
        """
        handle_rects_fn: Callable[[QRect], Dict[str, QRect]]
        """
        for i in reversed(range(len(self.annos))):
            r = self.qrect(i)
            for k, hr in handle_rects_fn(r).items():
                if hr.contains(pos):
                    logger.debug("hit handle idx=%d handle=%s", i, k)
                    return i, k
        return None, None

    def hit_body(self, pos: QtCore.QPoint) -> Optional[int]:
        for i in reversed(range(len(self.annos))):
            if self.qrect(i).contains(pos):
                logger.debug("hit body idx=%d", i)
                return i
        return None

    def hit_body_expanded(
        self,
        pos: QtCore.QPoint,
        pad: int = 6,
    ) -> Optional[int]:
        for i in reversed(range(len(self.annos))):
            r = self.qrect(i).adjusted(-pad, -pad, pad, pad)
            if r.contains(pos):
                logger.debug("hit body(expanded) idx=%d", i)
                return i
        return None

    def hit_close(
        self,
        pos: QtCore.QPoint,
        *,
        close_rect_fn,
    ) -> Optional[int]:
        """
        Hit-test close button.
        close_rect_fn: Callable[[QRect], QRect]
        """
        for i in reversed(range(len(self.annos))):
            r = self.qrect(i)
            rc = close_rect_fn(r)
            if rc.contains(pos):
                return i
        return None
