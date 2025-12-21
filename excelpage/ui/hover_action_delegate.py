from __future__ import annotations

from typing import Optional, Tuple

from PySide6.QtCore import Qt, QModelIndex, Signal, QEvent, QRect
from PySide6.QtGui import QPainter, QFont, QColor
from PySide6.QtWidgets import QStyledItemDelegate, QStyleOptionViewItem

from models.node_tag import NodeTag
from logger import get_logger

logger = get_logger("HoverActionDelegate")


class HoverActionDelegate(QStyledItemDelegate):
    """
    TreeView の行ホバー時に、右端へ「編集」「削除」ボタンを描画し、
    クリックされたら Signal を emit する Delegate。

    - QWidget を行に埋め込まない（重くなる＆DnD/パフォーマンス悪化しやすい）
    - paint() で描画
    - editorEvent() でクリック判定
    """

    edit_requested = Signal(QModelIndex)
    delete_requested = Signal(QModelIndex)

    def __init__(self, role_tag: int, parent=None):
        super().__init__(parent)
        self._role_tag = role_tag
        self._hover_index: Optional[QModelIndex] = None

        # 表示用（軽い）
        self._btn_w = 28
        self._btn_h = 22
        self._gap = 6
        self._right_pad = 8

        logger.info("[Delegate] initialized role_tag=%s", role_tag)

    # ----------------------------
    # Hover index set/clear
    # ----------------------------
    def set_hover_index(self, idx: Optional[QModelIndex]):
        self._hover_index = idx

    def clear_hover_index(self):
        self._hover_index = None

    # ----------------------------
    # Rect compute
    # ----------------------------
    def _calc_button_rects(self, rect: QRect) -> Tuple[QRect, QRect]:
        """
        右端に [edit][delete] を並べる
        """
        x2 = rect.right() - self._right_pad
        y = rect.center().y() - (self._btn_h // 2)

        del_rect = QRect(x2 - self._btn_w, y, self._btn_w, self._btn_h)
        edit_rect = QRect(
            del_rect.left() - self._gap - self._btn_w,
            y,
            self._btn_w,
            self._btn_h,
        )
        return edit_rect, del_rect

    def _should_show_actions(self, index: QModelIndex) -> Tuple[bool, bool]:
        """
        kind によって表示を制御
          - folder(group): edit/delete OK
          - file: edit/delete OK（表示名のみ編集）
          - sheet: NG（Excel構造に関わる破壊操作を避ける）
        """
        try:
            tag = index.data(self._role_tag)
            if isinstance(tag, NodeTag):
                if tag.kind == "sheet":
                    return (False, False)
                if tag.kind in ("folder", "file"):
                    return (True, True)
        except Exception as e:
            logger.error("[Delegate] _should_show_actions failed: %s", e, exc_info=True)
        return (False, False)

    def _paint_action_button(
        self,
        painter: QPainter,
        rect: QRect,
        text: str,
        hovered: bool,
        pressed: bool = False,
    ):
        """
        シンプルな角丸ボタン風（unicode アイコンで軽量）
        """
        # 背景色
        if pressed:
            bg = QColor(60, 110, 255, 255)
            fg = QColor(255, 255, 255, 255)
        elif hovered:
            bg = QColor(70, 70, 70, 220)
            fg = QColor(230, 230, 230, 255)
        else:
            bg = QColor(40, 40, 40, 200)
            fg = QColor(210, 210, 210, 255)

        painter.save()
        painter.setRenderHint(QPainter.Antialiasing, True)
        painter.setPen(QColor(80, 80, 80, 220))
        painter.setBrush(bg)
        painter.drawRoundedRect(rect, 6, 6)

        painter.setPen(fg)
        f = QFont(painter.font())
        f.setPointSize(max(8, f.pointSize()))
        f.setBold(True)
        painter.setFont(f)
        painter.drawText(rect, Qt.AlignCenter, text)
        painter.restore()

    # ----------------------------
    # Paint
    # ----------------------------
    def paint(self, painter: QPainter, option: QStyleOptionViewItem, index: QModelIndex):
        # まず通常描画
        super().paint(painter, option, index)

        # column 0 のみ
        if index.column() != 0:
            return

        # hover していない行は何もしない
        if self._hover_index is None:
            return

        # 同じ行か判定（同一モデル前提）
        if index.row() != self._hover_index.row() or index.parent() != self._hover_index.parent():
            return

        show_edit, show_del = self._should_show_actions(index)
        if not (show_edit or show_del):
            return

        try:
            rect = option.rect
            edit_rect, del_rect = self._calc_button_rects(rect)

            # 文字表示は軽い（フォント依存で絵文字が出ない環境もあるので最小表現）
            # edit: ✏ / delete: 🗑  が出ない場合もあるので E / X 併用
            edit_text = "✏"
            del_text = "🗑"

            # 絵文字が出ない環境向け fallback 表記（コード内で条件分岐はしない方針）
            # → 見えなければ単なる四角でもOK、操作性は editorEvent で担保

            # ホバー判定（マウス座標はここでは取れないので、見た目は常に hover 扱い）
            if show_edit:
                self._paint_action_button(painter, edit_rect, edit_text, hovered=True)
            if show_del:
                self._paint_action_button(painter, del_rect, del_text, hovered=True)

        except Exception as e:
            logger.error("[Delegate] paint failed: %s", e, exc_info=True)

    # ----------------------------
    # Click handling
    # ----------------------------
    def editorEvent(self, event, model, option, index):
        try:
            if index.column() != 0:
                return super().editorEvent(event, model, option, index)

            # hover 行でしか反応させない
            if self._hover_index is None:
                return super().editorEvent(event, model, option, index)

            if index.row() != self._hover_index.row() or index.parent() != self._hover_index.parent():
                return super().editorEvent(event, model, option, index)

            show_edit, show_del = self._should_show_actions(index)
            if not (show_edit or show_del):
                return super().editorEvent(event, model, option, index)

            if event.type() not in (QEvent.MouseButtonRelease, QEvent.MouseButtonPress):
                return super().editorEvent(event, model, option, index)

            rect = option.rect
            edit_rect, del_rect = self._calc_button_rects(rect)

            pos = event.pos()
            if show_edit and edit_rect.contains(pos):
                if event.type() == QEvent.MouseButtonRelease:
                    logger.info("[Delegate] edit clicked row=%s", index.row())
                    self.edit_requested.emit(index)
                return True

            if show_del and del_rect.contains(pos):
                if event.type() == QEvent.MouseButtonRelease:
                    logger.info("[Delegate] delete clicked row=%s", index.row())
                    self.delete_requested.emit(index)
                return True

            return super().editorEvent(event, model, option, index)

        except Exception as e:
            logger.error("[Delegate] editorEvent failed: %s", e, exc_info=True)
            return super().editorEvent(event, model, option, index)
