from __future__ import annotations
from PySide6 import QtCore, QtGui, QtWidgets


class MiniToolbar(QtWidgets.QWidget):
    # シグナル
    shotClicked = QtCore.Signal()
    rectClicked = QtCore.Signal()
    newColorPicked = QtCore.Signal(QtGui.QColor)
    newStrokePicked = QtCore.Signal(int)
    settingsClicked = QtCore.Signal()
    recToggleClicked = QtCore.Signal()
    playToggleClicked = QtCore.Signal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAttribute(QtCore.Qt.WA_StyledBackground, True)
        self.setStyleSheet("QWidget{background:rgba(255,255,255,240);border:1px solid #ddd;border-radius:8px;}")

        lay = QtWidgets.QHBoxLayout(self)
        lay.setContentsMargins(10, 6, 10, 6)
        lay.setSpacing(8)

        # 撮影
        self.btnShot = QtWidgets.QToolButton()
        self.btnShot.setText("Shot")
        self.btnShot.clicked.connect(self.shotClicked)

        # 矩形
        self.btnRect = QtWidgets.QToolButton()
        self.btnRect.setText("+Rect")
        self.btnRect.clicked.connect(self.rectClicked)

        # カラーパレット
        self.colorBtn = QtWidgets.QToolButton()
        self._col = QtGui.QColor("#FF3B30")
        self._update_color_icon()
        self.colorBtn.clicked.connect(self._on_pick_color)

        # 線幅
        self.spinStroke = QtWidgets.QSpinBox()
        self.spinStroke.setRange(1, 12)
        self.spinStroke.setValue(2)
        self.spinStroke.setFixedWidth(64)
        self.spinStroke.valueChanged.connect(self.newStrokePicked)

        # 録画トグル
        self.btnRecToggle = QtWidgets.QToolButton()
        self.btnRecToggle.setText("●")
        self.btnRecToggle.setToolTip("Record Start/Stop")
        self.btnRecToggle.clicked.connect(self.recToggleClicked)

        # 再生トグル
        self.btnPlayToggle = QtWidgets.QToolButton()
        self.btnPlayToggle.setText("▶")
        self.btnPlayToggle.setToolTip("Play/Stop")
        self.btnPlayToggle.clicked.connect(self.playToggleClicked)

        # 設定
        self.btnSettings = QtWidgets.QToolButton()
        self.btnSettings.setText("⚙")
        self.btnSettings.clicked.connect(self.settingsClicked)

        for w in (self.btnShot, self.btnRect, self.colorBtn, self.spinStroke,
                  self.btnRecToggle, self.btnPlayToggle, self.btnSettings):
            lay.addWidget(w)

    # --- トグル表示切替 ---

    def setRecording(self, active: bool):
        """録画中の状態に応じて ●⇔■ に切替"""
        if active:
            self.btnRecToggle.setText("■")
            self.btnRecToggle.setToolTip("Recording... Click to Stop")
        else:
            self.btnRecToggle.setText("●")
            self.btnRecToggle.setToolTip("Record Start")

    def setPlaying(self, active: bool):
        """再生中の状態に応じて ▶⇔■ に切替"""
        if active:
            self.btnPlayToggle.setText("■")
            self.btnPlayToggle.setToolTip("Playing... Click to Stop")
        else:
            self.btnPlayToggle.setText("▶")
            self.btnPlayToggle.setToolTip("Play Start")

    def _update_color_icon(self):
        pm = QtGui.QPixmap(18, 18)
        pm.fill(QtCore.Qt.transparent)
        p = QtGui.QPainter(pm)
        p.fillRect(1, 1, 16, 16, self._col)
        p.setPen(QtGui.QPen(QtGui.QColor("#333"), 1))
        p.drawRect(1, 1, 16, 16)
        p.end()
        self.colorBtn.setIcon(QtGui.QIcon(pm))

    def _on_pick_color(self):
        col = QtWidgets.QColorDialog.getColor(self._col, self, "Pick Color (for next Rect)")
        if col.isValid():
            self._col = col
            self._update_color_icon()
            self.newColorPicked.emit(col)
