from __future__ import annotations
from pathlib import Path
from typing import Optional
import json

from PySide6 import QtCore, QtGui, QtWidgets

from ui.preview import PreviewPane
from core.region_window import RegionWindow
from ui.export_dialog import ExportDialog
from export import ExportOptions, get as get_exporter  # ← レジストリから取得

ROOT = Path(__file__).resolve().parents[1]
CONFIG_FILE = ROOT / "config.json"


def load_config() -> dict:
    if CONFIG_FILE.exists():
        try:
            return json.loads(CONFIG_FILE.read_text(encoding="utf-8"))
        except Exception:
            return {}
    return {}


def save_config(cfg: dict):
    CONFIG_FILE.write_text(json.dumps(cfg, ensure_ascii=False, indent=2), encoding="utf-8")


class ListAppWindow(QtWidgets.QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Capture List")
        self.resize(1100, 720)

        cfg = load_config()
        self.current_dir = Path(cfg.get("last_folder", str(ROOT / "out")))
        # 起動時点で存在しなければ作成
        self.current_dir.mkdir(parents=True, exist_ok=True)

        # 中央：初期はプレースホルダを新規生成してセット
        self._show_placeholder("No captures.\nUse Folder… or take a Shot.")

        # ツールバー
        tb = QtWidgets.QToolBar("Main")
        tb.setIconSize(QtCore.QSize(18, 18))
        self.addToolBar(QtCore.Qt.TopToolBarArea, tb)

        act_show_region = QtGui.QAction("Show Region", self)
        act_hide_region = QtGui.QAction("Hide Region", self)
        act_reload      = QtGui.QAction("Reload List", self)
        act_folder      = QtGui.QAction("Folder…", self)        # ダイアログで選択
        act_export      = QtGui.QAction("Export…", self)         # 形式/タイトル/保存先指定

        tb.addAction(act_show_region)
        tb.addAction(act_hide_region)
        tb.addSeparator()
        tb.addAction(act_reload)
        tb.addAction(act_folder)
        tb.addSeparator()
        tb.addAction(act_export)
        tb.addSeparator()

        # パス表示：編集可能
        tb.addSeparator()
        self.path_edit = QtWidgets.QLineEdit(str(self.current_dir))
        self.path_edit.setMinimumWidth(480)
        self.path_edit.setClearButtonEnabled(True)
        self.path_edit.setToolTip("Current capture folder (editable)")
        tb.addWidget(self.path_edit)

        # ステータスバー
        self.status = self.statusBar()

        # オーバーレイ領域（遅延生成）
        self._region: Optional[RegionWindow] = None

        # signals
        act_show_region.triggered.connect(self.show_region)
        act_hide_region.triggered.connect(self.hide_region)
        act_reload.triggered.connect(self.reload_list)
        act_folder.triggered.connect(self.pick_folder)
        act_export.triggered.connect(self._export_dialog)
        self.path_edit.editingFinished.connect(self._path_edited_commit)

        # 初回ロード
        self.reload_list()

    # --- actions
    def ensure_region(self):
        if self._region is None:
            self._region = RegionWindow(preview=self._current_preview_or_none(), save_dir=self.current_dir)
            self._region.show()
            self._region.raise_()
            self.status.showMessage("Region shown", 1500)

    def show_region(self):
        self.ensure_region()
        if self._region and not self._region.isVisible():
            self._region.show()
            self._region.raise_()

    def hide_region(self):
        if self._region and self._region.isVisible():
            self._region.hide()
            self.status.showMessage("Region hidden", 1500)

    def reload_list(self):
        """
        現在フォルダから *.json を読み直してカードを再構築。
        PNG/読めるJSONがひとつも無ければリスト非表示（プレースホルダ）。
        無い場合はフォルダを作成。
        """
        # 必ず存在させる
        self.current_dir.mkdir(parents=True, exist_ok=True)

        has_png = any(self.current_dir.glob("*.png"))
        jsons = sorted(self.current_dir.glob("*.json"))

        # JSON読めるものだけ採用
        readable_jsons = []
        for jp in jsons:
            try:
                json.loads(jp.read_text(encoding="utf-8"))
                readable_jsons.append(jp)
            except Exception:
                continue

        if not has_png and not readable_jsons:
            self._show_placeholder("No captures.\nUse Folder… or take a Shot.")
            self.status.showMessage("No captures found", 1500)
            if self._region:
                self._region.set_preview_sink(None)  # type: ignore
            return

        # プレビュー表示
        new_preview = PreviewPane()
        for jp in readable_jsons:
            new_preview.add_capture(jp)
        self.setCentralWidget(new_preview)
        if self._region:
            self._region.set_preview_sink(new_preview)
        self.status.showMessage(f"Loaded {len(readable_jsons)} items from {self.current_dir}", 2000)

    def pick_folder(self):
        """フォルダ選択（QFileDialog）。選択＝保存先/表示先に採用、保存、リロード、Regionにも反映"""
        start_dir = str(self.current_dir if self.current_dir.exists() else ROOT)
        path = QtWidgets.QFileDialog.getExistingDirectory(self, "Select Capture Folder", start_dir)
        if not path:
            return
        self._switch_folder(Path(path))
        self.path_edit.setText(str(self.current_dir))

    def _switch_folder(self, p: Path):
        self.current_dir = p
        # 切替時も即作成
        self.current_dir.mkdir(parents=True, exist_ok=True)

        cfg = load_config()
        cfg["last_folder"] = str(self.current_dir)
        save_config(cfg)

        if self._region:
            self._region.set_save_dir(self.current_dir)

        self.reload_list()

    def _path_edited_commit(self):
        """パス編集確定（Enter/フォーカス外れ）で切替。存在しなければ作成。"""
        text = self.path_edit.text().strip()
        if not text:
            self.path_edit.setText(str(self.current_dir))
            return
        new_path = Path(text)
        if new_path == self.current_dir:
            return
        self._switch_folder(new_path)
        self.status.showMessage(f"Folder set to: {self.current_dir}", 2000)

    def _export_dialog(self):
        """
        エクスポート実行（タイトル／形式／出力ファイルをユーザー指定）
        - HTML は既存テンプレートにタブ追記、Excel は既存ブックにシート追記
        - タイトル未入力時は既定タイトル、ファイル名未指定時は既定名にフォールバック
        """
        dlg = ExportDialog(self, default_dir=self.current_dir)
        if dlg.exec() != QtWidgets.QDialog.Accepted:
            return

        title = dlg.result_title() or None            # None → 既定タイトルで出力
        fmt   = dlg.result_format()                   # "excel" or "html"
        path  = dlg.result_path()

        try:
            exporter = get_exporter(fmt)              # レジストリから取得（excel/html）
            # 拡張子調整
            if path.suffix.lower() != exporter.ext:
                path = path.with_suffix(exporter.ext)

            out = exporter.export(
                self.current_dir,
                ExportOptions(title=title, filename=path)
            )

            # 成功通知
            if hasattr(self, "status"):
                self.status.showMessage(f"Exported: {out}", 3000)
            QtWidgets.QMessageBox.information(self, "Export", f"Exported to:\n{out}")

        except Exception as ex:
            # 失敗通知
            if hasattr(self, "status"):
                self.status.showMessage("Export failed", 3000)
            QtWidgets.QMessageBox.critical(
                self, "Export Error", f"{type(ex).__name__}\n{ex}"
            )

    def closeEvent(self, e: QtGui.QCloseEvent):
        if self._region:
            try:
                self._region.close()
            except Exception:
                pass
        super().closeEvent(e)

    # --- helpers
    def _current_preview_or_none(self) -> Optional[PreviewPane]:
        w = self.centralWidget()
        return w if isinstance(w, PreviewPane) else None

    def _show_placeholder(self, text: str):
        """毎回新規の QLabel を作ってセット → 破棄済み参照を回避"""
        ph = QtWidgets.QLabel(text)
        ph.setAlignment(QtCore.Qt.AlignCenter)
        ph.setStyleSheet("QLabel{color:#666;font-size:14px;}")
        self.setCentralWidget(ph)
