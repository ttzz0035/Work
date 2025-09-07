# tests/conftest.py
import pytest
from types import SimpleNamespace
from pathlib import Path
import openpyxl

import os
import sys

# このファイル（conftest.py）のパス
_THIS_DIR = os.path.dirname(os.path.abspath(__file__))

# リポジトリルート（tests の 1 つ上を想定）
REPO_ROOT = os.path.abspath(os.path.join(_THIS_DIR, os.pardir))
if REPO_ROOT not in sys.path:
    sys.path.insert(0, REPO_ROOT)

# もし src レイアウト（src/excel_transfer/…）にしているなら、以下も追加
SRC_DIR = os.path.join(REPO_ROOT, "src")
if os.path.isdir(SRC_DIR) and SRC_DIR not in sys.path:
    sys.path.insert(0, SRC_DIR)


# ---- Excel可用性チェック（xlwings経由でExcel起動できなければskip）----
def _excel_available():
    try:
        import xlwings as xw
        app = xw.App(visible=False, add_book=False)
        app.kill()
        return True
    except Exception:
        return False

def pytest_collection_modifyitems(config, items):
    # Excelがなければ tests をすべて skip（logicはxlwings依存のため）
    if not _excel_available():
        skip = pytest.mark.skip(reason="Excel(xlwings) が利用できない環境のためスキップ")
        for item in items:
            item.add_marker(skip)

# ---- テスト用ユーティリティ ----
def wb_write(path: Path, sheets: dict):
    """
    openpyxl でシンプルな xlsx を作る。
    sheets = { "SheetName": [ [row1], [row2], ... ] }
    """
    wb = openpyxl.Workbook()
    # 既定のSheetを差し替え
    default = wb.active
    first = True
    for name, rows in sheets.items():
        ws = default if first else wb.create_sheet(title=name)
        ws.title = name
        for r in rows:
            ws.append(r)
        first = False
    wb.save(path)

@pytest.fixture
def ctx(tmp_path):
    base = tmp_path
    (base / "outputs").mkdir()
    (base / "data" / "config").mkdir(parents=True)
    # 最小限のContext（servicesが参照する属性だけ）
    return SimpleNamespace(
        base_dir=str(base),
        output_dir=str(base / "outputs"),
        user_paths_file=str(base / "user_paths.yaml"),
        app_settings={"app": {"default_dir": str(base)}},
        labels={},  # 不要
        save_user_path=lambda *args, **kwargs: None,
        default_dir_for=lambda _=None: str(base),
    )
