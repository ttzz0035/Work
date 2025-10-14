# -*- mode: python ; coding: utf-8 -*-
from pathlib import Path

app_name = "ope_browser"
entry_script = "ope_browser.py"
try:
    base_dir = Path(__file__).resolve().parent
except NameError:
    import os
    base_dir = Path(os.getcwd())

# --- ms-playwright フォルダを自動同梱 ---
# プロジェクト内に playwright/ms-playwright を配置しておく
datas = [
    (str(base_dir / "playwright" / "ms-playwright"), "playwright/ms-playwright"),
]
# --- runtime hook 登録 ---
runtime_hooks = [str(base_dir / "playwright" / "hook_playwright_env.py")]

a = Analysis(
    [entry_script],
    pathex=[str(base_dir)],
    binaries=[],
    datas=datas,
    hiddenimports=[],
    hookspath=[],
    runtime_hooks=runtime_hooks,
    excludes=[],
)

pyz = PYZ(a.pure, a.zipped_data)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name=app_name,
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=True,  # GUIならFalse
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=False,
    name=app_name
)
