# -*- mode: python ; coding: utf-8 -*-

a = Analysis(
    ["main_diff.py"],
    pathex=["."],          # ★ excel_transfer を cwd として扱う
    binaries=[],
    datas=[
        ("config", "config"),  # config 同梱
    ],
    hiddenimports=[
        "ui.app",
        "services.diff",
        "models.dto",
        "utils.log",
        "utils.configs",
    ],
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="diff_main",
    console=True,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    name="main_diff",
)
