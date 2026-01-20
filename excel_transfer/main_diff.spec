# -*- mode: python ; coding: utf-8 -*-

a = Analysis(
    ["main_diff.py"],
    pathex=["."],  # excel_transfer を cwd として扱う
    binaries=[],
    datas=[
        ("config", "config"),  # config 同梱
        ("licensing/THIRD_PARTY_LICENSES.txt", "licensing"),  # ★ 追加（必須）
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

pyz = PYZ(
    a.pure,
    a.zipped_data,
)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="main_diff",
    console=True,   # Diff 単体ツールなので OK
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    name="main_diff",
)
