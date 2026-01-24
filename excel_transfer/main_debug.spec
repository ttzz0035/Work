# -*- mode: python ; coding: utf-8 -*-

a = Analysis(
    ["main_debug.py"],
    pathex=["."],          # excel_transfer を cwd として扱う
    binaries=[],
    datas=[
        ("data/config", "data/config"),  # 既存 YAML / label.yml
    ],
    hiddenimports=[
        "ui.app",
        "utils.log",
        "utils.configs",
        "services.transfer",
        "services.diff",
        "models.dto",
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)

pyz = PYZ(a.pure, a.zipped_data)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="main",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name="main",
)
