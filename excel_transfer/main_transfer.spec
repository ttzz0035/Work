# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ["main_transfer.py"],
    pathex=["."],          # ★ excel_transfer を cwd として扱う
    binaries=[],
    datas=[
        ("config", "config"),  # config 同梱
    ],
    hiddenimports=[
        "ui.app",
        "services.transfer",
        "models.dto",
        "utils.log",
        "utils.configs",
    ],
    noarchive=False,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='main_transfer',
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
    name='main_transfer',
)
