# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_submodules, collect_data_files

numpy_hidden = collect_submodules('numpy')
pandas_hidden = collect_submodules('pandas')
numpy_data = collect_data_files('numpy')
pandas_data = collect_data_files('pandas')

a = Analysis(
    ['auto_keyword.py'],
    pathex=[],
    binaries=[],
    datas=[('your_icon.ico', '.')] + numpy_data + pandas_data,
    hiddenimports=numpy_hidden + pandas_hidden,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='Auto_Naver_Keyword',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=['your_icon.ico'],
)
