# -*- mode: python ; coding: utf-8 -*-

import certifi
import os

a = Analysis(
    ['得物对账单_sqlserver版.py'],
    pathex=[],
    binaries=[],
    datas=[(certifi.where(), 'certifi')],
    hiddenimports=['pyodbc', 'requests', 'certifi', 'tenacity', 'pandas', 'openpyxl', 'tkinter', 'tkinter.ttk', 'tkinter.messagebox', 'urllib.parse', 'datetime', 'hashlib', 'socket', 'threading', 'typing'],
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
    name='DewuBillSystem',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
