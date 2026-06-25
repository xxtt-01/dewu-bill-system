# -*- mode: python ; coding: utf-8 -*-
import certifi

a = Analysis(
    ['得物对账单_sqlserver版.py'],
    pathex=[],
    binaries=[],
    datas=[(certifi.where(), 'certifi')],
    hiddenimports=[
        'pyodbc', 'requests', 'certifi', 'tenacity', 'pandas', 'openpyxl',
        'PyQt6', 'PyQt6.QtWidgets', 'PyQt6.QtCore', 'PyQt6.QtGui',
        'urllib.parse', 'datetime', 'hashlib', 'socket', 'threading', 'typing', 'dotenv',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='得物对账单自动获取',
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
