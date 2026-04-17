# -*- mode: python ; coding: utf-8 -*-

# Canonical PyInstaller spec (base). Keep this file in sync with
# "QR CODER.spec", which is only a branding variant for the executable name.
COMMON_DATAS = [
    ('icon.ico', '.'),
    ('locales', 'locales'),
    ('logs/*.db', 'logs'),
]

a = Analysis(
    ['qr_generator.py'],
    pathex=[],
    binaries=[],
    datas=COMMON_DATAS,
    hiddenimports=[],
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
    name='qr_generator',
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
    icon=['icon.ico'],
)
