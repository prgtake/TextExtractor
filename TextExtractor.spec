# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['textextractor.py'],
    pathex=[],
    binaries=[],
    datas=[('textextractor.py', '.')],
    hiddenimports=['pandas', 'docx', 'pptx', 'extract_msg', 'olefile', 'PIL._tkinter_finder'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['pytest', 'py'],
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
    name='TextExtractor',
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
