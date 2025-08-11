# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['[v2] Certezza - Robos.py'],
    pathex=[],
    binaries=[],
    datas=[('img', 'img'), ('scripts', 'scripts'), ('documentacao\\v1.0.1\\Documentacao_Interface_Robos_Certezza_v1.0.1.pdf', '.')],
    hiddenimports=['pdfplumber', 'comtypes.client', 'screeninfo', 'screeninfo.screeninfo', 'ttkbootstrap'],
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
    [],
    exclude_binaries=True,
    name='Certezza - Robos',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=['img\\logoICO.ico'],
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='Certezza - Robos',
)
