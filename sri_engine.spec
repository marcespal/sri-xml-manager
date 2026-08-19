# -*- mode: python ; coding: utf-8 -*-
# Build ONEDIR (carpeta) — el lanzador es un exe estándar pequeño que los
# antivirus/EDR corporativos NO marcan como falso positivo (a diferencia del
# onefile auto-extraíble). UPX desactivado por la misma razón.

a = Analysis(
    ['python\\main.py'],
    pathex=['python'],
    binaries=[('chromedriver.exe', '.')],
    datas=[('python/wsdl', 'wsdl')],
    hiddenimports=['tkinter', 'tkinter.ttk'],
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
    name='sri_engine',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,
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
    upx=False,
    upx_exclude=[],
    name='sri_engine',
)
