# -*- mode: python ; coding: utf-8 -*-
# Compilar con: pyinstaller build.spec

block_cipher = None

EXCLUDES = [
    'matplotlib', 'scipy', 'IPython', 'notebook', 'jupyter',
    'pytest', 'test', 'tests', 'PyQt5', 'PySide2', 'PySide6',
    'tkinter.test', 'lib2to3',
]

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=[
        'openpyxl.cell._writer',
        'pandas._libs.tslibs.base',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=EXCLUDES,
    noarchive=False,
    optimize=1,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='ConciliadorDIAN',
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
    onefile=True,
)
