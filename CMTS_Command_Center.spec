# -*- mode: python ; coding: utf-8 -*-
# Build:  pyinstaller --clean --noconfirm CMTS_Command_Center.spec
# (or just run build.bat, which sets up a clean venv first)
from PyInstaller.utils.hooks import collect_data_files

# customtkinter ships theme .json files and fonts that PyInstaller does not
# pick up on its own. Without them the exe dies on launch.
datas = collect_data_files('customtkinter')

# matplotlib already has a built-in PyInstaller hook, so collect_all() is not
# needed. collect_all() pulled in every backend (Qt, wx, GTK, WebAgg) and
# fails with "multiple Qt bindings" on machines that have PyQt/PySide
# installed (Anaconda, Spyder, etc.). The app only uses the Tk backend.
hiddenimports = [
    'matplotlib.backends.backend_tkagg',
    'openpyxl.cell.rich_text',
    'openpyxl.cell.text',
]

excludes = [
    'PyQt5', 'PyQt6', 'PySide2', 'PySide6', 'qtpy',
    'wx', 'gi', 'IPython', 'jupyter', 'notebook',
    'matplotlib.tests', 'numpy.tests', 'pytest',
]

a = Analysis(
    ['CMTS_Command_Center.py'],
    pathex=[],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=excludes,
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
    name='CMTS_Command_Center',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    # UPX corrupts some Windows DLLs (vcruntime, Tcl/Tk, cryptography) and
    # makes Defender far more likely to quarantine the exe. Keep it off.
    upx=False,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
