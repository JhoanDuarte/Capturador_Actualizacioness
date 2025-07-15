# -*- mode: python -*-
import glob, os, sys
from pathlib import Path
from PyInstaller.utils.hooks import collect_all, collect_data_files
from PyQt5 import QtCore

block_cipher = None
SPEC_PATH = Path(sys.argv[0]).resolve()
BASE_DIR  = SPEC_PATH.parent

# DLLs de GTK/Cairo
gtk_bins = glob.glob(str(BASE_DIR / 'gtk3-runtime' / 'bin' / '*.dll'))

# Qt platform plugin
qt_plugins = QtCore.QLibraryInfo.location(QtCore.QLibraryInfo.PluginsPath)
qwindows_dll = os.path.join(qt_plugins, 'platforms', 'qwindows.dll')

# CustomTkinter
ctk_datas, ctk_binaries, ctk_hidden = collect_all('customtkinter')

# PyQt5 data (traducciones, plugins, etc)
pyqt5_datas = collect_data_files('PyQt5')

a = Analysis(
    # Incluimos tus scripts principales
    ['login_app.py', 'dashboard.py'],
    pathex=[str(BASE_DIR)],
    binaries=[
        *((path, '.') for path in gtk_bins),
        (qwindows_dll, 'platforms'),
        *ctk_binaries
    ],
    datas=[
        ('.env', '.'),
        ('FondoLoginDark.png', '.'),
        ('FondoLoginWhite.png', '.'),
        ('doc_black.png', '.'),
        ('doc_white.png', '.'),
        ('lock_black.png', '.'),
        ('lock_white.png', '.'),
        ('moon.png', '.'),
        ('sun.png', '.'),
        ('FondoDashboardDark.png', '.'),
        ('FondoDashboardWhite.png', '.'),
        # Qt platform plugin
        (qwindows_dll, 'platforms'),
        # Archivos de código y recursos adicionales
        ('dashboard.py', '.'),
        ('LogoImg_dark.png', '.'),
        ('LogoImg_light.png', '.'),
        ('latest.json', '.'),
        ('version.py', '.'),
        *ctk_datas
    ],
    hiddenimports=[
        'unicodedata',
        'io',
        'sip',
        'PyQt5.QtCore',
        'PyQt5.QtGui',
        'PyQt5.QtWidgets',
        *ctk_hidden
    ],
    hookspath=[],
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],                 # nada aquí
    exclude_binaries=True,
    name='login_app',
    debug=True,         # ← activa debug
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=True,       # ← muestra la consola para ver errores
    icon=str(BASE_DIR / 'Logo.ico')
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    name='Dashboard-Capturacion-Datos'
)
