import os
from pathlib import Path
from PyInstaller.utils.hooks import collect_all
from PyQt5 import QtCore

block_cipher = None

# Ruta base del proyecto
BASE_DIR = Path(__file__).resolve().parent

# Qt platform plugin (qwindows.dll)
qwindows_dll = QtCore.QLibraryInfo.location(QtCore.QLibraryInfo.PluginsPath)
qwindows_dll = os.path.join(qwindows_dll, 'platforms', 'qwindows.dll')

# 2) Todo CustomTkinter (código + datos)
ctk_datas, ctk_binaries, ctk_hidden = collect_all('customtkinter')

# Asegúrate de que las rutas a los archivos como `credentials.json` y `token.json` sean correctas
a = Analysis(
    ['login_app.py'],
    pathex=[str(BASE_DIR)],
    binaries=[(qwindows_dll, 'platforms'), *ctk_binaries],
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
        'request',
        'io',
        'sip',
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
    [],
    exclude_binaries=True,
    name='login_app',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
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
