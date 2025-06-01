import glob
from PyInstaller.utils.hooks import collect_all
import os

block_cipher = None

# 1) Todas las DLL de GTK/Cairo
gtk_bins = glob.glob(r'C:\Users\pysnepsdbs08\gtk3-runtime\bin\*.dll')
gtk_bins.append(
    r'C:\Users\pysnepsdbs08\AppData\Local\Programs\Python\Python313\DLLs\unicodedata.pyd'
)

# 2) Todo CustomTkinter (código + datos)
ctk_datas, ctk_binaries, ctk_hidden = collect_all('customtkinter')

# Asegúrate de que las rutas a los archivos como `credentials.json` y `token.json` sean correctas
a = Analysis(
    ['login_app.py'],
    pathex=[],
    binaries=[*( (path, '.') for path in gtk_bins ), *ctk_binaries],
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
        (r'C:\Users\pysnepsdbs08\AppData\Local\Programs\Python\Python313\Lib\site-packages\PyQt5\Qt5\plugins\platforms\qwindows.dll', 'platforms'),
        # Dashboard y logo
        (r'C:\Users\pysnepsdbs08\Downloads\Capturador_Actualizaciones\dashboard.py', '.'),
        (r'C:\Users\pysnepsdbs08\Downloads\Capturador_Actualizaciones\LogoImg_dark.png', '.'),
        (r'C:\Users\pysnepsdbs08\Downloads\Capturador_Actualizaciones\LogoImg_light.png', '.'),
        # JSONS: Se copian al mismo directorio donde estará el EXE
        (r'C:\Users\pysnepsdbs08\Downloads\Capturador_Actualizaciones\latest.json', '.'),
        # versionamiento
        (r'C:\Users\pysnepsdbs08\Downloads\Capturador_Actualizaciones\version.py', '.'),
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
    icon=r'C:\Users\pysnepsdbs08\Downloads\Capturador_Actualizaciones\Logo.ico'
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
