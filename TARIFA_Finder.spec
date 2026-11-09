# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_submodules

hiddenimports = ['win32com', 'win32com.client', 'pywintypes', 'pythoncom', 'openpyxl', 'openpyxl.styles', 'openpyxl.styles.differential', 'openpyxl.formatting.rule', 'openpyxl.utils', 'PIL', 'PIL.Image', 'PIL.ImageTk', 'pandas', 'pandas.core.frame', 'pandas.core.series', 'pandas.io.formats.excel', 'xlrd', 'xlrd.biffh']
hiddenimports += collect_submodules('win32com')


a = Analysis(
    ['G:\\CARPETA COMPRADORES\\COMPRADOR 1 - VICTOR MARTIN\\02.CONSOLA REG INFO SAP\\CODIGO\\tarifa_app.py'],
    pathex=[],
    binaries=[],
    datas=[('G:\\CARPETA COMPRADORES\\COMPRADOR 1 - VICTOR MARTIN\\02.CONSOLA REG INFO SAP\\CODIGO\\transformar_tarifa.py', '.'), ('G:\\CARPETA COMPRADORES\\COMPRADOR 1 - VICTOR MARTIN\\02.CONSOLA REG INFO SAP\\CODIGO\\tarifa_constants.py', '.'), ('G:\\CARPETA COMPRADORES\\COMPRADOR 1 - VICTOR MARTIN\\02.CONSOLA REG INFO SAP\\CODIGO\\tarifa_core.py', '.'), ('G:\\CARPETA COMPRADORES\\COMPRADOR 1 - VICTOR MARTIN\\02.CONSOLA REG INFO SAP\\CODIGO\\tarifa_dialogs.py', '.'), ('G:\\CARPETA COMPRADORES\\COMPRADOR 1 - VICTOR MARTIN\\02.CONSOLA REG INFO SAP\\CODIGO\\crear_registro_info.py', '.'), ('G:\\CARPETA COMPRADORES\\COMPRADOR 1 - VICTOR MARTIN\\02.CONSOLA REG INFO SAP\\CODIGO\\assets\\logo.png', 'assets'), ('G:\\CARPETA COMPRADORES\\COMPRADOR 1 - VICTOR MARTIN\\02.CONSOLA REG INFO SAP\\CODIGO\\assets\\logo.ico', 'assets'), ('G:\\CARPETA COMPRADORES\\COMPRADOR 1 - VICTOR MARTIN\\02.CONSOLA REG INFO SAP\\CODIGO\\assets\\logo_dali_articulos.png', 'assets')],
    hiddenimports=hiddenimports,
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
    name='TARIFA_Finder',
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
    icon=['G:\\CARPETA COMPRADORES\\COMPRADOR 1 - VICTOR MARTIN\\02.CONSOLA REG INFO SAP\\CODIGO\\assets\\logo.ico'],
)
