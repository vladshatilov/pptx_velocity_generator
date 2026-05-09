# -*- mode: python ; coding: utf-8 -*-
import os
from PyInstaller.utils.hooks import collect_data_files, collect_submodules

datas = []
hiddenimports = ['openpyxl']
datas += collect_data_files('pptx')
hiddenimports += collect_submodules('pptx')

# Pillow — needed for image rotation and WebP support
hiddenimports += ['PIL', 'PIL.Image', 'PIL.WebPImagePlugin',
                  'PIL.JpegImagePlugin', 'PIL.PngImagePlugin',
                  'PIL.GifImagePlugin', 'PIL.BmpImagePlugin']

# Bundle fallback placeholder image if present next to the spec
if os.path.exists('fallback.png'):
    datas += [('fallback.png', '.')]

# Bundle user-facing readme so it always lands next to hello_kitty.exe
if os.path.exists('user_readme.txt'):
    datas += [('user_readme.txt', '.')]


a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['pandas', 'numpy', 'matplotlib', 'scipy', 'IPython', 'tkinter'],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='hello_kitty',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    version='file_version_info.txt',
    icon=['icon_file.ico'],
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='hello_kitty',
)
