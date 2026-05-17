# -*- mode: python ; coding: utf-8 -*-
import os

# 项目根目录
ROOT = os.path.dirname(os.path.abspath(SPEC))

# 字体文件自动检测：按优先级搜索（用户安装 -> 系统安装 -> 本地 fonts/）
_FONT_SEARCH_DIRS = [
    os.path.expandvars(r"%LOCALAPPDATA%\Microsoft\Windows\Fonts"),
    os.path.expandvars(r"%WINDIR%\Fonts"),
    os.path.join(ROOT, "fonts"),
]
_FONT_FILES = [
    "HarmonyOS_Sans_SC_Medium.ttf",
    "HarmonyOS_Sans_SC_Black.ttf",
    "HarmonyOS_Sans_SC_Regular.ttf",
]

font_datas = []
for fn in _FONT_FILES:
    for d in _FONT_SEARCH_DIRS:
        full = os.path.join(d, fn)
        if os.path.exists(full):
            font_datas.append((full, "fonts"))
            break

a = Analysis(
    [os.path.join(ROOT, 'gui_main.py')],
    pathex=[ROOT],
    binaries=[],
    datas=[
        (os.path.join(ROOT, 'config_mgr.py'),            '.'),
        (os.path.join(ROOT, 'data_loader.py'),           '.'),
        (os.path.join(ROOT, 'charts.py'),                '.'),
        (os.path.join(ROOT, 'theme_matrix_analysis.py'), '.'),
        (os.path.join(ROOT, 'config.py'),                '.'),
    ] + font_datas,
    hiddenimports=[
        'customtkinter',
        'PIL', 'PIL.Image', 'PIL.ImageFilter', 'PIL.ImageDraw',
        'pandas', 'numpy', 'openpyxl', 'xlrd',
        'matplotlib', 'matplotlib.pyplot', 'matplotlib.font_manager',
        'matplotlib.patches', 'matplotlib.colors', 'seaborn',
        'tkinter', 'tkinter.filedialog', 'tkinter.messagebox',
        'json', 'threading', 're', 'warnings',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['scipy', 'IPython', 'notebook', 'pytest'],
    noarchive=False,
    optimize=1,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='AeroEBT_v3.3',
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
    icon=os.path.join(ROOT, r'..\AeroEBT\AEROEBT.ico') if os.path.exists(
        os.path.join(ROOT, r'..\AeroEBT\AEROEBT.ico')) else None,
)
