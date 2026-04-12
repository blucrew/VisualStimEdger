# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_all, collect_data_files

datas = [
    ('models/yolo-fastest.cfg', 'models'),
    ('models/best.weights',     'models'),
    ('icon.ico',                '.'),
    ('splash.png',              '.'),
    ('overlay.html',            '.'),
]
binaries = []
hiddenimports = [
    'pycaw.pycaw', 'comtypes.stream', 'win32timezone',
    'sounddevice', '_sounddevice_data',
    'websocket', 'websocket._abnf', 'websocket._core',
    'websocket._exceptions', 'websocket._http', 'websocket._logging',
    'websocket._socket', 'websocket._ssl_compat', 'websocket._utils',
    'ssl',
]

for pkg in ('pycaw', 'sounddevice', 'customtkinter', 'miniaudio'):
    tmp = collect_all(pkg)
    datas    += tmp[0]
    binaries += tmp[1]
    hiddenimports += tmp[2]

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=binaries,
    datas=datas,
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
    name='VisualStimEdger',
    icon='icon.ico',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=['opencv_world*.dll', 'cv2*.pyd'],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
