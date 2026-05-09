# -*- mode: python ; coding: utf-8 -*-
# macOS build spec — produces a proper .app bundle
from PyInstaller.utils.hooks import collect_all

datas = [
    ('models/yolo-fastest.cfg', 'models'),
    ('models/best.weights',     'models'),
    ('splash.png',              '.'),
    ('overlay.html',            '.'),
]
binaries = []
hiddenimports = [
    'sounddevice', '_sounddevice_data',
    'websocket', 'websocket._abnf', 'websocket._core',
    'websocket._exceptions', 'websocket._http', 'websocket._logging',
    'websocket._socket', 'websocket._ssl_compat', 'websocket._utils',
    'ssl',
    'bleak', 'bleak.backends', 'bleak.backends.corebluetooth',
    'bleak.backends.corebluetooth.client',
    'bleak.backends.corebluetooth.scanner',
]

for pkg in ('sounddevice', 'customtkinter', 'miniaudio', 'bleak'):
    tmp = collect_all(pkg)
    datas         += tmp[0]
    binaries      += tmp[1]
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
    # Exclude Windows-only packages so PyInstaller doesn't error on import
    excludes=['pycaw', 'win32timezone', 'pywin32', 'comtypes'],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    name='VisualStimEdger',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)

app = BUNDLE(
    exe,
    a.binaries,
    a.datas,
    name='VisualStimEdger.app',
    icon=None,
    bundle_identifier='com.blucrew.visualstimedger',
    info_plist={
        'NSHighResolutionCapable': True,
        'CFBundleShortVersionString': '1.7.9',
        'NSCameraUsageDescription': 'VisualStimEdger needs camera/screen access to capture the video feed.',
    },
)
