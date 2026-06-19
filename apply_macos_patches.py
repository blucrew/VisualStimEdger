"""
Applies all 16 macOS platform patches to main.py → main-macos.py.
Run with: python apply_macos_patches.py
"""
import pathlib, re, sys

SRC  = pathlib.Path("VSE.py")
DEST = pathlib.Path("main-macos.py")

src = SRC.read_text(encoding="utf-8")

# ─────────────────────────────────────────────────────────────────────────────
# PATCH 1 — platform flag (after `import sys`)
# ─────────────────────────────────────────────────────────────────────────────
src = src.replace(
    "import os\nimport sys\nimport tempfile\nimport ctypes",
    "import os\nimport sys\nimport tempfile\nimport ctypes\n\nWINDOWS = sys.platform == 'win32'",
    1
)
assert "WINDOWS = sys.platform == 'win32'" in src, "PATCH 1 FAILED"
print("PATCH 1 OK — WINDOWS flag")

# ─────────────────────────────────────────────────────────────────────────────
# PATCH 2 — guard DPI awareness
# ─────────────────────────────────────────────────────────────────────────────
old = """\
try:
    ctypes.windll.shcore.SetProcessDpiAwareness(2)   # PROCESS_PER_MONITOR_DPI_AWARE
except Exception:
    try:
        ctypes.windll.user32.SetProcessDPIAware()     # fallback for older Windows
    except Exception:
        pass"""
new = """\
if WINDOWS:
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(2)   # PROCESS_PER_MONITOR_DPI_AWARE
    except Exception:
        try:
            ctypes.windll.user32.SetProcessDPIAware()     # fallback for older Windows
        except Exception:
            pass"""
assert old in src, "PATCH 2 — source block not found"
src = src.replace(old, new, 1)
print("PATCH 2 OK — DPI guard")

# ─────────────────────────────────────────────────────────────────────────────
# PATCH 3 — guard comtypes cache
# ─────────────────────────────────────────────────────────────────────────────
old = """\
import comtypes.gen as _comtypes_gen
_cache = os.path.join(tempfile.gettempdir(), "VisualStimEdger_comtypes_gen")
os.makedirs(_cache, exist_ok=True)
_comtypes_gen.__path__ = [_cache]"""
new = """\
if WINDOWS:
    import comtypes.gen as _comtypes_gen
    _cache = os.path.join(tempfile.gettempdir(), "VisualStimEdger_comtypes_gen")
    os.makedirs(_cache, exist_ok=True)
    _comtypes_gen.__path__ = [_cache]"""
assert old in src, "PATCH 3 — source block not found"
src = src.replace(old, new, 1)
print("PATCH 3 OK — comtypes guard")

# ─────────────────────────────────────────────────────────────────────────────
# PATCH 4 — guard win32 and pycaw imports
# ─────────────────────────────────────────────────────────────────────────────
old = """\
import win32gui
import win32ui
import win32con
import win32api"""
new = """\
if WINDOWS:
    import win32gui
    import win32ui
    import win32con
    import win32api"""
assert old in src, "PATCH 4a — win32 imports not found"
src = src.replace(old, new, 1)

old = """\
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
from comtypes import CLSCTX_ALL"""
new = """\
if WINDOWS:
    from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
    from comtypes import CLSCTX_ALL"""
assert old in src, "PATCH 4b — pycaw import not found"
src = src.replace(old, new, 1)
print("PATCH 4 OK — win32/pycaw guards")

# ─────────────────────────────────────────────────────────────────────────────
# PATCH 5 — CONFIG_PATH platform split
# ─────────────────────────────────────────────────────────────────────────────
old = 'CONFIG_PATH = pathlib.Path(os.environ.get("APPDATA", ".")) / "VisualStimEdger" / "config.json"'
new = '''\
if WINDOWS:
    CONFIG_PATH = pathlib.Path(os.environ.get("APPDATA", ".")) / "VisualStimEdger" / "config.json"
else:
    CONFIG_PATH = pathlib.Path.home() / "Library" / "Application Support" / "VisualStimEdger" / "config.json"'''
assert old in src, "PATCH 5 — CONFIG_PATH not found"
src = src.replace(old, new, 1)
print("PATCH 5 OK — CONFIG_PATH")

# ─────────────────────────────────────────────────────────────────────────────
# PATCH 6 — guard MoveWindow in RegionSelector.__init__
# ─────────────────────────────────────────────────────────────────────────────
old = """\
        self.root.geometry(f"{w}x{h}+0+0")
        self.root.update_idletasks()
        ctypes.windll.user32.MoveWindow(
            self.root.winfo_id(), self.offset_x, self.offset_y, w, h, True)
        self.root.config(cursor="cross")"""
new = """\
        self.root.geometry(f"{w}x{h}+0+0")
        self.root.update_idletasks()
        if WINDOWS:
            try:
                ctypes.windll.user32.MoveWindow(
                    self.root.winfo_id(), self.offset_x, self.offset_y, w, h, True)
            except Exception:
                pass
        else:
            try:
                self.root.lift()
                self.root.focus_force()
            except Exception:
                pass
        self.root.config(cursor="cross")"""
assert old in src, "PATCH 6 — MoveWindow not found"
src = src.replace(old, new, 1)
print("PATCH 6 OK — MoveWindow guard")

# ─────────────────────────────────────────────────────────────────────────────
# PATCH 7 — select_region() macOS absolute-coordinate return (regex for trailing ws)
# ─────────────────────────────────────────────────────────────────────────────
p7 = re.compile(
    r'(    time\.sleep\(1\.0\)\s+)'
    r'(    x1, y1 = selector\.region\[.left.\], selector\.region\[.top.\]\s+)'
    r'(    x2, y2 = x1 \+ selector\.region\[.width.\], y1 \+ selector\.region\[.height.\]\s+)'
    r'(    cx = \(x1 \+ x2\) // 2\s+    cy = \(y1 \+ y2\) // 2\s+)'
    r'(    hwnd = win32gui\.WindowFromPoint.*?return hwnd, rel_box)',
    re.DOTALL
)
m7 = p7.search(src)
assert m7, "PATCH 7 — select_region return block not found"
replacement7 = (
    m7.group(1) +
    m7.group(2) +
    m7.group(3) +
    "\n    if not WINDOWS:\n"
    "        abs_box = {\n"
    "            'x1': x1, 'y1': y1, 'x2': x2, 'y2': y2,\n"
    "            'width': x2 - x1, 'height': y2 - y1,\n"
    "        }\n"
    "        return None, abs_box\n\n" +
    m7.group(4) +
    m7.group(5)
)
src = src[:m7.start()] + replacement7 + src[m7.end():]
print("PATCH 7 OK — select_region macOS abs coords")

# ─────────────────────────────────────────────────────────────────────────────
# PATCH 8 — capture_window_region() macOS mss-first path
# ─────────────────────────────────────────────────────────────────────────────
old = """\
def capture_window_region(hwnd, rel_box):
    try:
        left, top, right, bot = win32gui.GetWindowRect(hwnd)"""
new = """\
def capture_window_region(hwnd, rel_box):
    try:
        if not WINDOWS or hwnd is None:
            monitor = {
                "top":    int(rel_box.get('y1', rel_box.get('top', 0))),
                "left":   int(rel_box.get('x1', rel_box.get('left', 0))),
                "width":  int(rel_box['width']),
                "height": int(rel_box['height']),
            }
            grab = _get_sct().grab(monitor)
            return cv2.cvtColor(np.array(grab), cv2.COLOR_BGRA2BGR)
        left, top, right, bot = win32gui.GetWindowRect(hwnd)"""
assert old in src, "PATCH 8 — capture_window_region not found"
src = src.replace(old, new, 1)
print("PATCH 8 OK — capture_window_region macOS fallback")

# ─────────────────────────────────────────────────────────────────────────────
# PATCH 9 — select_head() macOS cv2.selectROI branch
# ─────────────────────────────────────────────────────────────────────────────
old = """\
def select_head(frame_cv, parent=None):
    if parent is None:
        root = tk.Tk()
        is_main = True
    else:
        root = tk.Toplevel(parent)
        is_main = False"""
new = """\
def select_head(frame_cv, parent=None):
    if not WINDOWS:
        display = frame_cv.copy()
        cv2.putText(display,
                    "Step 2: Draw a box around the target. ENTER/SPACE to confirm, C to cancel.",
                    (10, 25), cv2.FONT_HERSHEY_SIMPLEX, 0.55, (0, 0, 255), 2)
        win_name = "Step 2: Select Target"
        cv2.namedWindow(win_name, cv2.WINDOW_AUTOSIZE)
        roi = cv2.selectROI(win_name, display, fromCenter=False, showCrosshair=True)
        cv2.destroyWindow(win_name)
        cv2.waitKey(1)
        rx, ry, rw, rh = roi
        if rx + rw > frame_cv.shape[1] * 1.1 or ry + rh > frame_cv.shape[0] * 1.1:
            rx, ry, rw, rh = int(rx * 0.5), int(ry * 0.5), int(rw * 0.5), int(rh * 0.5)
        return (rx, ry, rw, rh)

    if parent is None:
        root = tk.Tk()
        is_main = True
    else:
        root = tk.Toplevel(parent)
        is_main = False"""
assert old in src, "PATCH 9 — select_head not found"
src = src.replace(old, new, 1)
print("PATCH 9 OK — select_head macOS OpenCV")

# ─────────────────────────────────────────────────────────────────────────────
# PATCHES 10-13 — XToysClient already HTTP; skip if already converted
# ─────────────────────────────────────────────────────────────────────────────
if "_WEBHOOK_URL = \"https://xtoys.app/webhook\"" in src:
    print("PATCH 10-13 SKIP — XToysClient already HTTP Private Webhook")
else:
    print("PATCH 10-13 WARNING — XToysClient may need manual check")

# ─────────────────────────────────────────────────────────────────────────────
# PATCH 14 — guard list_audio_devices() and WindowsAudioClient
# ─────────────────────────────────────────────────────────────────────────────
# Find and replace list_audio_devices function + WindowsAudioClient class with
# if WINDOWS: guarded versions plus macOS stubs.

old = """\
def list_audio_devices():
    \"\"\"Return all render (output) devices."""
new = """\
if WINDOWS:
 def list_audio_devices():
    \"\"\"Return all render (output) devices."""
# The above approach is fragile; use a smarter marker replacement instead.

# Wrap ONLY the audio block (list_audio_devices + WindowsAudioClient) in the
# Windows guard. The OBS OverlayServer is defined immediately after them but is
# pure asyncio/socket (cross-platform), so the guard must STOP before it —
# otherwise OverlayServer is undefined on macOS and App.__init__ NameErrors when
# it does `self._overlay = OverlayServer()`. (Was: marker_end="\nclass MusicPlayer:",
# which swept OverlayServer into the if-WINDOWS block.)
marker_start = "\ndef list_audio_devices():\n"
marker_end   = "\n# ── OBS overlay WebSocket server"

idx_start = src.find(marker_start)
idx_end   = src.find(marker_end)
assert idx_start > 0, "PATCH 14 — list_audio_devices not found"
assert idx_end   > 0, "PATCH 14 — OBS overlay marker (guard end) not found"

original_block = src[idx_start:idx_end]

# Indent every line by 4 spaces for the if WINDOWS: block
indented = "\n".join(
    ("    " + line if line.strip() else line)
    for line in original_block.splitlines()
)

macos_stubs = '''

else:
    def list_audio_devices():
        return []

    class WindowsAudioClient:
        def __init__(self, device):
            self._connected = False
        @property
        def connected(self):
            return False
        def get_volume(self):
            return None
        def set_volume(self, vol, floor=0.0, ceiling=1.0):
            pass
        def adjust_volume(self, delta, floor=0.0, ceiling=1.0):
            pass
'''

replacement = "\nif WINDOWS:" + indented + macos_stubs
src = src[:idx_start] + replacement + src[idx_end:]
print("PATCH 14 OK — list_audio_devices / WindowsAudioClient guarded")

# ─────────────────────────────────────────────────────────────────────────────
# PATCH 15 — guard _acquire_single_instance()
# ─────────────────────────────────────────────────────────────────────────────
old = """\
def _acquire_single_instance() -> bool:
    \"\"\"Create a named Windows mutex so only one VSE can run at once.
    Returns True if we got the lock, False if another instance owns it.\"\"\""""
new = """\
def _acquire_single_instance() -> bool:
    \"\"\"Create a named Windows mutex so only one VSE can run at once.
    Returns True if we got the lock, False if another instance owns it.\"\"\"
    if not WINDOWS:
        return True"""
assert old in src, "PATCH 15 — _acquire_single_instance not found"
src = src.replace(old, new, 1)
print("PATCH 15 OK — single instance guard")

# ─────────────────────────────────────────────────────────────────────────────
# PATCH 16 — main() macOS region selector (skip splash, use mss + selectROI)
# ─────────────────────────────────────────────────────────────────────────────
old = """\
    if not show_splash():
        log.info("Splash closed without starting — exiting")
        return

    hwnd, rel_box = select_region()
    if not hwnd or rel_box['width'] <= 10 or rel_box['height'] <= 10:
        log.warning("Invalid region selected — exiting")
        return

    initial_frame = capture_window_region(hwnd, rel_box)
    if initial_frame is None:
        log.error("Failed to capture window — ensure it is not fully minimised")
        return

    bbox = select_head(initial_frame)"""
new = """\
    if WINDOWS:
        if not show_splash():
            log.info("Splash closed without starting — exiting")
            return
        hwnd, rel_box = select_region()
        if not hwnd or rel_box['width'] <= 10 or rel_box['height'] <= 10:
            log.warning("Invalid region selected — exiting")
            return
        initial_frame = capture_window_region(hwnd, rel_box)
        if initial_frame is None:
            log.error("Failed to capture window — ensure it is not fully minimised")
            return
    else:
        # macOS: pick monitor, screenshot it, draw region with cv2.selectROI
        import tkinter as _tk
        _root = _tk.Tk()
        _root.withdraw()
        with mss() as _sct:
            monitors = _sct.monitors[1:]  # skip monitors[0] (virtual combined)
        if not monitors:
            log.error("No monitors found via mss")
            return
        chosen_mon = monitors[0]
        if len(monitors) > 1:
            # Simple Tk dialog to pick monitor
            _sel = [0]
            _dlg = _tk.Toplevel(_root)
            _dlg.title("Select Monitor")
            _tk.Label(_dlg, text="Which monitor is your video feed on?",
                      font=("Arial", 14)).pack(padx=20, pady=10)
            for _i, _m in enumerate(monitors):
                _tk.Button(
                    _dlg, text=f"Monitor {_i+1}  ({_m['width']}×{_m['height']})",
                    font=("Arial", 12),
                    command=lambda idx=_i: (_sel.__setitem__(0, idx), _dlg.destroy())
                ).pack(fill="x", padx=20, pady=3)
            _root.wait_window(_dlg)
            chosen_mon = monitors[_sel[0]]
        _root.destroy()

        log.info(f"macOS: screenshotting monitor {chosen_mon}")
        with mss() as _sct:
            _grab = _sct.grab(chosen_mon)
        _ss = cv2.cvtColor(np.array(_grab), cv2.COLOR_BGRA2BGR)
        _h, _w = _ss.shape[:2]

        # Draw capture region
        _display = _ss.copy()
        cv2.putText(_display, "Step 1: Draw box around your video feed. ENTER/SPACE to confirm.",
                    (10, 30), cv2.FONT_HERSHEY_SIMPLEX, 0.6, (255, 255, 0), 2)
        _win = "Step 1: Select Region"
        cv2.namedWindow(_win, cv2.WINDOW_AUTOSIZE)
        _roi = cv2.selectROI(_win, _display, fromCenter=False, showCrosshair=False)
        cv2.destroyWindow(_win)
        cv2.waitKey(1)

        _rx, _ry, _rw, _rh = _roi
        if _rw < 10 or _rh < 10:
            log.warning("Invalid region selected — exiting")
            return
        # Handle Retina 2× scaling
        if _rx + _rw > _w * 1.1 or _ry + _rh > _h * 1.1:
            _rx, _ry, _rw, _rh = int(_rx*0.5), int(_ry*0.5), int(_rw*0.5), int(_rh*0.5)

        rel_box = {
            'x1': chosen_mon['left'] + _rx,
            'y1': chosen_mon['top']  + _ry,
            'x2': chosen_mon['left'] + _rx + _rw,
            'y2': chosen_mon['top']  + _ry + _rh,
            'width': _rw, 'height': _rh,
        }
        hwnd = None
        initial_frame = _ss[_ry:_ry+_rh, _rx:_rx+_rw].copy()
        if initial_frame is None or initial_frame.size == 0:
            log.error("Failed to capture region from screenshot")
            return

    bbox = select_head(initial_frame)"""
assert old in src, "PATCH 16 — main() region block not found"
src = src.replace(old, new, 1)
print("PATCH 16 OK — main() macOS region selector")

# ─────────────────────────────────────────────────────────────────────────────
# Also guard the MessageBoxW single-instance popup in main()
# ─────────────────────────────────────────────────────────────────────────────
old = """\
        try:
            import ctypes
            ctypes.windll.user32.MessageBoxW(
                None,
                "VisualStimEdger is already running.\\n\\nClose the existing window first.",
                "VisualStimEdger",
                0x40,  # MB_ICONINFORMATION
            )
        except Exception:
            pass"""
new = """\
        if WINDOWS:
            try:
                import ctypes
                ctypes.windll.user32.MessageBoxW(
                    None,
                    "VisualStimEdger is already running.\\n\\nClose the existing window first.",
                    "VisualStimEdger",
                    0x40,  # MB_ICONINFORMATION
                )
            except Exception:
                pass"""
if old in src:
    src = src.replace(old, new, 1)
    print("PATCH 16b OK — MessageBoxW guard")

# ─────────────────────────────────────────────────────────────────────────────
# Write output
# ─────────────────────────────────────────────────────────────────────────────
DEST.write_text(src, encoding="utf-8")
print(f"\nWrote {DEST} ({DEST.stat().st_size // 1024} KB)")

# Syntax check
import ast
try:
    ast.parse(src)
    print("AST parse: OK")
except SyntaxError as e:
    print(f"AST parse FAILED: {e}")
    sys.exit(1)

print("\nAll patches applied successfully.")
