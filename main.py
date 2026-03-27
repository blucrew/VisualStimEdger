import os
import sys
import tempfile

# Redirect comtypes generated-interface cache to a writable location BEFORE
# pycaw/comtypes are imported.  In a frozen exe (PyInstaller or Nuitka) the
# bundled comtypes/gen directory is read-only, so COM interface generation
# silently fails and device enumeration returns nothing.
import comtypes.gen as _comtypes_gen
_cache = os.path.join(tempfile.gettempdir(), "VisualStimEdger_comtypes_gen")
os.makedirs(_cache, exist_ok=True)
_comtypes_gen.__path__ = [_cache]

import cv2
import numpy as np
import time
import threading
import webbrowser
import requests
import websocket
import tkinter as tk
from tkinter import messagebox, ttk
import customtkinter as ctk
from PIL import Image, ImageTk
import logging
import argparse

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("dark-blue")
from mss import mss
import win32gui
import win32ui
import win32con
import win32api
import json
import pathlib
import queue
from collections import deque
import ctypes
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
from comtypes import CLSCTX_ALL

import atexit

log = logging.getLogger("VisualStimEdger")

_sct = None

def _get_sct():
    """Lazy singleton mss instance, closed cleanly on process exit."""
    global _sct
    if _sct is None:
        _sct = mss()
        atexit.register(_sct.close)
    return _sct


def resource_path(relative):
    """Resolve path to bundled resource — works both in dev and PyInstaller .exe."""
    base = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, relative)


class DickDetector:
    """
    Runs YOLOFastest (Darknet) inference via cv2.dnn to detect 'dick-head'.
    Used to periodically reanchor the CSRT tracker so it can't drift onto hands.
    """
    CLASSES = ["dick", "dick-head"]
    INPUT_SIZE = (320, 320)

    def __init__(self, conf_threshold=0.40, nms_threshold=0.45):
        self.last_conf = 0.0  # confidence of most recent detection (0 if none)
        cfg     = resource_path(os.path.join("models", "yolo-fastest.cfg"))
        weights = resource_path(os.path.join("models", "best.weights"))
        self.conf_threshold = conf_threshold
        self.nms_threshold  = nms_threshold
        self._net = None
        self._output_layers = []
        try:
            net = cv2.dnn.readNetFromDarknet(cfg, weights)
            net.setPreferableBackend(cv2.dnn.DNN_BACKEND_OPENCV)
            net.setPreferableTarget(cv2.dnn.DNN_TARGET_CPU)
            layer_names = net.getLayerNames()
            self._output_layers = [layer_names[i - 1] for i in net.getUnconnectedOutLayers().flatten()]
            self._net = net
            log.info("DickDetector: model loaded OK")
        except Exception as e:
            log.warning(f"DickDetector: failed to load model: {e}")

    @property
    def available(self):
        return self._net is not None

    def detect_head(self, frame):
        """
        Returns the highest-confidence 'dick-head' bbox as (x, y, w, h) in frame
        pixel coordinates, or None if nothing found above the confidence threshold.
        """
        if not self.available:
            return None

        h, w = frame.shape[:2]
        blob = cv2.dnn.blobFromImage(frame, 1 / 255.0, self.INPUT_SIZE,
                                     swapRB=True, crop=False)
        self._net.setInput(blob)
        outputs = self._net.forward(self._output_layers)

        boxes, confidences = [], []
        for output in outputs:
            for det in output:
                scores   = det[5:]
                class_id = int(np.argmax(scores))
                conf     = float(scores[class_id])
                if class_id == 1 and conf >= self.conf_threshold:   # class 1 = dick-head
                    cx = int(det[0] * w)
                    cy = int(det[1] * h)
                    bw = int(det[2] * w)
                    bh = int(det[3] * h)
                    boxes.append([cx - bw // 2, cy - bh // 2, bw, bh])
                    confidences.append(conf)

        if not boxes:
            return None

        indices = cv2.dnn.NMSBoxes(boxes, confidences,
                                   self.conf_threshold, self.nms_threshold)
        if len(indices) == 0:
            return None

        # Return the highest-confidence surviving detection
        best = max(indices.flatten(), key=lambda i: confidences[i])
        self.last_conf = confidences[best]
        return tuple(boxes[best])

# --- CONFIGURATION ---
VERSION = "1.4.1"
GITHUB_REPO = "blucrew/VisualStimEdger"
RESTIM_HOST = '127.0.0.1'
RESTIM_PORT = 12346
TCODE_AXIS = 'L0'
VOLUME_STEP = 0.05
VOLUME_UPDATE_INTERVAL = 0.5


# Aggressiveness levels: name → delta multiplier
AGGR_LEVELS = {
    "Easy":   0.4,
    "Middle": 1.0,
    "Hard":   2.0,
    "Expert": 4.0,
}

CONFIG_PATH = pathlib.Path(os.environ.get("APPDATA", ".")) / "VisualStimEdger" / "config.json"

HEAD_Y_SMOOTH      = 8    # rolling average window for head Y before volume logic

class RegionSelector:
    def __init__(self, parent=None):
        if parent is None:
            self.root = tk.Tk()
            self.is_main = True
        else:
            self.root = tk.Toplevel(parent)
            self.is_main = False
            
        self.root.attributes('-alpha', 0.4)
        
        with mss() as sct:
            mon = sct.monitors[0]  # monitors[0] = combined virtual screen (all monitors)
            self.offset_x = mon["left"]
            self.offset_y = mon["top"]
            w, h = mon["width"], mon["height"]

        # Set size first, then use MoveWindow to handle negative coords
        # (secondary monitor left of primary gives negative offset_x which
        # tkinter geometry strings like "+-1920+0" cannot express correctly).
        self.root.geometry(f"{w}x{h}+0+0")
        self.root.update_idletasks()
        ctypes.windll.user32.MoveWindow(
            self.root.winfo_id(), self.offset_x, self.offset_y, w, h, False)
        self.root.overrideredirect(True)
        
        self.root.configure(background='black')
        self.root.attributes("-topmost", True)
        self.root.config(cursor="cross")

        self.canvas = tk.Canvas(self.root, cursor="cross", bg="black")
        self.canvas.pack(fill="both", expand=True)

        self.canvas.bind("<ButtonPress-1>", self.on_press)
        self.canvas.bind("<B1-Motion>", self.on_drag)
        self.canvas.bind("<ButtonRelease-1>", self.on_release)

        self.start_x = None
        self.start_y = None
        self.rect = None
        self.region = None
        
        self.label = tk.Label(self.root, text="Step 1: Draw a box around the video feed on any monitor. Release to lock.", font=("Arial", 28), bg="white", fg="black")
        self.label.pack(pady=50)
        self.root.bind("<Escape>", lambda e: self.root.destroy())

    def on_press(self, event):
        self.start_x = self.canvas.canvasx(event.x)
        self.start_y = self.canvas.canvasy(event.y)
        if self.rect:
            self.canvas.delete(self.rect)
        self.rect = self.canvas.create_rectangle(self.start_x, self.start_y, 1, 1, outline='red', width=3, fill="gray50")

    def on_drag(self, event):
        cur_x, cur_y = (self.canvas.canvasx(event.x), self.canvas.canvasy(event.y))
        self.canvas.coords(self.rect, self.start_x, self.start_y, cur_x, cur_y)

    def on_release(self, event):
        end_x, end_y = (self.canvas.canvasx(event.x), self.canvas.canvasy(event.y))
        x1, x2 = sorted([self.start_x, end_x])
        y1, y2 = sorted([self.start_y, end_y])
        
        abs_x1 = int(x1) + self.offset_x
        abs_y1 = int(y1) + self.offset_y
        
        self.region = {'top': abs_y1, 'left': abs_x1, 'width': int(x2 - x1), 'height': int(y2 - y1)}
        self.root.destroy()

def select_region(parent=None):
    selector = RegionSelector(parent)
    if selector.is_main:
        selector.root.mainloop()
    else:
        parent.wait_window(selector.root)
    
    if not selector.region:
        return None, None
        
    time.sleep(1.0)
    
    x1, y1 = selector.region['left'], selector.region['top']
    x2, y2 = x1 + selector.region['width'], y1 + selector.region['height']
    
    cx = (x1 + x2) // 2
    cy = (y1 + y2) // 2
    
    hwnd = win32gui.WindowFromPoint((cx, cy))
    hwnd = win32gui.GetAncestor(hwnd, win32con.GA_ROOT)
    
    wx, wy, wr, wb = win32gui.GetWindowRect(hwnd)
    
    rel_box = {
        'x1': x1 - wx,
        'y1': y1 - wy,
        'x2': x2 - wx,
        'y2': y2 - wy,
        'width': x2 - x1,
        'height': y2 - y1
    }
    return hwnd, rel_box

def capture_window_region(hwnd, rel_box):
    try:
        left, top, right, bot = win32gui.GetWindowRect(hwnd)
        w = right - left
        h = bot - top
        if w <= 0 or h <= 0:
            return None

        # ── PrintWindow (works for occluded windows) ──────────────────────────
        # GDI objects are cleaned up in finally so they never leak, even when
        # GetBitmapBits / reshape / cvtColor throw on bad GPU-window bitmaps.
        hwndDC = saveDC = mfcDC = saveBitMap = None
        frame = None
        try:
            hwndDC    = win32gui.GetWindowDC(hwnd)
            mfcDC     = win32ui.CreateDCFromHandle(hwndDC)
            saveDC    = mfcDC.CreateCompatibleDC()
            saveBitMap = win32ui.CreateBitmap()
            saveBitMap.CreateCompatibleBitmap(mfcDC, w, h)
            saveDC.SelectObject(saveBitMap)

            result = ctypes.windll.user32.PrintWindow(hwnd, saveDC.GetSafeHdc(), 2)
            if result == 1:
                bmpinfo = saveBitMap.GetInfo()
                bmpstr  = saveBitMap.GetBitmapBits(True)
                img     = np.frombuffer(bmpstr, dtype=np.uint8).reshape(
                              (bmpinfo['bmHeight'], bmpinfo['bmWidth'], 4))
                frame   = cv2.cvtColor(img, cv2.COLOR_BGRA2BGR)
        except Exception as e:
            log.debug(f"PrintWindow failed: {e}")
            frame = None
        finally:
            if saveBitMap: win32gui.DeleteObject(saveBitMap.GetHandle())
            if saveDC:     saveDC.DeleteDC()
            if mfcDC:      mfcDC.DeleteDC()
            if hwndDC:     win32gui.ReleaseDC(hwnd, hwndDC)

        # ── Crop the PrintWindow result ───────────────────────────────────────
        if frame is not None and frame.any():
            x1 = max(0, min(w, rel_box['x1']))
            y1 = max(0, min(h, rel_box['y1']))
            x2 = max(0, min(w, rel_box['x2']))
            y2 = max(0, min(h, rel_box['y2']))
            if x2 - x1 > 0 and y2 - y1 > 0:
                return frame[y1:y2, x1:x2].copy()

        # ── mss fallback (screen-level grab — works for GPU windows on-screen) ─
        abs_x1 = left + rel_box['x1']
        abs_y1 = top  + rel_box['y1']
        monitor = {
            "top":    abs_y1,
            "left":   abs_x1,
            "width":  rel_box['width'],
            "height": rel_box['height'],
        }
        grab = _get_sct().grab(monitor)
        return cv2.cvtColor(np.array(grab), cv2.COLOR_BGRA2BGR)

    except Exception as e:
        log.debug(f"capture_window_region: {e}")
        return None

def select_head(frame_cv, parent=None):
    if parent is None:
        root = tk.Tk()
        is_main = True
    else:
        root = tk.Toplevel(parent)
        is_main = False
    
    root.title("Step 2: Select Cock Head")
    root.attributes("-topmost", True)
    
    frame_rgb = cv2.cvtColor(frame_cv, cv2.COLOR_BGR2RGB)
    img = Image.fromarray(frame_rgb)
    
    canvas = tk.Canvas(root, width=img.width, height=img.height, cursor="cross")
    canvas.pack()
    
    photo = ImageTk.PhotoImage(image=img)
    canvas.image = photo # Keep reference to prevent garbage collection!
    canvas.create_image(0, 0, image=photo, anchor=tk.NW)
    
    state = {'start_x': None, 'start_y': None, 'rect': None, 'bbox': (0, 0, 0, 0)}
    
    mw = img.width // 2
    mh = img.height // 2
    # Pre-draw explicit centered default fallback
    state['rect'] = canvas.create_rectangle(mw-30, mh-30, mw+30, mh+30, outline='green', width=2)
    state['bbox'] = (mw-30, mh-30, 60, 60)
    
    def on_press(event):
        state['start_x'] = event.x
        state['start_y'] = event.y
        if state['rect']:
            canvas.delete(state['rect'])
        state['rect'] = canvas.create_rectangle(event.x, event.y, event.x, event.y, outline='green', width=2)
        
    def on_drag(event):
        if state['rect'] and state['start_x'] is not None and state['start_y'] is not None:
            canvas.coords(state['rect'], state['start_x'], state['start_y'], event.x, event.y)
        
    def on_release(event):
        if state['start_x'] is not None and state['start_y'] is not None:
            x1, y1 = state['start_x'], state['start_y']
            x2, y2 = event.x, event.y
            state['bbox'] = (int(min(x1, x2)), int(min(y1, y2)), int(abs(x2-x1)), int(abs(y2-y1)))
        
    canvas.bind("<ButtonPress-1>", on_press)
    canvas.bind("<B1-Motion>", on_drag)
    canvas.bind("<ButtonRelease-1>", on_release)

    def _cancel():
        state['bbox'] = (0, 0, 0, 0)
        root.destroy()

    root.bind("<Escape>", lambda e: _cancel())
    root.protocol("WM_DELETE_WINDOW", _cancel)

    btn_frame = tk.Frame(root)
    btn_frame.pack(fill=tk.X)
    tk.Button(btn_frame, text="Confirm Head Area ✅", command=root.destroy, font=("Arial", 12, "bold"), bg="#4CAF50", fg="white").pack(pady=5)
    
    if is_main:
        root.eval('tk::PlaceWindow . center')
        root.mainloop()
    else:
        # Centers Toplevel relative to system active screen
        root.update_idletasks()
        w = root.winfo_width()
        h = root.winfo_height()
        x = root.winfo_screenwidth() // 2 - w // 2
        y = root.winfo_screenheight() // 2 - h // 2
        root.geometry(f"+{x}+{y}")
        parent.wait_window(root)
        
    return state['bbox']

class RestimClient:
    _BACKOFF_INITIAL = 1.0
    _BACKOFF_MAX     = 30.0

    def __init__(self, host, port, axis):
        self.host   = host
        self.port   = port
        self.axis   = axis
        self.volume = 0.5
        self.ws     = None
        self._lock        = threading.Lock()
        self._connecting  = False
        self._backoff     = self._BACKOFF_INITIAL
        self._next_attempt = 0.0  # connect immediately on first call

    def maybe_reconnect(self):
        """Call from the main thread. Spawns a connect thread when backoff allows."""
        with self._lock:
            if self.ws is not None or self._connecting:
                return
            if time.time() < self._next_attempt:
                return
            self._connecting = True
        threading.Thread(target=self._connect_bg, daemon=True).start()

    def _connect_bg(self):
        try:
            ws_url = f"ws://{self.host}:{self.port}"
            ws = websocket.create_connection(ws_url, timeout=2.0)
            with self._lock:
                self.ws         = ws
                self._backoff   = self._BACKOFF_INITIAL  # reset on success
                self._connecting = False
            log.info(f"Restim: connected at {ws_url}")
            self.set_volume(self.volume, instant=True)
        except Exception as e:
            with self._lock:
                self._backoff      = min(self._backoff * 2, self._BACKOFF_MAX)
                self._next_attempt = time.time() + self._backoff
                self._connecting   = False
            log.debug(f"Restim: connect failed, retry in {self._backoff:.0f}s: {e}")

    def set_volume(self, vol, instant=False, floor=0.0, ceiling=1.0):
        with self._lock:
            self.volume = max(floor, min(ceiling, vol))
            if self.ws:
                try:
                    val_int  = int(round(self.volume * 9999))
                    interval = 0 if instant else int(VOLUME_UPDATE_INTERVAL * 1000)
                    self.ws.send(f"{self.axis}{val_int:04d}I{interval}")
                except Exception:
                    self.ws = None

    def adjust_volume(self, delta, floor=0.0, ceiling=1.0):
        self.set_volume(self.volume + delta, floor=floor, ceiling=ceiling)


class _DefaultAudioDevice:
    """Minimal stand-in returned when full device enumeration fails."""
    FriendlyName = "Default Output Device"
    def __init__(self, dev):
        self._dev = dev


class _SounddeviceAudioDevice:
    """Device entry populated via sounddevice when all pycaw enumeration paths fail.
    _dev is intentionally None — WindowsAudioClient resolves it by name at init time."""
    def __init__(self, name):
        self.FriendlyName = name
        self._dev = None


def list_audio_devices():
    """Return all render (output) devices.

    sounddevice is tried first — it bypasses COM entirely and is rock-solid
    for listing.  pycaw is kept only as a fallback for volume control resolution.
    """
    # Level 1: sounddevice — most reliable for enumeration across all Windows versions
    try:
        import sounddevice as sd
        devs = [_SounddeviceAudioDevice(d['name'])
                for d in sd.query_devices()
                if d['max_output_channels'] > 0]
        if devs:
            log.info(f"WinAudio: sounddevice found {len(devs)} output device(s)")
            return devs
    except Exception as e:
        log.error(f"WinAudio: sounddevice listing failed: {e}")

    # Level 2 & 3: pycaw full enumeration (fallback)
    log.warning("WinAudio: sounddevice unavailable — falling back to pycaw")
    try:
        devices = AudioUtilities.GetAllDevices()
        render = [d for d in devices if int(d.flow) == 0]
        if render:
            return render
        if devices:
            return list(devices)
    except Exception as e:
        log.error(f"WinAudio: GetAllDevices failed: {e}")

    # Level 4: default speakers only
    try:
        default = AudioUtilities.GetSpeakers()
        if default:
            return [_DefaultAudioDevice(default)]
    except Exception as e:
        log.error(f"WinAudio: GetSpeakers fallback also failed: {e}")

    return []


class WindowsAudioClient:
    def __init__(self, device):
        self._volume_interface = None
        dev = device._dev

        # _dev is None when we came via the sounddevice fallback path.
        # Try to find the matching IMMDevice by name; fall back to default speakers.
        if dev is None:
            dev = self._resolve_dev(device.FriendlyName)

        if dev is None:
            log.error(f"WinAudio: could not resolve IMMDevice for '{device.FriendlyName}'")
            return
        try:
            interface = dev.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
            self._volume_interface = interface.QueryInterface(IAudioEndpointVolume)
            log.info(f"WinAudio: connected to {device.FriendlyName}")
        except Exception as e:
            log.error(f"WinAudio: failed to activate device: {e}")

    @staticmethod
    def _resolve_dev(name):
        """Find an IMMDevice for the given friendly name.

        Tries in order:
        1. Exact match against GetAllDevices() FriendlyName.
        2. Substring match — handles PortAudio vs WASAPI naming differences
           e.g. "Speakers (Realtek)" vs "Speakers (Realtek(R) Audio)".
        3. First available render device.
        4. GetSpeakers() — default output endpoint.
        """
        try:
            all_devs = [d for d in AudioUtilities.GetAllDevices() if d._dev is not None]
            # Pass 1: exact
            for d in all_devs:
                if d.FriendlyName == name:
                    return d._dev
            # Pass 2: fuzzy — one name is a substring of the other
            nl = name.lower()
            for d in all_devs:
                fl = d.FriendlyName.lower()
                if nl in fl or fl in nl:
                    log.debug(f"WinAudio: fuzzy matched '{name}' → '{d.FriendlyName}'")
                    return d._dev
            # Pass 3: first render device
            render = [d for d in all_devs if int(d.flow) == 0]
            if render:
                log.debug(f"WinAudio: no name match, using first render device '{render[0].FriendlyName}'")
                return render[0]._dev
        except Exception as e:
            log.debug(f"WinAudio: _resolve_dev enumeration error: {e}")
        # Pass 4: default speakers
        try:
            return AudioUtilities.GetSpeakers()
        except Exception:
            return None

    @property
    def connected(self):
        return self._volume_interface is not None

    def get_volume(self):
        try:
            return self._volume_interface.GetMasterVolumeLevelScalar()
        except Exception:
            return 0.0

    def set_volume(self, vol, floor=0.0, ceiling=1.0):
        vol = max(floor, min(ceiling, vol))
        try:
            self._volume_interface.SetMasterVolumeLevelScalar(vol, None)
        except Exception as e:
            log.warning(f"WinAudio: set_volume failed: {e}")

    def adjust_volume(self, delta, floor=0.0, ceiling=1.0):
        self.set_volume(self.get_volume() + delta, floor=floor, ceiling=ceiling)


def check_for_update(on_update_available):
    """Runs in a background thread. Calls on_update_available(latest_version, url) if a newer release exists."""
    try:
        resp = requests.get(
            f"https://api.github.com/repos/{GITHUB_REPO}/releases/latest",
            timeout=5,
            headers={"Accept": "application/vnd.github+json"}
        )
        if resp.status_code != 200:
            return
        data = resp.json()
        latest = data.get("tag_name", "").lstrip("v")
        url = data.get("html_url", f"https://github.com/{GITHUB_REPO}/releases/latest")
        try:
            # Guard against pre-release tags like "1.2.0-beta.1"
            latest_tuple  = tuple(int(x) for x in latest.split(".")[:3] if x.isdigit())
            current_tuple = tuple(int(x) for x in VERSION.split(".")[:3] if x.isdigit())
            if latest_tuple and latest_tuple > current_tuple:
                on_update_available(latest, url)
        except Exception:
            pass
    except Exception:
        pass  # silently ignore — no internet, rate limit, etc.


# ── MP3 player ────────────────────────────────────────────────────────────────
try:
    import miniaudio as _miniaudio
    _MINIAUDIO_OK = True
except ImportError:
    _miniaudio = None
    _MINIAUDIO_OK = False
    log.warning("miniaudio not installed — MP3 mode unavailable")


class MusicPlayer:
    """Streams audio files with per-chunk volume control via miniaudio."""

    EXTS = {'.mp3', '.wav', '.ogg', '.flac', '.m4a'}

    def __init__(self):
        self._device    = None
        self._stop_flag = False
        self._state     = "stopped"   # stopped | playing | paused
        self._skip      = 0           # +1 next, -1 prev (set from main thread)
        self._playlist  = []          # list[pathlib.Path]
        self._idx       = 0
        self.loop_mode  = "folder"    # "track" | "folder"
        self.volume     = 0.5         # 0.0–1.0, written from main thread, read from audio thread
        # UI notification — set by audio thread, read by main thread (strings are atomic)
        self.track_name = ""
        self.track_info = ""          # e.g. "3 / 12"

    # ── Generator ─────────────────────────────────────────────────────────────
    def _gen(self):
        required_frames = yield b""   # prime
        while not self._stop_flag:
            if not self._playlist or self._state == "stopped":
                required_frames = yield bytes(required_frames * 8)  # silence
                continue

            path = self._playlist[self._idx]
            self.track_name = path.stem
            self.track_info = f"{self._idx + 1} / {len(self._playlist)}"

            try:
                src = _miniaudio.stream_file(
                    str(path),
                    output_format=_miniaudio.SampleFormat.FLOAT32,
                    nchannels=2, sample_rate=44100,
                    frames_to_read=4096,
                )
            except Exception as e:
                log.error(f"MusicPlayer: cannot open {path.name}: {e}")
                self._idx = (self._idx + 1) % len(self._playlist)
                continue

            for chunk in src:
                if self._stop_flag:
                    return

                # Handle skip (next/prev)
                skip = self._skip
                if skip:
                    self._skip = 0
                    self._idx  = (self._idx + skip) % len(self._playlist)
                    break

                # Pause — yield silence, stay in loop
                while self._state == "paused" and not self._stop_flag and not self._skip:
                    required_frames = yield bytes(required_frames * 8)

                if self._stop_flag:
                    return

                vol  = max(0.0, min(1.0, self.volume))
                data = np.frombuffer(chunk.samples, dtype=np.float32).copy()
                data *= vol
                required_frames = yield data.tobytes()
            else:
                # Natural track end
                if self.loop_mode == "track":
                    continue          # replay same index
                self._idx = (self._idx + 1) % len(self._playlist)

    # ── Device control ────────────────────────────────────────────────────────
    def _start(self):
        self._stop_flag = False
        gen = self._gen()
        next(gen)
        try:
            self._device = _miniaudio.PlaybackDevice(
                output_format=_miniaudio.SampleFormat.FLOAT32,
                nchannels=2, sample_rate=44100,
                buffersize_msec=150,
            )
            self._device.start(gen)
        except Exception as e:
            log.error(f"MusicPlayer: device start failed: {e}")
            self._device = None

    def _stop_device(self):
        self._stop_flag = True
        self._state     = "stopped"
        if self._device:
            try:
                self._device.stop()
            except Exception:
                pass
            self._device = None

    # ── Public API ────────────────────────────────────────────────────────────
    def load_file(self, path: str):
        self._stop_device()
        self._playlist  = [pathlib.Path(path)]
        self._idx       = 0
        self.loop_mode  = "track"
        self._state     = "playing"
        self._start()

    def load_folder(self, folder: str):
        files = sorted(
            [f for f in pathlib.Path(folder).iterdir()
             if f.is_file() and f.suffix.lower() in self.EXTS],
            key=lambda f: f.name.lower(),
        )
        if not files:
            log.warning("MusicPlayer: no audio files in folder")
            return
        self._stop_device()
        self._playlist = files
        self._idx      = 0
        self.loop_mode = "folder"
        self._state    = "playing"
        self._start()

    def play(self):
        if self._state == "paused":
            self._state = "playing"
        elif self._state == "stopped" and self._playlist:
            self._state = "playing"
            if not self._device:
                self._start()

    def pause(self):
        if self._state == "playing":
            self._state = "paused"

    def stop(self):
        self._stop_device()

    def next_track(self):
        if self._playlist:
            self._skip = 1

    def prev_track(self):
        if self._playlist:
            self._skip = -1

    def adjust_volume(self, delta: float, floor: float, ceiling: float):
        self.volume = max(floor, min(ceiling, self.volume + delta))

    def cleanup(self):
        self._stop_device()


class App:
    # ── Colour palette ────────────────────────────────────────────────────────
    _C_BG        = "#0d0d0d"
    _C_SURFACE   = "#1a1a1a"
    _C_SURFACE2  = "#242424"
    _C_RED       = "#cc2200"
    _C_RED_HOV   = "#991800"
    _C_YELLOW    = "#ffcc00"
    _C_YELLOW_H  = "#e6b500"
    _C_TEXT      = "#eeeeee"
    _C_TEXT_DIM  = "#666666"
    _C_BORDER    = "#2e2e2e"

    # ── YOLO reanchoring
    _YOLO_INTERVAL = 15    # run detector every N frames
    _YOLO_CONFIRM  = 2     # consecutive detections in same area before reanchoring
    _YOLO_MAX_JUMP = 2.0   # max allowed jump as multiple of current bbox diagonal
    # Tracker plausibility
    _SIZE_RATIO_MAX  = 2.5
    _MAX_JUMP_FACTOR = 2.5

    def __init__(self, hwnd, rel_box, initial_frame, bbox):
        # Ensure COM is initialised on this thread (fixes pycaw failures on some Win11 setups)
        try:
            from comtypes import CoInitialize
            CoInitialize()
        except Exception:
            pass

        # Capture / window
        self.hwnd    = hwnd
        self.rel_box = rel_box

        # Tracking state
        self.head_y          = bbox[1] + bbox[3] // 2
        self.heights         = {"Edging": None, "Erect": None, "Flaccid": None}

        self.last_vol_time   = time.time()
        self.tracking_paused    = False
        self.last_bbox          = tuple(int(v) for v in bbox)
        self.tracking_ok        = True
        self._track_msg         = ""
        self.yolo_frame_counter = 0
        self.yolo_candidate     = None  # (bbox, hits) pending confirmation
        self._head_y_history    = deque(maxlen=HEAD_Y_SMOOTH)
        self._frame_times       = deque(maxlen=30)  # for FPS calculation

        # Session stats
        self.session_start   = time.time()
        self.state_times     = {"Edging": 0.0, "Erect": 0.0, "Flaccid": 0.0}
        self.edge_count      = 0
        self._prev_state     = "Erect"
        self._last_state_time = time.time()
        self._stats_tick     = 0   # throttle label updates

        # Hold
        self.hold_active = False

        # Cum cooldown  (None = not active)
        self._cum_time: float | None = None

        # AUTO calibration
        self._auto_mode       = True
        self._auto_min_y: float | None = None
        self._auto_max_y: float | None = None
        self._auto_obs_start: float | None = None
        self._auto_last_apply = 0.0

        # Background capture — keeps the main thread free for tracking + UI
        self._frame_queue    = queue.Queue(maxsize=2)
        self._running        = True
        self._capture_thread = threading.Thread(target=self._capture_loop, daemon=True)
        self._capture_thread.start()

        # CV
        self.tracker  = cv2.TrackerCSRT_create()
        self.tracker.init(initial_frame, bbox)
        self.detector = DickDetector()

        # Output clients
        self.restim           = RestimClient(RESTIM_HOST, RESTIM_PORT, TCODE_AXIS)
        self.win_audio        = None
        self.win_devices      = []
        self._orig_win_volume = None  # restored on exit
        self.music_player     = MusicPlayer() if _MINIAUDIO_OK else None

        # Root window
        self.root = ctk.CTk()
        self.root.title("VisualStimEdger")
        _icon = pathlib.Path(resource_path("icon.ico"))
        if _icon.exists():
            self.root.iconbitmap(str(_icon))

        # tkinter vars — must be created after root exists
        self.min_vol_var = tk.DoubleVar(value=0.0)
        self.max_vol_var = tk.DoubleVar(value=100.0)
        self.aggr_var    = tk.StringVar(value="Middle")
        self.mode_var    = tk.StringVar(value="restim")
        self.port_var    = tk.StringVar(value="12346")

        self._build_ui()
        self._on_aggr_change()   # set initial aggressiveness colour
        self._load_config()
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        self._start_update_check()

    def run(self):
        self.root.after(5, self._update_frame)
        self.root.mainloop()

    # ------------------------------------------------------------------ UI

    def _build_ui(self):
        root = self.root
        root.configure(fg_color=self._C_BG)

        P   = 12          # standard outer padding
        lbl = ctk.CTkFont(size=11, weight="bold")
        btn = ctk.CTkFont(size=12, weight="bold")

        # ── Update banner (hidden until needed) ───────────────────────────────
        self._update_banner = ctk.CTkFrame(root, fg_color="#7a5c00", corner_radius=0)
        self._update_label  = ctk.CTkLabel(self._update_banner, text="",
                                           text_color="white", font=lbl)
        self._update_label.pack(side=tk.LEFT, padx=P, pady=6)
        self._update_btn = ctk.CTkButton(self._update_banner, text="Download", width=100,
                                         fg_color="#b8860b", hover_color="#8B6914",
                                         text_color="white", font=lbl, corner_radius=4)
        self._update_btn.pack(side=tk.RIGHT, padx=P, pady=6)

        # ── Video + calibration buttons ───────────────────────────────────────
        top_frame = ctk.CTkFrame(root, fg_color="transparent")
        top_frame.pack(padx=P, pady=P, fill=tk.X)
        self._first_widget = top_frame

        # video — plain tk.Label so ImageTk works without wrapping
        vid_shell = ctk.CTkFrame(top_frame, fg_color=self._C_SURFACE,
                                 corner_radius=8, border_width=1, border_color=self._C_BORDER)
        vid_shell.pack(side=tk.LEFT)
        self.video_label = tk.Label(vid_shell, bg=self._C_SURFACE)
        self.video_label.pack(padx=3, pady=3)

        # height buttons — stacked right of video
        hbf = ctk.CTkFrame(top_frame, fg_color="transparent")
        hbf.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(10, 0))

        def _hbtn(parent, text, cmd, color, hover):
            ctk.CTkButton(parent, text=text, command=cmd, font=btn, height=38,
                          fg_color=color, hover_color=hover,
                          text_color="white", corner_radius=6).pack(fill=tk.X, pady=3)

        self._auto_btn = ctk.CTkButton(
            hbf, text="AUTO  (observing...)", command=self._toggle_auto,
            font=btn, height=38, corner_radius=6,
            fg_color="#b8a000", hover_color="#8a7800", text_color="white")
        self._auto_btn.pack(fill=tk.X, pady=(0, 6))

        _hbtn(hbf, "Set Edging Height",  self._set_edging,  self._C_RED,    self._C_RED_HOV)
        _hbtn(hbf, "Set Erect Height",   self._set_erect,   "#1a5c2e",      "#114420")
        _hbtn(hbf, "Set Flaccid Height", self._set_flaccid, "#1a2e5c",      "#11203e")

        # spacer pushes cum button to bottom
        ctk.CTkFrame(hbf, fg_color="transparent").pack(fill=tk.BOTH, expand=True)
        ctk.CTkButton(hbf, text="I've CUM", command=self._on_cum,
                      font=btn, height=38, corner_radius=6,
                      fg_color="white", hover_color="#dddddd",
                      text_color="#0d0d0d").pack(fill=tk.X, pady=3)

        def _divider():
            ctk.CTkFrame(root, height=1, fg_color=self._C_BORDER).pack(fill=tk.X, padx=P, pady=3)

        _divider()

        # ── Volume floor / ceiling ────────────────────────────────────────────
        vol_card = ctk.CTkFrame(root, fg_color=self._C_SURFACE, corner_radius=8)
        vol_card.pack(fill=tk.X, padx=P, pady=4)

        def _slider_row(parent, label_text, var, val_attr, callback):
            row = ctk.CTkFrame(parent, fg_color="transparent")
            row.pack(fill=tk.X, padx=12, pady=(6, 3))
            ctk.CTkLabel(row, text=label_text, font=lbl,
                         text_color=self._C_TEXT, width=90, anchor="w").pack(side=tk.LEFT)
            vl = ctk.CTkLabel(row, text=f"{var.get():.0f}%", font=lbl,
                              text_color=self._C_YELLOW, width=44, anchor="e")
            vl.pack(side=tk.RIGHT)
            setattr(self, val_attr, vl)
            ctk.CTkSlider(row, from_=0, to=100, variable=var, command=callback,
                          button_color=self._C_YELLOW, button_hover_color=self._C_YELLOW_H,
                          progress_color=self._C_YELLOW, fg_color=self._C_SURFACE2,
                          width=220).pack(side=tk.RIGHT, padx=(0, 8))

        _slider_row(vol_card, "Vol Floor",   self.min_vol_var, "_floor_lbl", self._on_floor_change)
        _slider_row(vol_card, "Vol Ceiling", self.max_vol_var, "_ceil_lbl",  self._on_ceil_change)

        _divider()

        # ── Aggressiveness ────────────────────────────────────────────────────
        aggr_card = ctk.CTkFrame(root, fg_color=self._C_SURFACE, corner_radius=8)
        aggr_card.pack(fill=tk.X, padx=P, pady=4)
        aggr_row = ctk.CTkFrame(aggr_card, fg_color="transparent")
        aggr_row.pack(fill=tk.X, padx=12, pady=10)
        ctk.CTkLabel(aggr_row, text="Aggressiveness", font=lbl,
                     text_color=self._C_TEXT).pack(side=tk.LEFT)
        self._aggr_seg = ctk.CTkSegmentedButton(
            aggr_row, values=list(AGGR_LEVELS.keys()), variable=self.aggr_var,
            command=self._on_aggr_change,
            selected_color="#1a8a1a", selected_hover_color="#145c14",
            unselected_color=self._C_SURFACE2, unselected_hover_color="#333",
            font=ctk.CTkFont(size=11, weight="bold"),
            text_color=self._C_TEXT,
        )
        self._aggr_seg.pack(side=tk.RIGHT)

        _divider()

        # ── Output mode ───────────────────────────────────────────────────────
        mode_card = ctk.CTkFrame(root, fg_color=self._C_SURFACE, corner_radius=8)
        self._mode_card = mode_card
        mode_card.pack(fill=tk.X, padx=P, pady=4)
        mode_row = ctk.CTkFrame(mode_card, fg_color="transparent")
        mode_row.pack(fill=tk.X, padx=12, pady=10)
        ctk.CTkLabel(mode_row, text="Output", font=lbl,
                     text_color=self._C_TEXT).pack(side=tk.LEFT)
        _mode_values = ["restim", "windows"] + (["mp3"] if _MINIAUDIO_OK else [])
        ctk.CTkSegmentedButton(
            mode_row, values=_mode_values, variable=self.mode_var,
            command=lambda _: self._on_mode_change(),
            selected_color=self._C_RED, selected_hover_color=self._C_RED_HOV,
            unselected_color=self._C_SURFACE2, unselected_hover_color="#333",
            font=ctk.CTkFont(size=11), text_color=self._C_TEXT,
        ).pack(side=tk.RIGHT)

        # Restim options panel
        self._restim_opts = ctk.CTkFrame(root, fg_color=self._C_SURFACE2, corner_radius=6)
        ro = ctk.CTkFrame(self._restim_opts, fg_color="transparent")
        ro.pack(padx=12, pady=7)
        ctk.CTkLabel(ro, text="Port:", font=lbl, text_color=self._C_TEXT).pack(side=tk.LEFT, padx=(0, 6))
        ctk.CTkEntry(ro, textvariable=self.port_var, width=72,
                     fg_color=self._C_SURFACE, border_color=self._C_BORDER,
                     text_color=self._C_TEXT).pack(side=tk.LEFT)
        self.port_var.trace_add("write", self._on_port_change)

        # Windows Audio options panel
        self._windows_opts = ctk.CTkFrame(root, fg_color=self._C_SURFACE2, corner_radius=6)
        wo = ctk.CTkFrame(self._windows_opts, fg_color="transparent")
        wo.pack(padx=12, pady=7, fill=tk.X)
        ctk.CTkLabel(wo, text="Device:", font=lbl, text_color=self._C_TEXT).pack(side=tk.LEFT, padx=(0, 6))
        self._device_combo = ctk.CTkComboBox(
            wo, values=[], command=self._on_device_select, width=280,
            fg_color=self._C_SURFACE, border_color=self._C_BORDER,
            button_color=self._C_RED, button_hover_color=self._C_RED_HOV,
            dropdown_fg_color=self._C_SURFACE2, text_color=self._C_TEXT,
        )
        self._device_combo.pack(side=tk.LEFT, padx=(0, 6))
        ctk.CTkButton(wo, text="Refresh", command=self._refresh_devices, width=72,
                      fg_color=self._C_SURFACE2, hover_color="#333",
                      text_color=self._C_TEXT, border_width=1, border_color=self._C_BORDER,
                      font=ctk.CTkFont(size=10)).pack(side=tk.LEFT)

        self._restim_opts.pack(fill=tk.X, padx=P, pady=(0, 4))

        # ── MP3 player options panel ───────────────────────────────────────────
        self._mp3_opts = ctk.CTkFrame(root, fg_color=self._C_SURFACE2, corner_radius=6)
        if _MINIAUDIO_OK:
            mo = ctk.CTkFrame(self._mp3_opts, fg_color="transparent")
            mo.pack(fill=tk.X, padx=12, pady=(8, 4))

            # File / folder load buttons
            load_row = ctk.CTkFrame(mo, fg_color="transparent")
            load_row.pack(fill=tk.X, pady=(0, 4))
            ctk.CTkButton(load_row, text="📁 Load File", width=110, height=28,
                          fg_color=self._C_SURFACE, hover_color="#333",
                          text_color=self._C_TEXT, border_width=1, border_color=self._C_BORDER,
                          font=ctk.CTkFont(size=10),
                          command=self._mp3_load_file).pack(side=tk.LEFT, padx=(0, 6))
            ctk.CTkButton(load_row, text="📂 Load Folder", width=120, height=28,
                          fg_color=self._C_SURFACE, hover_color="#333",
                          text_color=self._C_TEXT, border_width=1, border_color=self._C_BORDER,
                          font=ctk.CTkFont(size=10),
                          command=self._mp3_load_folder).pack(side=tk.LEFT, padx=(0, 10))
            self._mp3_track_lbl = ctk.CTkLabel(load_row, text="No file loaded",
                                               font=ctk.CTkFont(size=10),
                                               text_color=self._C_TEXT_DIM,
                                               anchor="w")
            self._mp3_track_lbl.pack(side=tk.LEFT, fill=tk.X, expand=True)

            # Transport controls
            ctrl_row = ctk.CTkFrame(mo, fg_color="transparent")
            ctrl_row.pack(pady=(0, 4))
            btn_kw = dict(width=44, height=32, fg_color=self._C_SURFACE,
                          hover_color="#333", text_color=self._C_TEXT,
                          border_width=1, border_color=self._C_BORDER,
                          font=ctk.CTkFont(size=13))
            ctk.CTkButton(ctrl_row, text="⏮", command=self._mp3_prev, **btn_kw).pack(side=tk.LEFT, padx=2)
            self._mp3_play_btn = ctk.CTkButton(ctrl_row, text="▶", command=self._mp3_play_pause, **btn_kw)
            self._mp3_play_btn.pack(side=tk.LEFT, padx=2)
            ctk.CTkButton(ctrl_row, text="⏹", command=self._mp3_stop, **btn_kw).pack(side=tk.LEFT, padx=2)
            ctk.CTkButton(ctrl_row, text="⏭", command=self._mp3_next, **btn_kw).pack(side=tk.LEFT, padx=2)

            # Loop mode
            loop_row = ctk.CTkFrame(mo, fg_color="transparent")
            loop_row.pack(pady=(0, 6))
            ctk.CTkLabel(loop_row, text="Loop:", font=ctk.CTkFont(size=10),
                         text_color=self._C_TEXT_DIM).pack(side=tk.LEFT, padx=(0, 6))
            self._mp3_loop_var = tk.StringVar(value="folder")
            ctk.CTkSegmentedButton(
                loop_row, values=["track", "folder"],
                variable=self._mp3_loop_var,
                command=self._mp3_on_loop_change,
                selected_color=self._C_RED, selected_hover_color=self._C_RED_HOV,
                unselected_color=self._C_SURFACE, unselected_hover_color="#333",
                font=ctk.CTkFont(size=10), text_color=self._C_TEXT,
            ).pack(side=tk.LEFT)

        _divider()

        # ── Controls row ──────────────────────────────────────────────────────
        ctrl = ctk.CTkFrame(root, fg_color="transparent")
        ctrl.pack(fill=tk.X, padx=P, pady=6)

        def _ghost_btn(parent, text, cmd, **kw):
            kw.setdefault("font", ctk.CTkFont(size=10))
            return ctk.CTkButton(
                parent, text=text, command=cmd, height=34,
                fg_color=self._C_SURFACE2, hover_color="#333",
                text_color=self._C_TEXT, border_width=1, border_color=self._C_BORDER,
                **kw,
            )

        _ghost_btn(ctrl, "Re-Select Feed", self._reselect_feed).pack(
            side=tk.LEFT, expand=True, fill=tk.X, padx=(0, 4))
        _ghost_btn(ctrl, "Re-Select Head", self._reselect_head).pack(
            side=tk.LEFT, expand=True, fill=tk.X, padx=4)
        self._hold_btn = _ghost_btn(ctrl, "Hold Volume", self._toggle_hold,
                                    font=ctk.CTkFont(size=10, weight="bold"), width=120)
        self._hold_btn.pack(side=tk.LEFT, padx=(4, 0))
        self._about_btn = _ghost_btn(ctrl, "ⓘ", self._show_about_menu, width=34)
        self._about_btn.pack(side=tk.LEFT, padx=(4, 0))

        _divider()

        # ── Status labels ─────────────────────────────────────────────────────
        self.info_label = ctk.CTkLabel(
            root, text="State: --  |  Vol: --  |  WS: Disconnected",
            font=ctk.CTkFont(size=15, weight="bold"), text_color=self._C_TEXT,
        )
        self.info_label.pack(pady=(8, 2))

        self.stats_label = ctk.CTkLabel(
            root, text="Session: 00:00  |  Edges: 0",
            font=ctk.CTkFont(size=10), text_color=self._C_TEXT_DIM,
        )
        self.stats_label.pack(pady=(0, 10))

    def _start_update_check(self):
        def callback(latest, url):
            self.root.after(0, self._show_update_banner, latest, url)
        threading.Thread(target=check_for_update, args=(callback,), daemon=True).start()

    # ------------------------------------------------------------------ config persistence

    def _save_config(self):
        try:
            CONFIG_PATH.parent.mkdir(parents=True, exist_ok=True)
            data = {
                "heights":     self.heights,
                "min_vol":     self.min_vol_var.get(),
                "max_vol":     self.max_vol_var.get(),
                "aggressiveness": self.aggr_var.get(),
                "mode":        self.mode_var.get(),
                "port":        self.port_var.get(),
                "device_name": self._device_combo.get(),
                "mp3_path":    getattr(self, '_mp3_last_path', ""),
                "mp3_path_type": getattr(self, '_mp3_last_type', "file"),
                "mp3_loop":    self._mp3_loop_var.get() if _MINIAUDIO_OK else "folder",
            }
            CONFIG_PATH.write_text(json.dumps(data, indent=2))
        except Exception as e:
            log.warning(f"Config: save failed: {e}")

    def _load_config(self):
        try:
            if not CONFIG_PATH.exists():
                return
            data = json.loads(CONFIG_PATH.read_text())
            # Heights
            for key in ("Edging", "Erect", "Flaccid"):
                if key in data.get("heights", {}):
                    self.heights[key] = data["heights"][key]
            # Sliders / controls
            if "min_vol"        in data: self.min_vol_var.set(data["min_vol"])
            if "max_vol"        in data: self.max_vol_var.set(data["max_vol"])
            if "aggressiveness" in data:
                val = data["aggressiveness"]
                # Handle old int format from previous sessions
                if isinstance(val, int):
                    val = list(AGGR_LEVELS.keys())[min(val, len(AGGR_LEVELS) - 1)]
                self.aggr_var.set(val)
            if "mode" in data:
                self.mode_var.set(data["mode"])
                self._on_mode_change()
            if "port"        in data: self.port_var.set(data["port"])
            if "device_name" in data: self._device_combo.set(data["device_name"])
            if _MINIAUDIO_OK and self.music_player:
                if "mp3_loop" in data:
                    self._mp3_loop_var.set(data["mp3_loop"])
                    self.music_player.loop_mode = data["mp3_loop"]
                mp3_path = data.get("mp3_path", "")
                mp3_type = data.get("mp3_path_type", "file")
                if mp3_path and pathlib.Path(mp3_path).exists():
                    self._mp3_last_path = mp3_path
                    self._mp3_last_type = mp3_type
                    # Don't auto-play on load — just prime the label
                    if mp3_type == "folder":
                        files = sorted(
                            [f for f in pathlib.Path(mp3_path).iterdir()
                             if f.is_file() and f.suffix.lower() in MusicPlayer.EXTS],
                            key=lambda f: f.name.lower(),
                        )
                        if files:
                            self.music_player._playlist = files
                            self.music_player._idx = 0
                            self.music_player.track_name = files[0].stem
                            self.music_player.track_info = f"1 / {len(files)}"
                    else:
                        p = pathlib.Path(mp3_path)
                        self.music_player._playlist = [p]
                        self.music_player._idx = 0
                        self.music_player.track_name = p.stem
                        self.music_player.track_info = "1 / 1"
                    self._mp3_update_track_label()
            log.info(f"Config: loaded from {CONFIG_PATH}")
        except Exception as e:
            log.warning(f"Config: load failed: {e}")

    def _cleanup(self):
        """Non-UI cleanup — safe to call from atexit or _on_close."""
        if getattr(self, '_cleaned_up', False):
            return
        self._cleaned_up = True
        self._running = False

        # Wait for capture thread to finish (max 1.5 s)
        t = getattr(self, '_capture_thread', None)
        if t and t.is_alive():
            t.join(timeout=1.5)

        # Zero out Restim before closing socket
        try:
            self.restim.set_volume(0.0, instant=True)
        except Exception:
            pass

        # Close WebSocket
        with self.restim._lock:
            old_ws = self.restim.ws
            self.restim.ws = None
        if old_ws:
            try:
                old_ws.close()
            except Exception:
                pass

        # Stop MP3 player
        if self.music_player:
            try:
                self.music_player.cleanup()
            except Exception:
                pass

        # Restore Windows audio to pre-session level
        if self.win_audio and self.win_audio.connected and self._orig_win_volume is not None:
            try:
                self.win_audio._volume_interface.SetMasterVolumeLevelScalar(
                    self._orig_win_volume, None)
                log.info(f"WinAudio: restored volume to {self._orig_win_volume:.2f}")
            except Exception:
                pass

    def _on_close(self):
        self._save_config()
        self._cleanup()
        self.root.destroy()

    # ------------------------------------------------------------------ capture thread

    def _capture_loop(self):
        """Runs in background thread — captures frames and drops them in the queue."""
        while self._running:
            if self.tracking_paused:
                time.sleep(0.05)
                continue
            frame = capture_window_region(self.hwnd, self.rel_box)
            if frame is not None:
                # If the queue is full, discard the stale frame so we always
                # feed the tracker the freshest capture available
                if self._frame_queue.full():
                    try:
                        self._frame_queue.get_nowait()
                    except queue.Empty:
                        pass
                self._frame_queue.put(frame)
            else:
                time.sleep(0.05)

    # ------------------------------------------------------------------ callbacks

    def _show_about_menu(self):
        menu = tk.Menu(self.root, tearoff=0,
                       bg=self._C_SURFACE2, fg=self._C_TEXT,
                       activebackground=self._C_RED, activeforeground="white",
                       borderwidth=1, relief="flat",
                       font=("Segoe UI", 10))
        menu.add_command(label=f"Version  v{VERSION}", state="disabled")
        menu.add_separator()
        menu.add_command(label="Support Us ❤   ko-fi.com/stimstation",
                         command=lambda: webbrowser.open("https://ko-fi.com/stimstation"))
        btn = self._about_btn
        x = btn.winfo_rootx()
        y = btn.winfo_rooty() + btn.winfo_height()
        try:
            menu.tk_popup(x, y)
        finally:
            menu.grab_release()

    def _show_update_banner(self, latest, url):
        self._update_label.configure(text=f"Update available: v{latest}")
        self._update_btn.configure(command=lambda: webbrowser.open(url))
        self._update_banner.pack(fill=tk.X, before=self._first_widget)


    _AUTO_SETTLE   = 5.0   # seconds of observation before first height apply
    _AUTO_INTERVAL = 2.0   # seconds between re-applies

    def _toggle_auto(self):
        self._auto_mode = not self._auto_mode
        if self._auto_mode:
            self._auto_min_y = self._auto_max_y = None
            self._auto_obs_start = None
            self._auto_last_apply = 0.0
            self._auto_btn.configure(text="AUTO  (observing...)",
                                     fg_color="#b8a000", hover_color="#8a7800")
        else:
            self._auto_btn.configure(text="AUTO  (off)",
                                     fg_color=self._C_SURFACE2, hover_color="#333")

    def _auto_feed(self, y: float):
        """Called every heavy frame while AUTO is on and tracking is good."""
        now = time.time()
        if self._auto_obs_start is None:
            self._auto_obs_start = now

        # Expand observed range
        if self._auto_min_y is None or y < self._auto_min_y:
            self._auto_min_y = y
        if self._auto_max_y is None or y > self._auto_max_y:
            self._auto_max_y = y

        elapsed = now - self._auto_obs_start

        # Update button label while settling — only once per second to avoid UI stall
        if elapsed < self._AUTO_SETTLE:
            remaining = int(self._AUTO_SETTLE - elapsed) + 1
            if remaining != getattr(self, '_auto_last_remaining', -1):
                self._auto_last_remaining = remaining
                self._auto_btn.configure(text=f"AUTO  (observing {remaining}s...)")
            return

        # Throttle actual height application
        if now - self._auto_last_apply < self._AUTO_INTERVAL:
            return
        self._auto_last_apply = now

        rng = self._auto_max_y - self._auto_min_y
        if rng < 8:   # not enough vertical range yet — keep observing
            return

        self.heights["Edging"]  = self._auto_min_y + 0.05 * rng
        self.heights["Erect"]   = (self._auto_min_y + self._auto_max_y) / 2
        self.heights["Flaccid"] = self._auto_max_y

        if getattr(self, '_auto_btn_state', None) != "active":
            self._auto_btn_state = "active"
            self._auto_btn.configure(text="AUTO  \u2713", fg_color="#1a8a1a", hover_color="#145c14")
        log.debug(f"AUTO heights: edging={self.heights['Edging']:.0f} "
                  f"erect={self.heights['Erect']:.0f} flaccid={self.heights['Flaccid']:.0f}")

    def _set_edging(self):
        self.heights["Edging"] = self.head_y
        log.info(f"Edging height set at Y={self.head_y}")
        self._save_config()

    def _set_erect(self):
        self.heights["Erect"] = self.head_y
        log.info(f"Erect height set at Y={self.head_y}")
        self._save_config()

    def _set_flaccid(self):
        self.heights["Flaccid"] = self.head_y
        log.info(f"Flaccid height set at Y={self.head_y}")
        self._save_config()

    def _reset_heights(self):
        """Clear calibrated heights — called whenever the feed changes."""
        self.heights = {"Edging": None, "Erect": None, "Flaccid": None, "Ruin": None}
        self._auto_min_y = None
        self._auto_max_y = None
        self._auto_frame_count = 0
        self._auto_start_time = None
        self._auto_last_shown_sec = -1
        if hasattr(self, '_auto_btn_state'):
            self._auto_btn_state = None
        if hasattr(self, '_auto_btn'):
            self._auto_btn.configure(text="AUTO", fg_color=self._C_SURFACE2,
                                     hover_color="#333")
        log.info("Heights reset after feed re-selection")

    def _reselect_feed(self):
        self.tracking_paused = True
        new_hwnd, new_rel_box = select_region(self.root)
        if new_hwnd and new_rel_box['width'] > 10 and new_rel_box['height'] > 10:
            self.hwnd    = new_hwnd
            self.rel_box = new_rel_box
            self._reset_heights()
            self._reselect_head()
        else:
            self.tracking_paused = False

    def _reselect_head(self):
        self.tracking_paused = True
        try:
            pause_frame = capture_window_region(self.hwnd, self.rel_box)
            if pause_frame is None:
                messagebox.showwarning(
                    "Capture Failed",
                    "Could not grab a frame from the video feed.\n\n"
                    "Make sure the feed window is not minimised, then try again.",
                    parent=self.root,
                )
                return
            new_bbox = select_head(pause_frame, parent=self.root)
            if new_bbox[2] > 0 and new_bbox[3] > 0:
                self.tracker.init(pause_frame, new_bbox)
                self.last_bbox   = new_bbox
                self.head_y      = new_bbox[1] + new_bbox[3] // 2
                self.tracking_ok = True
        finally:
            self.tracking_paused = False

    _AGGR_COLORS = {
        "Easy":   ("#1a8a1a", "#145c14"),
        "Middle": ("#b8a000", "#8a7800"),
        "Hard":   ("#cc6600", "#994c00"),
        "Expert": ("#cc2200", "#991800"),
    }

    def _on_aggr_change(self, val=None):
        val = val or self.aggr_var.get()
        col, hov = self._AGGR_COLORS.get(val, (self._C_RED, self._C_RED_HOV))
        self._aggr_seg.configure(selected_color=col, selected_hover_color=hov)
        self._save_config()

    def _on_cum(self):
        """Drop volume to 0, hold 5 min, ramp back to floor over 5 min."""
        self._cum_time = time.time()
        floor_val = min(self.min_vol_var.get(), self.max_vol_var.get()) / 100.0
        ceil_val  = self.max_vol_var.get() / 100.0
        mode = self.mode_var.get()
        if mode == "restim":
            self.restim.set_volume(0.0, instant=True, floor=0.0, ceiling=ceil_val)
        elif self.win_audio and self.win_audio.connected:
            self.win_audio.set_volume(0.0, 0.0, 1.0)
        log.info("Cum triggered — volume silenced for 5 min")

    def _on_floor_change(self, val):
        self._floor_lbl.configure(text=f"{float(val):.0f}%")
        self._save_config()

    def _on_ceil_change(self, val):
        self._ceil_lbl.configure(text=f"{float(val):.0f}%")
        self._save_config()

    def _on_mode_change(self):
        mode = self.mode_var.get()
        P    = 12
        # Hide all panels first
        self._restim_opts.pack_forget()
        self._windows_opts.pack_forget()
        self._mp3_opts.pack_forget()
        # Stop music if leaving mp3 mode
        if mode != "mp3" and self.music_player:
            self.music_player.stop()
        if mode == "restim":
            self._restim_opts.pack(fill=tk.X, padx=P, pady=(0, 4),
                                   after=self._mode_card)
        elif mode == "windows":
            self._windows_opts.pack(fill=tk.X, padx=P, pady=(0, 4),
                                    after=self._mode_card)
            if not self.win_devices:
                self._refresh_devices()
        elif mode == "mp3" and _MINIAUDIO_OK:
            self._mp3_opts.pack(fill=tk.X, padx=P, pady=(0, 4),
                                after=self._mode_card)
        self._save_config()

    def _on_port_change(self, *_):
        val = self.port_var.get().strip()
        if val.isdigit():
            new_port = int(val)
            if not (1 <= new_port <= 65535):
                return
            if new_port != self.restim.port:
                self.restim.port = new_port
                with self.restim._lock:
                    old_ws = self.restim.ws
                    self.restim.ws = None
                if old_ws:
                    try:
                        old_ws.close()
                    except Exception:
                        pass

    def _refresh_devices(self):
        try:
            self.win_devices = list_audio_devices()
        except Exception as e:
            log.error(f"Audio device scan failed: {e}")
            messagebox.showerror(
                "Audio Device Error",
                f"Could not scan audio devices:\n{e}\n\n"
                "Try running as Administrator if this keeps happening."
            )
            self.win_devices = []
        names = [d.FriendlyName for d in self.win_devices]
        self._device_combo.configure(values=names)
        if names:
            self._device_combo.set(names[0])
            self._on_device_select(names[0])
        else:
            log.warning("No audio output devices found")

    def _on_device_select(self, value):
        for d in self.win_devices:
            if d.FriendlyName == value:
                self.win_audio = WindowsAudioClient(d)
                if self.win_audio.connected:
                    self._orig_win_volume = self.win_audio.get_volume()
                else:
                    self._orig_win_volume = None
                self._save_config()
                break

    def _toggle_hold(self):
        self.hold_active = not self.hold_active
        if self.hold_active:
            self._hold_btn.configure(text="HELD [|]",
                                     fg_color=self._C_RED, hover_color=self._C_RED_HOV,
                                     border_color=self._C_RED)
        else:
            self._hold_btn.configure(text="Hold Volume",
                                     fg_color=self._C_SURFACE2, hover_color="#333",
                                     border_color=self._C_BORDER)

    # ------------------------------------------------------------------ MP3 transport callbacks

    def _mp3_load_file(self):
        from tkinter import filedialog
        path = filedialog.askopenfilename(
            title="Select audio file",
            filetypes=[
                ("Audio files", "*.mp3 *.wav *.ogg *.flac *.m4a"),
                ("All files", "*.*"),
            ],
        )
        if path and self.music_player:
            self.music_player.load_file(path)
            self._mp3_last_path = path
            self._mp3_last_type = "file"
            self._mp3_loop_var.set(self.music_player.loop_mode)
            self._mp3_update_track_label()

    def _mp3_load_folder(self):
        from tkinter import filedialog
        folder = filedialog.askdirectory(title="Select music folder")
        if folder and self.music_player:
            self.music_player.load_folder(folder)
            self._mp3_last_path = folder
            self._mp3_last_type = "folder"
            self._mp3_loop_var.set(self.music_player.loop_mode)
            self._mp3_update_track_label()

    def _mp3_play_pause(self):
        if not self.music_player:
            return
        if self.music_player._state == "playing":
            self.music_player.pause()
            self._mp3_play_btn.configure(text="▶")
        else:
            self.music_player.play()
            self._mp3_play_btn.configure(text="⏸")

    def _mp3_stop(self):
        if self.music_player:
            self.music_player.stop()
            self._mp3_play_btn.configure(text="▶")

    def _mp3_prev(self):
        if self.music_player:
            self.music_player.prev_track()

    def _mp3_next(self):
        if self.music_player:
            self.music_player.next_track()

    def _mp3_on_loop_change(self, val):
        if self.music_player:
            self.music_player.loop_mode = val

    def _mp3_update_track_label(self):
        if not self.music_player:
            return
        name = self.music_player.track_name
        info = self.music_player.track_info
        if name:
            text = f"{name}  [{info}]" if info else name
        else:
            text = "No file loaded"
        self._mp3_track_lbl.configure(text=text)
        # Keep play button in sync
        if self.music_player._state == "playing":
            self._mp3_play_btn.configure(text="⏸")
        else:
            self._mp3_play_btn.configure(text="▶")

    # ------------------------------------------------------------------ frame loop

    # Throttle constants (ms)
    _DISPLAY_INTERVAL_MS  = 50   # max 20 fps for the preview panel
    _STATUS_INTERVAL_MS   = 250  # max 4 fps for the status label text

    def _update_frame(self):
        # Always reschedule — even if an exception occurs mid-frame the loop
        # must not die, otherwise the UI freezes permanently.
        try:
            if self.tracking_paused:
                return

            try:
                frame = self._frame_queue.get_nowait()
            except queue.Empty:
                return

            now = time.time()
            self._frame_times.append(now)

            # Tracking + volume: every other frame in non-Expert modes
            self._proc_frame_count = getattr(self, '_proc_frame_count', 0) + 1
            heavy = (self.aggr_var.get() == "Expert" or self._proc_frame_count % 2 == 1)

            if heavy:
                self._maybe_yolo_reanchor(frame)
                self._run_tracker(frame)
                if self._auto_mode and self.tracking_ok:
                    self._auto_feed(self.head_y)

            state = self._determine_state(self.head_y)

            if heavy:
                self._update_stats(state)
                self._tick_volume()

            # Status label — throttled to _STATUS_INTERVAL_MS
            _last_status = getattr(self, '_last_status_time', 0.0)
            if now - _last_status >= self._STATUS_INTERVAL_MS / 1000:
                self._update_status_label(state)
                self._last_status_time = now

            # Video display — throttled to _DISPLAY_INTERVAL_MS
            _last_disp = getattr(self, '_last_display_time', 0.0)
            if now - _last_disp >= self._DISPLAY_INTERVAL_MS / 1000:
                self._draw_height_lines(frame)
                self._draw_tracking_overlay(frame)
                self._display_frame(frame)
                self._last_display_time = now

        except Exception:
            log.exception("_update_frame: unhandled exception — loop continues")
        finally:
            interval = 100 if self.tracking_paused else 16
            self.root.after(interval, self._update_frame)

    def _maybe_yolo_reanchor(self, frame):
        self.yolo_frame_counter += 1
        if not self.detector.available or self.yolo_frame_counter < self._YOLO_INTERVAL:
            return
        self.yolo_frame_counter = 0

        yolo_bbox = self.detector.detect_head(frame)
        if yolo_bbox is None:
            self.yolo_candidate = None
            return

        px, py, pw, ph = self.last_bbox
        cur_cx, cur_cy = px + pw // 2, py + ph // 2
        yx, yy, yw, yh = yolo_bbox
        det_cx, det_cy = yx + yw // 2, yy + yh // 2
        diag = np.sqrt(pw ** 2 + ph ** 2)

        if np.sqrt((det_cx - cur_cx) ** 2 + (det_cy - cur_cy) ** 2) > diag * self._YOLO_MAX_JUMP:
            self.yolo_candidate = None
            return

        if self.yolo_candidate is not None:
            prev_bbox, hits = self.yolo_candidate
            pcx = prev_bbox[0] + prev_bbox[2] // 2
            pcy = prev_bbox[1] + prev_bbox[3] // 2
            hits = hits + 1 if np.sqrt((det_cx - pcx) ** 2 + (det_cy - pcy) ** 2) <= diag else 1
        else:
            hits = 1

        self.yolo_candidate = (yolo_bbox, hits)

        if hits >= self._YOLO_CONFIRM:
            self.tracker.init(frame, yolo_bbox)
            self.last_bbox      = yolo_bbox
            self.yolo_candidate = None
            x, y, w, h = yolo_bbox
            cv2.rectangle(frame, (x, y), (x + w, y + h), (255, 255, 0), 1)

    def _run_tracker(self, frame):
        """Update tracker state — does NOT draw; call _draw_tracking_overlay every frame."""
        success, new_bbox = self.tracker.update(frame)
        if success:
            x, y, w, h_box = [int(v) for v in new_bbox]
            px, py, pw, ph = self.last_bbox
            prev_cx, prev_cy = px + pw // 2, py + ph // 2
            new_cx,  new_cy  = x  + w  // 2, y  + h_box // 2
            diag = np.sqrt(pw ** 2 + ph ** 2)

            size_ok = (
                0 < w <= frame.shape[1] and 0 < h_box <= frame.shape[0] and
                (1 / self._SIZE_RATIO_MAX) < (w / max(pw, 1)) < self._SIZE_RATIO_MAX and
                (1 / self._SIZE_RATIO_MAX) < (h_box / max(ph, 1)) < self._SIZE_RATIO_MAX
            )
            jump_ok = np.sqrt((new_cx - prev_cx) ** 2 + (new_cy - prev_cy) ** 2) < diag * self._MAX_JUMP_FACTOR

            if size_ok and jump_ok:
                self.last_bbox    = (x, y, w, h_box)
                self.tracking_ok  = True
                self._track_msg   = ""
                self.head_y       = new_cy
                self._head_y_history.append(new_cy)
            else:
                reason = "size" if not size_ok else "jump"
                self._track_msg  = f"TRACKING SUSPECT ({reason}) - Frozen"
                self.tracking_ok = False
                self.tracker.init(frame, self.last_bbox)
        else:
            self._track_msg  = "TRACKING LOST - Frozen at last position"
            self.tracking_ok = False
            self.tracker.init(frame, self.last_bbox)

    def _draw_tracking_overlay(self, frame):
        """Draw the current tracking bbox every frame so there is no flicker
        when _run_tracker is skipped on non-heavy frames."""
        px, py, pw, ph = self.last_bbox
        if self.tracking_ok:
            cv2.rectangle(frame, (px, py), (px + pw, py + ph), (0, 255, 0), 2)
            cx, cy = px + pw // 2, py + ph // 2
            cv2.circle(frame, (cx, cy), 4, (0, 0, 255), -1)
        else:
            msg = getattr(self, '_track_msg', '')
            colour = (0, 165, 255) if "SUSPECT" in msg else (0, 0, 255)
            if msg:
                cv2.putText(frame, msg, (10, 30),
                            cv2.FONT_HERSHEY_SIMPLEX, 0.6, colour, 2)
            cv2.rectangle(frame, (px, py), (px + pw, py + ph), colour, 2)

    def _determine_state(self, y_pos):
        if any(self.heights.get(k) is None for k in ("Edging", "Erect", "Flaccid")):
            return "Erect (Needs Calibration)"
        dist_edging  = abs(y_pos - self.heights["Edging"])
        dist_erect   = abs(y_pos - self.heights["Erect"])
        dist_flaccid = abs(y_pos - self.heights["Flaccid"])
        minimum = min(dist_edging, dist_erect, dist_flaccid)
        if minimum == dist_edging:  return "Edging"
        if minimum == dist_flaccid: return "Flaccid"
        return "Erect"

    def _update_stats(self, state):
        """Accumulate time-in-state and count edge events. Throttled label refresh."""
        now     = time.time()
        elapsed = now - self._last_state_time
        self._last_state_time = now

        if self._prev_state in self.state_times:
            self.state_times[self._prev_state] += elapsed

        if state == "Edging" and self._prev_state != "Edging":
            self.edge_count += 1

        self._prev_state = state

        # Refresh the label every 30 frames (~1 s at 30 fps) to avoid thrashing
        self._stats_tick += 1
        if self._stats_tick < 30:
            return
        self._stats_tick = 0

        total_session = now - self.session_start
        m, s = divmod(int(total_session), 60)

        total_state = sum(self.state_times.values())
        if total_state > 0:
            pcts = {k: v / total_state * 100 for k, v in self.state_times.items()}
        else:
            pcts = {"Edging": 0.0, "Erect": 0.0, "Flaccid": 0.0}

        self.stats_label.configure(
            text=(f"Session: {m:02d}:{s:02d}  |  Edges: {self.edge_count}  |  "
                  f"Edging {pcts['Edging']:.0f}%  "
                  f"Erect {pcts['Erect']:.0f}%  "
                  f"Flaccid {pcts['Flaccid']:.0f}%")
        )

    def _draw_height_lines(self, frame):
        fw = frame.shape[1]
        if self.heights["Edging"]  is not None: cv2.line(frame, (0, int(self.heights["Edging"])),  (fw, int(self.heights["Edging"])),  (0, 0, 255), 2)
        if self.heights["Erect"]   is not None: cv2.line(frame, (0, int(self.heights["Erect"])),   (fw, int(self.heights["Erect"])),   (0, 255, 0), 2)
        if self.heights["Flaccid"] is not None: cv2.line(frame, (0, int(self.heights["Flaccid"])), (fw, int(self.heights["Flaccid"])), (255, 0, 0), 2)

    def _compute_volume_delta(self):
        if any(self.heights.get(k) is None for k in ("Edging", "Erect", "Flaccid")):
            return 0.0
        full_range = self.heights["Flaccid"] - self.heights["Edging"]
        if abs(full_range) < 1:
            return 0.0

        history = self._head_y_history
        smoothed_y = sum(history) / len(history) if history else self.head_y

        position   = (smoothed_y             - self.heights["Edging"]) / full_range
        erect_norm = (self.heights["Erect"]   - self.heights["Edging"]) / full_range
        aggr_mult = AGGR_LEVELS.get(self.aggr_var.get(), 1.0)

        # Velocity: normalised rate of change across the history window.
        # Positive = moving toward flaccid, negative = moving toward edging.
        if len(history) >= 4:
            velocity = (history[-1] - history[0]) / (len(history) * abs(full_range))
        else:
            velocity = 0.0

        if 0.0 <= position <= erect_norm:
            # Inside the sweet zone — only react if clearly trending flaccid
            vel_nudge = max(0.0, velocity) * 0.4
            return VOLUME_STEP * vel_nudge * aggr_mult

        if position < 0.0:
            # Past edging — ease off; dampen further if still moving toward edging
            vel_damp = max(0.0, -velocity) * 0.3
            return -VOLUME_STEP * (0.5 + vel_damp) * aggr_mult

        # Past erect, drifting toward flaccid
        dist      = (position - erect_norm) / max(1.0 - erect_norm, 0.01)
        vel_boost = max(0.0, velocity) * 0.5   # moving fast toward flaccid = respond harder
        return VOLUME_STEP * (min(dist, 1.0) + vel_boost) * aggr_mult


    _CUM_SILENCE = 300.0   # seconds at zero
    _CUM_RAMP    = 300.0   # seconds ramping back to floor

    def _tick_volume(self):
        cur_time = time.time()
        if cur_time - self.last_vol_time < VOLUME_UPDATE_INTERVAL:
            return
        self.last_vol_time = cur_time

        # ── Cum cooldown overrides everything ─────────────────────────────────
        if self._cum_time is not None:
            elapsed   = cur_time - self._cum_time
            floor_val = min(self.min_vol_var.get(), self.max_vol_var.get()) / 100.0
            ceil_val  = self.max_vol_var.get() / 100.0
            if elapsed < self._CUM_SILENCE:
                target = 0.0
            elif elapsed < self._CUM_SILENCE + self._CUM_RAMP:
                progress = (elapsed - self._CUM_SILENCE) / self._CUM_RAMP
                target   = floor_val * progress
            else:
                self._cum_time = None   # cooldown over — fall through to normal
                target = None
            if target is not None:
                mode = self.mode_var.get()
                if mode == "restim":
                    self.restim.set_volume(target, floor=0.0, ceiling=ceil_val)
                elif mode == "windows" and self.win_audio and self.win_audio.connected:
                    self.win_audio.set_volume(target, 0.0, 1.0)
                elif mode == "mp3" and self.music_player:
                    self.music_player.volume = target
                return

        if self.hold_active:
            return  # Volume frozen by user

        history    = self._head_y_history
        smoothed_y = sum(history) / len(history) if history else self.head_y


        floor_val = min(self.min_vol_var.get(), self.max_vol_var.get()) / 100.0
        ceil_val  = self.max_vol_var.get() / 100.0
        delta     = self._compute_volume_delta()
        mode      = self.mode_var.get()

        if mode == "restim":
            if delta != 0.0:
                self.restim.adjust_volume(delta, floor=floor_val, ceiling=ceil_val)
            self.restim.maybe_reconnect()
        elif mode == "windows" and self.win_audio and self.win_audio.connected and delta != 0.0:
            self.win_audio.adjust_volume(delta, floor=floor_val, ceiling=ceil_val)
        elif mode == "mp3" and self.music_player and delta != 0.0:
            self.music_player.adjust_volume(delta, floor=floor_val, ceiling=ceil_val)

    def _update_status_label(self, state):
        quality_str = "Track: OK" if self.tracking_ok else "Track: LOST"
        mode = self.mode_var.get()
        if mode == "restim":
            conn_color = "#00ff00" if self.restim.ws else "#ff0000"
            vol_str    = f"{self.restim.volume * 100:.0f}%"
            src_str    = f"WS: {'Connected' if self.restim.ws else 'Disconnected'}"
        elif mode == "windows" and self.win_audio and self.win_audio.connected:
            conn_color = "#00ff00"
            vol_str    = f"{self.win_audio.get_volume() * 100:.0f}%"
            src_str    = "Win Audio: OK"
        elif mode == "mp3" and self.music_player:
            self._mp3_update_track_label()
            conn_color = "#00ff00" if self.music_player._state == "playing" else "#ffaa00"
            vol_str    = f"{self.music_player.volume * 100:.0f}%"
            src_str    = f"MP3: {self.music_player.track_name or '—'}"
        else:
            conn_color = "#ffaa00"
            vol_str    = "--"
            src_str    = "Win Audio: No Device"

        ft = self._frame_times
        fps = (len(ft) - 1) / max(ft[-1] - ft[0], 1e-9) if len(ft) >= 2 else 0.0
        yolo_str = (f"YOLO: {self.detector.last_conf:.0%}"
                    if self.detector.available and self.detector.last_conf > 0 else "YOLO: --")

        self.info_label.configure(
            text=(f"State: {state}  |  Vol: {vol_str}  |  {quality_str}"
                  f"  |  {src_str}  |  {fps:.0f} fps  |  {yolo_str}"),
            text_color=conn_color,
        )

    def _display_frame(self, frame):
        img   = Image.fromarray(cv2.cvtColor(frame, cv2.COLOR_BGR2RGB))
        imgtk = ImageTk.PhotoImage(image=img)
        self.video_label.imgtk = imgtk  # prevent GC
        self.video_label.configure(image=imgtk)


def show_splash() -> bool:
    """Welcome / setup checklist screen. Returns True if user clicked Start.

    If splash.png exists, displays it as a background image with the version
    overlaid dynamically (so the user can Photoshop the design and we always
    stamp the current version at runtime).  Falls back to plain tkinter widgets
    when the PNG is absent.
    """
    import tkinter as tk
    from tkinter import font as tkfont

    root = tk.Tk()
    root.title("VisualStimEdger")
    root.configure(bg="#0d0d0d")
    root.resizable(False, False)
    _icon = pathlib.Path(resource_path("icon.ico"))
    if _icon.exists():
        try:
            root.iconbitmap(str(_icon))
        except Exception:
            pass

    started = False

    def _start():
        nonlocal started
        started = True
        root.destroy()

    # ── PNG path (bundled resource) ────────────────────────────────────────────
    splash_path = pathlib.Path(resource_path("splash.png"))

    if splash_path.exists():
        try:
            from PIL import Image, ImageDraw, ImageFont, ImageTk

            img = Image.open(splash_path).convert("RGBA")
            W, H = img.size

            # Stamp version text onto the image at the space the designer left
            draw = ImageDraw.Draw(img)
            try:
                f_ver = ImageFont.truetype("C:/Windows/Fonts/segoeui.ttf", 11)
            except Exception:
                f_ver = ImageFont.load_default()
            ver_txt = f"v{VERSION}  ·  edge smarter"
            bb = draw.textbbox((0, 0), ver_txt, font=f_ver)
            tx = (W - (bb[2] - bb[0])) // 2
            draw.text((tx, 170), ver_txt, fill=(100, 100, 100, 255), font=f_ver)

            # Flatten RGBA onto the dark background colour (#0d0d0d = 13,13,13)
            bg_flat = Image.new("RGB", (W, H), (13, 13, 13))
            bg_flat.paste(img, mask=img.split()[3])

            photo = ImageTk.PhotoImage(bg_flat)

            lbl = tk.Label(root, image=photo, bg="#0d0d0d", cursor="hand2",
                           borderwidth=0, highlightthickness=0)
            lbl.image = photo  # keep reference
            lbl.pack()

            # Transparent overlay button that sits on top of the baked-in
            # red Start button (coordinates match make_splash.py BX/BY values
            # scaled to actual image size; default canvas is 700×640).
            scale = W / 700
            bx1 = int(88  * scale)
            by1 = int(540 * scale)
            bx2 = int(612 * scale)
            by2 = int(596 * scale)
            overlay = tk.Button(
                root, text="", bg="#cc2200", activebackground="#991800",
                relief="flat", bd=0, cursor="hand2", command=_start,
            )
            overlay.place(x=bx1, y=by1, width=bx2 - bx1, height=by2 - by1)

            root.protocol("WM_DELETE_WINDOW", root.destroy)
            sw = root.winfo_screenwidth()
            sh = root.winfo_screenheight()
            root.geometry(f"{W}x{H}+{(sw - W) // 2}+{(sh - H) // 2}")
            root.attributes("-topmost", True)
            root.after(400, lambda: root.attributes("-topmost", False))
            root.mainloop()
            return started
        except Exception as e:
            log.warning(f"splash.png display failed ({e}) — falling back to widget splash")

    # ── Widget fallback ────────────────────────────────────────────────────────
    BG       = "#0d0d0d"
    CARD     = "#1a1a1a"
    RED      = "#cc2200"
    YELLOW   = "#ffcc00"
    TEXT     = "#eeeeee"
    DIM      = "#666666"
    SUBDIM   = "#888888"
    P        = 24

    f_title  = tkfont.Font(family="Segoe UI", size=20, weight="bold")
    f_sub    = tkfont.Font(family="Segoe UI", size=10)
    f_head   = tkfont.Font(family="Segoe UI", size=12, weight="bold")
    f_body   = tkfont.Font(family="Segoe UI", size=10)
    f_num    = tkfont.Font(family="Segoe UI", size=10, weight="bold")
    f_btn    = tkfont.Font(family="Segoe UI", size=13, weight="bold")

    root.resizable(False, True)

    tk.Label(root, text="VisualStimEdger", font=f_title,
             bg=BG, fg=TEXT).pack(pady=(P, 2))
    tk.Label(root, text=f"v{VERSION}  ·  edge smarter", font=f_sub,
             bg=BG, fg=DIM).pack(pady=(0, P//2))

    card = tk.Frame(root, bg=CARD, padx=16, pady=12)
    card.pack(fill="x", padx=P, pady=(0, 10))

    tk.Label(card, text="Before you hit Start", font=f_head,
             bg=CARD, fg=YELLOW, anchor="w").pack(anchor="w", pady=(0, 8))

    steps = [
        ("1", "Open your camera feed",
              "OBS, webcam app, browser stream — anything showing your cock in a window.\n"
              "It does NOT need to be full-screen."),
        ("2", "Keep it visible and not minimised",
              "The app will ask you to draw a box around that window.\n"
              "Bring it to the front before hitting Start."),
        ("3", "You'll mark the tip of your cock",
              "Draw a small box around the head. An electrode or ring tracks\n"
              "even better than skin alone."),
        ("4", "Calibrate Flaccid / Erect / Edging heights",
              "Do this during your session. AUTO mode can handle it for you.\n"
              "Re-calibrate any time without restarting."),
    ]

    for num, title, body in steps:
        row = tk.Frame(card, bg=CARD)
        row.pack(fill="x", pady=3)
        tk.Label(row, text=num, font=f_num, bg=RED, fg="white",
                 width=2, relief="flat").pack(side="left", anchor="n", padx=(0, 8))
        col = tk.Frame(row, bg=CARD)
        col.pack(side="left", fill="x", expand=True)
        tk.Label(col, text=title, font=f_head, bg=CARD,
                 fg=TEXT, anchor="w", wraplength=440).pack(anchor="w")
        tk.Label(col, text=body, font=f_body, bg=CARD,
                 fg=SUBDIM, anchor="w", justify="left", wraplength=440).pack(anchor="w")

    tk.Label(root,
             text="Controls volume only — does not generate e-stim signals.\n"
                  "You need Restim, xToys, electron-redrive, an .mp3, etc. already running.",
             font=f_body, bg=BG, fg=DIM, justify="center", wraplength=500).pack(pady=(4, 12))

    btn = tk.Button(root, text="I'm ready — select my camera feed  \u2192",
                    font=f_btn, bg=RED, fg="white", activebackground="#991800",
                    activeforeground="white", relief="flat", bd=0,
                    cursor="hand2", command=_start, pady=12)
    btn.pack(fill="x", padx=P, pady=(0, P))

    root.protocol("WM_DELETE_WINDOW", root.destroy)
    W, H = 580, 640
    sw = root.winfo_screenwidth()
    sh = root.winfo_screenheight()
    root.geometry(f"{W}x{H}+{(sw - W) // 2}+{(sh - H) // 2}")
    root.attributes("-topmost", True)
    root.after(400, lambda: root.attributes("-topmost", False))
    root.mainloop()
    return started


def main():
    parser = argparse.ArgumentParser(description="VisualStimEdger")
    parser.add_argument("--debug", action="store_true",
                        help="Enable verbose debug logging to console")
    args = parser.parse_args()

    logging.basicConfig(
        level=logging.DEBUG if args.debug else logging.INFO,
        format="%(asctime)s %(levelname)-8s %(name)s: %(message)s",
        datefmt="%H:%M:%S",
    )

    log.info(f"VisualStimEdger v{VERSION} starting")

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

    bbox = select_head(initial_frame)
    if bbox[2] == 0 or bbox[3] == 0:
        log.info("No head selected — exiting")
        return

    app = App(hwnd, rel_box, initial_frame, bbox)
    atexit.register(app._cleanup)   # catches crashes / force-kills
    app.run()


if __name__ == "__main__":
    main()
