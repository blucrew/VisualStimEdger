import os
import sys
import tempfile
import ctypes

# ── DPI awareness (must be set before any GUI / coordinate work) ──────────
# Without this, Windows virtualises coordinates on multi-monitor setups with
# scaling enabled, causing mss, win32gui and tkinter to disagree on positions.
try:
    ctypes.windll.shcore.SetProcessDpiAwareness(2)   # PROCESS_PER_MONITOR_DPI_AWARE
except Exception:
    try:
        ctypes.windll.user32.SetProcessDPIAware()     # fallback for older Windows
    except Exception:
        pass

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
import ssl
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
VERSION = "1.7.7"
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

# Edge-count threshold before pre-emptive sweet-zone denial unlocks.
# None = never unlocks (pure reactive mechanic). Lower = meaner sooner.
# After unlocking, strength escalates slightly per additional edge (capped),
# so the tool tightens progressively across a long session.
PREEMPT_UNLOCK = {
    "Easy":   None,   # disabled — reactive only, gentle strength
    "Middle": None,   # disabled — reactive only, normal strength
    "Hard":   3,      # 3 reactive edges, then predictive, escalating
    "Expert": 0,      # predictive from the first moment
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

        # overrideredirect must be set before geometry so the window manager
        # never adds a title-bar (which would shrink the client area).
        self.root.overrideredirect(True)
        self.root.configure(background='black')
        self.root.attributes("-topmost", True)

        # Set size via geometry, then use MoveWindow to handle negative coords
        # (secondary monitor left of primary gives negative offset_x which
        # tkinter geometry strings like "+-1920+0" cannot express correctly).
        self.root.geometry(f"{w}x{h}+0+0")
        self.root.update_idletasks()
        ctypes.windll.user32.MoveWindow(
            self.root.winfo_id(), self.offset_x, self.offset_y, w, h, True)
        self.root.config(cursor="cross")

        self.canvas = tk.Canvas(self.root, cursor="cross", bg="black",
                                highlightthickness=0)
        self.canvas.pack(fill="both", expand=True)

        # Overlay label with place() so it doesn't steal space from the canvas.
        # The canvas must span the full virtual desktop for multi-monitor selection.
        self.label = tk.Label(self.root, text="Step 1: Draw a box around the video feed on any monitor. Release to lock.", font=("Arial", 28), bg="white", fg="black")
        self.label.place(relx=0.5, y=50, anchor="n")

        self.canvas.bind("<ButtonPress-1>", self.on_press)
        self.canvas.bind("<B1-Motion>", self.on_drag)
        self.canvas.bind("<ButtonRelease-1>", self.on_release)

        self.start_x = None
        self.start_y = None
        self.rect = None
        self.region = None
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
        if frame is not None and frame.size > 0 and frame.any():
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
        self.volume = 0.0
        self.ws     = None
        self._lock        = threading.Lock()
        self._connecting  = False
        self._backoff     = self._BACKOFF_INITIAL
        self._next_attempt = 0.0  # connect immediately on first call
        # Last integer value we actually wrote to the wire, so we can dedupe
        # repeat sends. Some Restim bindings interpret each send as a fresh
        # pulse / re-trigger, which pulses at our tick rate when VSE is
        # pegged at ceiling or floor. None means "nothing sent yet".
        self._last_sent_int = None

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
        ws_url = f"ws://{self.host}:{self.port}/tcode"
        log.info(f"Restim: connect attempt → {ws_url} (axis={self.axis})")
        try:
            ws = websocket.create_connection(ws_url, timeout=2.0)
            with self._lock:
                self.ws         = ws
                self._backoff   = self._BACKOFF_INITIAL  # reset on success
                self._connecting = False
            log.info(f"Restim: CONNECTED at {ws_url}")
            self.set_volume(self.volume, instant=True)
        except Exception as e:
            with self._lock:
                self._backoff      = min(self._backoff * 2, self._BACKOFF_MAX)
                self._next_attempt = time.time() + self._backoff
                self._connecting   = False
            log.warning(f"Restim: connect FAILED ({type(e).__name__}: {e}) — retry in {self._backoff:.0f}s")

    def set_volume(self, vol, instant=False, floor=0.0, ceiling=1.0):
        with self._lock:
            self.volume = max(floor, min(ceiling, vol))
            if not self.ws:
                return
            val_int  = int(round(self.volume * 9999))
            # Dedupe — don't spam Restim with the same value when the math
            # has pegged at a boundary. `instant` bypasses dedupe so manual
            # resets (cum stop, reconnect priming) always land.
            if not instant and val_int == self._last_sent_int:
                return
            interval = 0 if instant else int(VOLUME_UPDATE_INTERVAL * 1000)
            cmd = f"{self.axis}{val_int:04d}I{interval}"
            try:
                self.ws.send(cmd)
                self._last_sent_int = val_int
                log.info(f"Restim: SEND {cmd!r}  (vol={self.volume:.3f})")
            except Exception as e:
                log.warning(f"Restim: SEND FAILED {cmd!r} ({type(e).__name__}: {e}) — dropping socket")
                self.ws = None
                self._last_sent_int = None

    def adjust_volume(self, delta, floor=0.0, ceiling=1.0):
        self.set_volume(self.volume + delta, floor=floor, ceiling=ceiling)


class XToysClient:
    """Sends intensity to xToys via the Private Webhook HTTP endpoint.
    Endpoint : https://xtoys.app/webhook?id=<webhook_id>&action=set-intensity&intensity=<0-100>
    The webhook_id is shown in xToys → https://xtoys.app/me → Private Webhook.
    Stateless HTTP GET — no persistent connection needed.
    """
    _WEBHOOK_URL = "https://xtoys.app/webhook"

    def __init__(self, webhook_id="", port=None):
        self.webhook_id      = webhook_id
        self.volume          = 0.0
        self._lock           = threading.Lock()
        self._last_sent_int  = None
        self._last_error     = None
        self._last_success_t = 0.0
        self._session        = requests.Session()

    @property
    def enabled(self):
        return bool(self.webhook_id.strip())

    @property
    def connected(self):
        """True if a successful send happened in the last 30 s."""
        with self._lock:
            return (time.time() - self._last_success_t) < 30.0

    def maybe_reconnect(self):
        pass  # HTTP is stateless — nothing to reconnect

    def set_volume(self, vol, instant=False, floor=0.0, ceiling=1.0):
        with self._lock:
            self.volume = max(floor, min(ceiling, vol))
            if not self.enabled:
                return
            val_int = int(round(self.volume * 100))
            if not instant and val_int == self._last_sent_int:
                return
            params = {
                "id":        self.webhook_id.strip(),
                "action":    "set-intensity",
                "intensity": val_int,
            }
            try:
                r = self._session.get(self._WEBHOOK_URL, params=params, timeout=5.0)
                if r.status_code == 200:
                    self._last_sent_int  = val_int
                    self._last_error     = None
                    self._last_success_t = time.time()
                    log.info(f"xToys: SEND intensity={val_int}  (vol={self.volume:.3f})")
                else:
                    self._last_error = f"HTTP {r.status_code}"
                    log.warning(f"xToys: SEND failed — {self._last_error}")
            except Exception as e:
                self._last_error = str(e)
                log.warning(f"xToys: SEND failed — {e}")

    def adjust_volume(self, delta, floor=0.0, ceiling=1.0):
        self.set_volume(self.volume + delta, floor=floor, ceiling=ceiling)

    def disconnect(self):
        with self._lock:
            self._last_sent_int  = None
            self._last_success_t = 0.0
            try:
                self._session.close()
            except Exception:
                pass
            self._session = requests.Session()


class HRClient:
    """
    Receives live heart-rate data from Pulsoid via WebSocket.
    Token: https://pulsoid.net/ui/keys  (free tier works with most HR monitors)
    WS endpoint: wss://dev.pulsoid.net/api/v1/data/real_time?access_token=<token>
    Message format: {"measured_at":..., "data":{"heart_rate": 75}}

    modifier() returns a multiplier applied to denial deltas in _tick_volume:
      1.0 — at/below resting BPM (no effect)
      2.0 — at/above peak BPM (denial doubled, rewards halved)
    Linear between the two thresholds, creating natural tightening as arousal builds.
    """
    _WS_URL  = "wss://dev.pulsoid.net/api/v1/data/real_time?access_token={token}"
    _STALE_S = 12.0   # readings older than this are considered disconnected

    def __init__(self, token="", resting_bpm=70, peak_bpm=100):
        self.token       = token.strip()
        self.resting_bpm = resting_bpm
        self.peak_bpm    = peak_bpm
        self._bpm        = None
        self._bpm_hist   = deque(maxlen=6)
        self._last_rx    = 0.0
        self._lock       = threading.Lock()
        self._ws         = None
        self._thread     = None
        self._stop       = threading.Event()

    # ── properties ────────────────────────────────────────────────────────────

    @property
    def enabled(self):
        return bool(self.token)

    @property
    def connected(self):
        with self._lock:
            return self._bpm is not None and (time.time() - self._last_rx) < self._STALE_S

    @property
    def bpm(self):
        with self._lock:
            return self._bpm

    def smooth_bpm(self):
        """Rolling average of recent BPM readings, or None if no data."""
        with self._lock:
            if not self._bpm_hist:
                return None
            return sum(self._bpm_hist) / len(self._bpm_hist)

    def modifier(self):
        """
        Denial multiplier in [1.0, 2.0].
        Returns 1.0 if not connected or peak <= resting.
        """
        if not self.connected:
            return 1.0
        sbpm = self.smooth_bpm()
        if sbpm is None:
            return 1.0
        resting = float(self.resting_bpm)
        peak    = float(self.peak_bpm)
        if peak <= resting:
            return 1.0
        t = max(0.0, min(1.0, (sbpm - resting) / (peak - resting)))
        return 1.0 + t

    # ── lifecycle ──────────────────────────────────────────────────────────────

    def start(self):
        if not self.enabled:
            return
        self._stop.clear()
        if self._thread and self._thread.is_alive():
            return
        self._thread = threading.Thread(target=self._run, daemon=True, name="HRClient")
        self._thread.start()
        log.info(f"HRClient: connecting (token={self.token[:8]}…)")

    def stop(self):
        self._stop.set()
        ws = self._ws
        if ws:
            try:
                ws.close()
            except Exception:
                pass

    def restart(self, token=None, resting_bpm=None, peak_bpm=None):
        """Update config and reconnect."""
        self.stop()
        if token       is not None: self.token       = token.strip()
        if resting_bpm is not None: self.resting_bpm = resting_bpm
        if peak_bpm    is not None: self.peak_bpm    = peak_bpm
        with self._lock:
            self._bpm = None
            self._bpm_hist.clear()
            self._last_rx = 0.0
        if self.enabled:
            self.start()

    # ── WebSocket loop ─────────────────────────────────────────────────────────

    def _run(self):
        while not self._stop.is_set():
            if not self.token:
                time.sleep(2.0)
                continue
            url = self._WS_URL.format(token=self.token)
            try:
                ws = websocket.WebSocketApp(
                    url,
                    on_message=self._on_message,
                    on_error=lambda ws, e: log.warning(f"HRClient WS error: {e}"),
                    on_close=lambda ws, c, m: log.info("HRClient WS closed"),
                )
                self._ws = ws
                ws.run_forever(
                    sslopt={"cert_reqs": ssl.CERT_NONE},
                    ping_interval=20,
                    ping_timeout=10,
                )
            except Exception as e:
                log.warning(f"HRClient: connect error: {e}")
            finally:
                self._ws = None
            if not self._stop.is_set():
                log.debug("HRClient: reconnecting in 8 s…")
                time.sleep(8.0)

    def _on_message(self, ws, raw):
        try:
            data = json.loads(raw)
            bpm  = data.get("data", {}).get("heart_rate")
            if bpm is not None:
                bpm = int(bpm)
                with self._lock:
                    self._bpm = bpm
                    self._bpm_hist.append(bpm)
                    self._last_rx = time.time()
                log.debug(f"HRClient: {bpm} bpm")
        except Exception:
            pass


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
            return None

    def set_volume(self, vol, floor=0.0, ceiling=1.0):
        vol = max(floor, min(ceiling, vol))
        try:
            self._volume_interface.SetMasterVolumeLevelScalar(vol, None)
        except Exception as e:
            log.warning(f"WinAudio: set_volume failed: {e}")

    def adjust_volume(self, delta, floor=0.0, ceiling=1.0):
        cur = self.get_volume()
        if cur is None:
            return
        self.set_volume(cur + delta, floor=floor, ceiling=ceiling)


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


# ── OBS overlay WebSocket server ──────────────────────────────────────────────
import asyncio, struct, hashlib, base64, socket as _socket

class OverlayServer:
    """Tiny WebSocket server that broadcasts JSON state to OBS browser sources."""
    PORT = 12347

    def __init__(self):
        self._clients: list = []
        self._lock = threading.Lock()
        self._loop: asyncio.AbstractEventLoop | None = None
        self._thread: threading.Thread | None = None

    def start(self):
        self._thread = threading.Thread(target=self._run, daemon=True)
        self._thread.start()

    def _run(self):
        self._loop = asyncio.new_event_loop()
        asyncio.set_event_loop(self._loop)
        self._loop.run_until_complete(self._serve())

    async def _serve(self):
        server = await asyncio.start_server(self._handle, '127.0.0.1', self.PORT)
        log.info(f"Overlay WS server listening on ws://127.0.0.1:{self.PORT}")
        async with server:
            await server.serve_forever()

    async def _handle(self, reader: asyncio.StreamReader, writer: asyncio.StreamWriter):
        try:
            request = await asyncio.wait_for(reader.readuntil(b'\r\n\r\n'), timeout=5)
            headers = request.decode(errors='ignore')
            key = None
            for line in headers.split('\r\n'):
                if line.lower().startswith('sec-websocket-key:'):
                    key = line.split(':', 1)[1].strip()

            if not key:
                # Regular HTTP request — serve overlay.html
                html_path = os.path.join(os.path.dirname(__file__), "overlay.html")
                try:
                    with open(html_path, 'rb') as f:
                        body = f.read()
                    writer.write(
                        b'HTTP/1.1 200 OK\r\n'
                        b'Content-Type: text/html; charset=utf-8\r\n'
                        b'Access-Control-Allow-Origin: *\r\n'
                        b'Content-Length: ' + str(len(body)).encode() + b'\r\n'
                        b'Connection: close\r\n\r\n' + body
                    )
                except FileNotFoundError:
                    writer.write(b'HTTP/1.1 404 Not Found\r\nConnection: close\r\n\r\n')
                await writer.drain()
                writer.close()
                return

            # WebSocket upgrade
            accept = base64.b64encode(
                hashlib.sha1((key + '258EAFA5-E914-47DA-95CA-C5AB0DC85B11').encode()).digest()
            ).decode()
            writer.write(
                f'HTTP/1.1 101 Switching Protocols\r\n'
                f'Upgrade: websocket\r\n'
                f'Connection: Upgrade\r\n'
                f'Sec-WebSocket-Accept: {accept}\r\n\r\n'.encode()
            )
            await writer.drain()
        except Exception:
            writer.close()
            return

        with self._lock:
            self._clients.append(writer)
        try:
            # Keep connection alive — read and discard client frames
            while True:
                data = await reader.read(1024)
                if not data:
                    break
        except Exception:
            pass
        finally:
            with self._lock:
                if writer in self._clients:
                    self._clients.remove(writer)
            try:
                writer.close()
            except Exception:
                pass

    def broadcast(self, payload: str):
        """Send a text WebSocket frame to all connected clients."""
        if not self._loop:
            return
        frame = self._ws_text_frame(payload)

        async def _send():
            with self._lock:
                dead = []
                for w in self._clients:
                    try:
                        w.write(frame)
                        await w.drain()
                    except Exception:
                        dead.append(w)
                for w in dead:
                    self._clients.remove(w)

        asyncio.run_coroutine_threadsafe(_send(), self._loop)

    @staticmethod
    def _ws_text_frame(text: str) -> bytes:
        data = text.encode()
        length = len(data)
        if length < 126:
            header = struct.pack('!BB', 0x81, length)
        elif length < 65536:
            header = struct.pack('!BBH', 0x81, 126, length)
        else:
            header = struct.pack('!BBQ', 0x81, 127, length)
        return header + data

    def stop(self):
        if self._loop:
            self._loop.call_soon_threadsafe(self._loop.stop)


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
                data = np.frombuffer(chunk, dtype=np.float32).copy()
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


class Tooltip:
    """Hover tooltip for any tkinter widget."""
    def __init__(self, widget, text, delay=400):
        self.widget = widget
        self.text = text
        self.delay = delay
        self._tip = None
        self._after_id = None
        widget.bind("<Enter>", self._schedule, add="+")
        widget.bind("<Leave>", self._cancel, add="+")

    def _schedule(self, _event=None):
        self._cancel()
        self._after_id = self.widget.after(self.delay, self._show)

    def _cancel(self, _event=None):
        if self._after_id:
            self.widget.after_cancel(self._after_id)
            self._after_id = None
        if self._tip:
            self._tip.destroy()
            self._tip = None

    def _show(self):
        if self._tip:
            return
        x = self.widget.winfo_rootx() + 20
        y = self.widget.winfo_rooty() + self.widget.winfo_height() + 4
        self._tip = tw = tk.Toplevel(self.widget)
        tw.wm_overrideredirect(True)
        tw.wm_geometry(f"+{x}+{y}")
        tw.attributes("-topmost", True)
        lbl = tk.Label(tw, text=self.text, background="#444444", foreground="#eeeeee",
                       relief="solid", borderwidth=1, font=("Segoe UI", 9),
                       padx=6, pady=3, wraplength=280, justify="left")
        lbl.pack()


THEMES = {
    "ReThorn": {
        "BG": "#2d2d2d", "SURFACE": "#383838", "SURFACE2": "#404040",
        "ACCENT": "#F5A623", "ACCENT_H": "#d48e1a",
        "RED": "#FF4444", "RED_HOV": "#cc3636",
        "GREEN": "#3EC941", "GREEN_H": "#32a435",
        "BLUE": "#444CFC", "BLUE_H": "#363dca",
        "YELLOW": "#F5A623", "YELLOW_H": "#d48e1a",
        "TEXT": "#e0e0e8", "TEXT_DIM": "#8a8a8a", "BORDER": "#9C9C9C",
    },
    "OG": {
        "BG": "#1a1a2e", "SURFACE": "#222238", "SURFACE2": "#2a2a42",
        "ACCENT": "#e040fb", "ACCENT_H": "#c030d8",
        "RED": "#cc2200", "RED_HOV": "#991800",
        "GREEN": "#2d8a2d", "GREEN_H": "#1f6b1f",
        "BLUE": "#2a2a55", "BLUE_H": "#222248",
        "YELLOW": "#e040fb", "YELLOW_H": "#c030d8",
        "TEXT": "#e0e0e8", "TEXT_DIM": "#6a6a80", "BORDER": "#3a3a55",
    },
    "Mono": {
        "BG": "#1e1e1e", "SURFACE": "#2a2a2a", "SURFACE2": "#333333",
        "ACCENT": "#aaaaaa", "ACCENT_H": "#888888",
        "RED": "#5a4444", "RED_HOV": "#4a3838",
        "GREEN": "#445a44", "GREEN_H": "#384a38",
        "BLUE": "#44445a", "BLUE_H": "#38384a",
        "YELLOW": "#aaaaaa", "YELLOW_H": "#888888",
        "TEXT": "#cccccc", "TEXT_DIM": "#666666", "BORDER": "#555555",
    },
}
DEFAULT_THEME = "ReThorn"


class App:
    # ── Colour palette (set from theme) ──────────────────────────────────────
    _C_BG        = "#2d2d2d"
    _C_SURFACE   = "#383838"
    _C_SURFACE2  = "#404040"
    _C_ACCENT    = "#F5A623"
    _C_ACCENT_H  = "#d48e1a"
    _C_RED       = "#FF4444"
    _C_RED_HOV   = "#cc3636"
    _C_GREEN     = "#3EC941"
    _C_GREEN_H   = "#32a435"
    _C_BLUE      = "#444CFC"
    _C_BLUE_H    = "#363dca"
    _C_YELLOW    = "#F5A623"
    _C_YELLOW_H  = "#d48e1a"
    _C_TEXT      = "#e0e0e8"
    _C_TEXT_DIM  = "#8a8a8a"
    _C_BORDER    = "#9C9C9C"

    # ── YOLO reanchoring
    _YOLO_INTERVAL = 15    # run detector every N frames
    _YOLO_CONFIRM  = 2     # consecutive detections in same area before reanchoring
    _YOLO_MAX_JUMP = 2.0   # max allowed jump as multiple of current bbox diagonal
    # Tracker plausibility
    _SIZE_RATIO_MAX  = 2.5
    _MAX_JUMP_FACTOR = 2.5

    def _apply_theme(self, name):
        """Apply a colour theme by name."""
        t = THEMES.get(name, THEMES[DEFAULT_THEME])
        self._theme_name = name
        self._C_BG       = t["BG"]
        self._C_SURFACE  = t["SURFACE"]
        self._C_SURFACE2 = t["SURFACE2"]
        self._C_ACCENT   = t["ACCENT"]
        self._C_ACCENT_H = t["ACCENT_H"]
        self._C_RED      = t["RED"]
        self._C_RED_HOV  = t["RED_HOV"]
        self._C_GREEN    = t["GREEN"]
        self._C_GREEN_H  = t["GREEN_H"]
        self._C_BLUE     = t["BLUE"]
        self._C_BLUE_H   = t["BLUE_H"]
        self._C_YELLOW   = t["YELLOW"]
        self._C_YELLOW_H = t["YELLOW_H"]
        self._C_TEXT     = t["TEXT"]
        self._C_TEXT_DIM = t["TEXT_DIM"]
        self._C_BORDER   = t["BORDER"]

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
        self._cum_count = 0
        self._denial_count = 0
        self._cum_override_range = True   # True = cum goes to 100%, False = respect ceiling
        self._cum_stopped = False         # True after "I've CUM" until "Resume"
        self._ui_font_size = 11           # base font size for UI
        self._theme_name = DEFAULT_THEME
        self._cum_time: float | None = None
        self._cum_allowed = False
        self._cum_odds = dict(self._CUM_ODDS_DEFAULT)
        self._denial_phrases = list(self._DENIAL_PHRASES_DEFAULT)

        # AUTO calibration
        self._auto_mode       = True
        self._auto_min_y: float | None = None
        self._auto_max_y: float | None = None
        self._auto_obs_start: float | None = None
        self._auto_last_apply = 0.0

        # Click-to-set height picking mode: None or "Edging"/"Erect"/"Flaccid"
        self._pick_height: str | None = None

        # OBS overlay WebSocket server
        self._overlay = OverlayServer()
        self._overlay.start()
        self._last_overlay_broadcast = 0.0

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
        self.xtoys            = XToysClient()
        self.win_audio        = None
        self.win_devices      = []
        self._orig_win_volume = None  # restored on exit
        self.music_player     = MusicPlayer() if _MINIAUDIO_OK else None
        self.hr_client        = HRClient()

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
        self.restim_on   = tk.BooleanVar(value=True)
        self.xtoys_on    = tk.BooleanVar(value=False)
        self.audio_on    = tk.BooleanVar(value=False)
        self.mp3_on      = tk.BooleanVar(value=False)
        self.hr_on           = tk.BooleanVar(value=False)
        self.hr_token_var    = tk.StringVar(value="")
        self.hr_resting_var  = tk.IntVar(value=70)
        self.hr_peak_var     = tk.IntVar(value=100)
        self.port_var    = tk.StringVar(value="12346")
        # Edging sensitivity: pushes the effective Edging line DOWN (toward
        # Erect) by N pixels, so the Edging state trips earlier than the raw
        # calibrated line. 0 = strict (the line you actually placed).
        self.edge_sens_var = tk.IntVar(value=0)
        # T-code axis VSE sends volume updates on. Restim maps each axis to a
        # parameter ("Volume", "Vibration 0", etc.) in its Websocket panel.
        # Default V0 has historically been bound to Volume in most Restim
        # session files — users who've changed that mapping can point VSE at
        # whichever axis their session uses.
        self.tcode_axis_var    = tk.StringVar(value=TCODE_AXIS)
        self.xtoys_id_var      = tk.StringVar(value="")

        # Pre-load font size, theme, cum override before UI build
        try:
            if CONFIG_PATH.exists():
                _pre = json.loads(CONFIG_PATH.read_text())
                if "ui_font_size" in _pre:
                    self._ui_font_size = int(_pre["ui_font_size"])
                if "cum_override_range" in _pre:
                    self._cum_override_range = bool(_pre["cum_override_range"])
                if "theme" in _pre and _pre["theme"] in THEMES:
                    self._theme_name = _pre["theme"]
        except Exception:
            pass
        self._apply_theme(self._theme_name)
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
        root.minsize(540, 500)

        P   = 12          # standard outer padding
        fs  = self._ui_font_size
        lbl = ctk.CTkFont(size=fs, weight="bold")
        btn = ctk.CTkFont(size=fs + 1, weight="bold")

        # ── Menu bar ─────────────────────────────────────────────────────────
        menubar = tk.Menu(root, bg=self._C_BG, fg=self._C_TEXT,
                          activebackground=self._C_ACCENT, activeforeground="white",
                          relief="flat", borderwidth=0)
        settings_menu = tk.Menu(menubar, tearoff=0,
                                bg=self._C_BG, fg=self._C_TEXT,
                                activebackground=self._C_ACCENT, activeforeground="white")
        settings_menu.add_command(label="Settings...", command=self._open_settings)
        menubar.add_cascade(label="Settings", menu=settings_menu)

        about_menu = tk.Menu(menubar, tearoff=0,
                             bg=self._C_BG, fg=self._C_TEXT,
                             activebackground=self._C_ACCENT, activeforeground="white")
        about_menu.add_command(label=f"Version  v{VERSION}", state="disabled")
        about_menu.add_command(label="Dev: Sir Thorn", state="disabled")
        about_menu.add_separator()
        about_menu.add_command(label="GitHub (latest release)",
                               command=lambda: webbrowser.open("https://github.com/blucrew/VisualStimEdger/releases/latest"))
        about_menu.add_command(label="Ko-fi (support)",
                               command=lambda: webbrowser.open("https://ko-fi.com/stimstation"))
        menubar.add_cascade(label="About", menu=about_menu)

        root.configure(menu=menubar)

        # ── Update banner (hidden until needed) ───────────────────────────────
        self._update_banner = ctk.CTkFrame(root, fg_color=self._C_SURFACE, corner_radius=0)
        self._update_label  = ctk.CTkLabel(self._update_banner, text="",
                                           text_color="white", font=lbl)
        self._update_label.pack(side=tk.LEFT, padx=P, pady=6)
        self._update_btn = ctk.CTkButton(self._update_banner, text="Download", width=100,
                                         fg_color=self._C_ACCENT, hover_color=self._C_ACCENT_H,
                                         text_color="white", font=lbl, corner_radius=4)
        self._update_btn.pack(side=tk.RIGHT, padx=P, pady=6)

        # ── Video + calibration buttons ───────────────────────────────────────
        top_frame = ctk.CTkFrame(root, fg_color="transparent")
        top_frame.pack(padx=P, pady=(P, 4), fill=tk.X)
        self._first_widget = top_frame

        # video — plain tk.Label so ImageTk works without wrapping
        vid_shell = ctk.CTkFrame(top_frame, fg_color=self._C_SURFACE,
                                 corner_radius=8, border_width=1, border_color=self._C_BORDER,
                                 width=330, height=260)
        vid_shell.pack(side=tk.LEFT)
        vid_shell.pack_propagate(False)
        self.video_label = tk.Label(vid_shell, bg=self._C_SURFACE)
        self.video_label.pack(padx=3, pady=(3, 0), fill=tk.BOTH, expand=True)
        self._snark_label = ctk.CTkLabel(vid_shell, text="", font=ctk.CTkFont(size=11, slant="italic"),
                                          text_color="#ff4444", height=20)
        self._snark_label.pack(padx=3, pady=(0, 3))

        # height buttons — stacked right of video
        hbf = ctk.CTkFrame(top_frame, fg_color="transparent", width=180, height=260)
        hbf.pack(side=tk.LEFT, fill=tk.Y, padx=(10, 0))
        hbf.pack_propagate(False)

        def _hbtn_row(parent, text, set_cmd, pick_cmd, color, hover):
            row = ctk.CTkFrame(parent, fg_color="transparent")
            row.pack(fill=tk.X, pady=1)
            row.columnconfigure(0, weight=1, uniform="hbtn")
            row.columnconfigure(1, weight=1, uniform="hbtn")
            ctk.CTkButton(row, text=text, command=set_cmd, font=btn, height=28,
                          fg_color=color, hover_color=hover,
                          text_color="white", corner_radius=4
                          ).grid(row=0, column=0, sticky="ew", padx=(0, 2))
            ctk.CTkButton(row, text="Manual \u271a", command=pick_cmd, font=ctk.CTkFont(size=10),
                          height=28,
                          fg_color=color, hover_color=hover,
                          text_color="white", corner_radius=4
                          ).grid(row=0, column=1, sticky="ew")
            return row

        self._auto_btn = ctk.CTkButton(
            hbf, text="AUTO  (observing...)", command=self._toggle_auto,
            font=btn, height=28, corner_radius=4,
            fg_color=self._C_ACCENT, hover_color=self._C_ACCENT_H, text_color="white")
        self._auto_btn.pack(fill=tk.X, pady=(0, 3))
        Tooltip(self._auto_btn, "Auto-calibrate heights by observing motion range")

        _hbtn_row(hbf, "Edging",  self._set_edging,  lambda: self._start_pick("Edging"),  self._C_RED,    self._C_RED_HOV)
        _hbtn_row(hbf, "Erect",   self._set_erect,   lambda: self._start_pick("Erect"),   self._C_GREEN,  self._C_GREEN_H)
        _hbtn_row(hbf, "Flaccid", self._set_flaccid, lambda: self._start_pick("Flaccid"), self._C_BLUE,   self._C_BLUE_H)

        # cum buttons anchored at bottom
        cum_row = ctk.CTkFrame(hbf, fg_color="transparent")
        cum_row.pack(side=tk.BOTTOM, fill=tk.X, pady=1)
        cum_row.columnconfigure(0, weight=1, uniform="cumbtn")
        cum_row.columnconfigure(1, weight=1, uniform="cumbtn")
        self._letmecum_btn = ctk.CTkButton(
            cum_row, text="Let me cum?", command=self._on_letmecum,
            font=btn, height=30, corner_radius=4,
            fg_color=self._C_GREEN, hover_color=self._C_GREEN_H, text_color="white")
        self._letmecum_btn.grid(row=0, column=0, sticky="ew", padx=(0, 2))
        Tooltip(self._letmecum_btn, "Roll the dice — odds depend on aggressiveness. Win = temporary full volume permission")
        self._cum_btn = ctk.CTkButton(cum_row, text="I've CUM", command=self._on_cum,
                      font=btn, height=30, corner_radius=4,
                      fg_color="#e0e0e8", hover_color="#c8c8d0",
                      text_color=self._C_BG)
        self._cum_btn.grid(row=0, column=1, sticky="ew")
        Tooltip(self._cum_btn, "Press after finishing — volume drops to 0 and stays there. Press again (as 'Resume') if you want to go another round.")

        def _divider():
            ctk.CTkFrame(root, height=1, fg_color=self._C_BORDER).pack(fill=tk.X, padx=P, pady=1)

        _divider()

        # ── Volume floor / ceiling ────────────────────────────────────────────
        vol_card = ctk.CTkFrame(root, fg_color=self._C_SURFACE, corner_radius=8)
        vol_card.pack(fill=tk.X, padx=P, pady=2)

        vol_row = ctk.CTkFrame(vol_card, fg_color="transparent")
        vol_row.pack(fill=tk.X, padx=12, pady=6)
        ctk.CTkLabel(vol_row, text="Volume Range", font=lbl,
                     text_color=self._C_TEXT, anchor="w").pack(side=tk.LEFT)
        self._range_lbl = ctk.CTkLabel(vol_row, text="0% – 100%", font=lbl,
                                       text_color=self._C_YELLOW, anchor="e")
        self._range_lbl.pack(side=tk.RIGHT)

        # ── Custom dual-handle range slider ──
        _track_h, _handle_r = 4, 7
        _marker_h = 10
        _canvas_h = _handle_r * 2 + 4 + _marker_h
        self._range_cv = tk.Canvas(vol_card, height=_canvas_h, bg=self._C_SURFACE,
                                   highlightthickness=0)
        self._range_cv.pack(fill=tk.X, padx=18, pady=(0, 8))

        self._range_drag = None  # 'lo' or 'hi'

        def _range_draw(event=None):
            c = self._range_cv
            c.delete("all")
            w = c.winfo_width()
            if w < 20:
                return
            pad = _handle_r + 2
            track_w = w - pad * 2
            cy = _marker_h + (_canvas_h - _marker_h) // 2
            lo = self.min_vol_var.get() / 100.0
            hi = self.max_vol_var.get() / 100.0
            lx = pad + lo * track_w
            hx = pad + hi * track_w
            # background track
            c.create_line(pad, cy, pad + track_w, cy, fill=self._C_SURFACE2,
                          width=_track_h, capstyle="round")
            # active range
            c.create_line(lx, cy, hx, cy, fill=self._C_ACCENT,
                          width=_track_h, capstyle="round")
            # handles
            for x in (lx, hx):
                c.create_oval(x - _handle_r, cy - _handle_r, x + _handle_r, cy + _handle_r,
                              fill=self._C_ACCENT, outline="")
            # cum release marker — inverted triangle above the bar
            cum_frac = 1.0 if self._cum_override_range else hi
            cum_x = pad + cum_frac * track_w
            ts = 5  # triangle half-width
            c.create_polygon(cum_x - ts, 0, cum_x + ts, 0, cum_x, _marker_h - 2,
                             fill="#3EC941", outline="")

        def _range_press(e):
            w = self._range_cv.winfo_width()
            pad = _handle_r + 2
            track_w = w - pad * 2
            if track_w < 1:
                return
            frac = max(0.0, min(1.0, (e.x - pad) / track_w))
            lo = self.min_vol_var.get() / 100.0
            hi = self.max_vol_var.get() / 100.0
            if abs(frac - lo) < abs(frac - hi):
                self._range_drag = 'lo'
            else:
                self._range_drag = 'hi'
            _range_move(e)

        def _range_move(e):
            if not self._range_drag:
                return
            w = self._range_cv.winfo_width()
            pad = _handle_r + 2
            track_w = w - pad * 2
            if track_w < 1:
                return
            frac = max(0.0, min(1.0, (e.x - pad) / track_w))
            val = round(frac * 100)
            if self._range_drag == 'lo':
                val = min(val, int(self.max_vol_var.get()))
                self.min_vol_var.set(val)
            else:
                val = max(val, int(self.min_vol_var.get()))
                self.max_vol_var.set(val)
            lo = int(self.min_vol_var.get())
            hi = int(self.max_vol_var.get())
            self._range_lbl.configure(text=f"{lo}% – {hi}%")
            _range_draw()

        def _range_release(e):
            self._range_drag = None
            self._save_config()

        self._range_cv.bind("<ButtonPress-1>", _range_press)
        self._range_cv.bind("<B1-Motion>", _range_move)
        self._range_cv.bind("<ButtonRelease-1>", _range_release)
        self._range_cv.bind("<Configure>", _range_draw)
        self._range_draw = _range_draw
        Tooltip(self._range_cv,
               "Drag left handle = volume floor, right handle = ceiling.\n"
               "Green triangle = where 'Let me cum?' sends volume.\n"
               "Change override behavior in Settings > Cum Volume Behavior.")

        _divider()

        # ── Aggressiveness ────────────────────────────────────────────────────
        aggr_card = ctk.CTkFrame(root, fg_color=self._C_SURFACE, corner_radius=8)
        aggr_card.pack(fill=tk.X, padx=P, pady=2)
        Tooltip(aggr_card,
               "Controls how fast volume ramps and 'Let me cum?' odds.\n"
               "Easy: gentle ramps, 1-in-2 odds, 5 min grant\n"
               "Middle: moderate ramps, 1-in-4 odds, 5 min grant\n"
               "Hard: fast ramps, 1-in-6 odds, 3 min grant\n"
               "Expert: aggressive ramps, 1-in-30 odds, 1 min grant")
        aggr_row = ctk.CTkFrame(aggr_card, fg_color="transparent")
        aggr_row.pack(fill=tk.X, padx=12, pady=10)
        ctk.CTkLabel(aggr_row, text="Aggressiveness", font=lbl,
                     text_color=self._C_TEXT).pack(side=tk.LEFT)
        self._aggr_seg = ctk.CTkSegmentedButton(
            aggr_row, values=list(AGGR_LEVELS.keys()), variable=self.aggr_var,
            command=self._on_aggr_change,
            selected_color=self._C_ACCENT, selected_hover_color=self._C_ACCENT_H,
            unselected_color=self._C_SURFACE2, unselected_hover_color="#4a4a4a",
            font=ctk.CTkFont(size=11, weight="bold"),
            text_color=self._C_TEXT,
        )
        self._aggr_seg.pack(side=tk.RIGHT)

        _divider()

        # ── Output mode ───────────────────────────────────────────────────────
        mode_card = ctk.CTkFrame(root, fg_color=self._C_SURFACE, corner_radius=8)
        self._mode_card = mode_card
        mode_card.pack(fill=tk.X, padx=P, pady=2)
        mode_row = ctk.CTkFrame(mode_card, fg_color="transparent")
        mode_row.pack(fill=tk.X, padx=12, pady=10)
        ctk.CTkLabel(mode_row, text="Output", font=lbl,
                     text_color=self._C_TEXT).pack(side=tk.LEFT)

        _mode_defs = [
            ("restim",  "\u26a1 restim",            self.restim_on, self._C_ACCENT,  self._C_ACCENT_H),
            ("xtoys",   "xToys",                    self.xtoys_on,  self._C_GREEN,   self._C_GREEN_H),
            ("windows", "\u229e\U0001f50a windows",  self.audio_on,  self._C_BLUE,    self._C_BLUE_H),
            ("hr",      "\u2665 HR",                self.hr_on,     "#e91e63",       "#c2185b"),
        ]
        if _MINIAUDIO_OK:
            _mode_defs.append(("mp3", "\u266b mp3", self.mp3_on, "#9c27b0", "#7b1fa2"))

        self._mode_btns = {}
        mbf = ctk.CTkFrame(mode_row, fg_color="transparent")
        mbf.pack(side=tk.RIGHT)
        for mode_key, mode_label, boolvar, color, hover in _mode_defs:
            def _toggle(m=mode_key, bv=boolvar):
                new_val = not bv.get()
                bv.set(new_val)
                if new_val:
                    # Audio / MP3 are mutually exclusive
                    if m == "windows":
                        self.mp3_on.set(False)
                        if self.music_player: self.music_player.stop()
                    elif m == "mp3":
                        self.audio_on.set(False)
                else:
                    if m == "mp3" and self.music_player:
                        self.music_player.stop()
                self._on_output_change()
                self._update_output_btns()
            b = ctk.CTkButton(mbf, text=mode_label, command=_toggle,
                              font=ctk.CTkFont(size=11), height=28, corner_radius=4,
                              fg_color=color, hover_color=hover,
                              text_color="white", width=10)
            b.pack(side=tk.LEFT, padx=1)
            self._mode_btns[mode_key] = (b, color, hover, boolvar)

        def _update_output_btns():
            for mk, (bt, col, hov, bv) in self._mode_btns.items():
                if bv.get():
                    bt.configure(fg_color=col, hover_color=hov)
                else:
                    bt.configure(fg_color=self._C_SURFACE2, hover_color="#4a4a4a")
        self._update_output_btns = _update_output_btns
        _update_output_btns()

        # Restim options panel
        self._restim_opts = ctk.CTkFrame(root, fg_color=self._C_SURFACE2, corner_radius=6)
        ro = ctk.CTkFrame(self._restim_opts, fg_color="transparent")
        ro.pack(padx=12, pady=7)
        ctk.CTkLabel(ro, text="Restim Port:", font=lbl, text_color=self._C_TEXT).pack(side=tk.LEFT, padx=(0, 6))
        ctk.CTkEntry(ro, textvariable=self.port_var, width=72,
                     fg_color=self._C_SURFACE, border_color=self._C_BORDER,
                     text_color=self._C_TEXT).pack(side=tk.LEFT)
        self.port_var.trace_add("write", self._on_port_change)
        # OBS overlay URL on same row
        _overlay_url = f"http://127.0.0.1:{OverlayServer.PORT}"
        ctk.CTkLabel(ro, text="OBS Overlay:", font=lbl,
                     text_color=self._C_TEXT).pack(side=tk.LEFT, padx=(16, 6))
        _obs_entry = ctk.CTkEntry(ro, width=170,
                                  fg_color=self._C_SURFACE, border_color=self._C_BORDER,
                                  text_color=self._C_TEXT, font=ctk.CTkFont(size=11))
        _obs_entry.insert(0, _overlay_url)
        _obs_entry.configure(state="readonly")
        _obs_entry.pack(side=tk.LEFT, padx=(0, 4))
        def _copy_overlay_url():
            self.root.clipboard_clear()
            self.root.clipboard_append(_overlay_url)
            _copy_btn.configure(text="Copied!")
            self.root.after(1500, lambda: _copy_btn.configure(text="Copy"))
        _copy_btn = ctk.CTkButton(ro, text="Copy", command=_copy_overlay_url, width=48,
                                  fg_color=self._C_SURFACE, hover_color="#4a4a4a",
                                  text_color=self._C_TEXT, border_width=1, border_color=self._C_BORDER,
                                  font=ctk.CTkFont(size=10))
        _copy_btn.pack(side=tk.LEFT)

        # xToys options panel
        self._xtoys_opts = ctk.CTkFrame(root, fg_color=self._C_SURFACE2, corner_radius=6)
        xo = ctk.CTkFrame(self._xtoys_opts, fg_color="transparent")
        xo.pack(fill=tk.X, padx=12, pady=(7, 2))
        ctk.CTkLabel(xo, text="Webhook ID:", font=lbl,
                     text_color=self._C_TEXT).pack(side=tk.LEFT, padx=(0, 6))
        ctk.CTkEntry(xo, textvariable=self.xtoys_id_var, width=240,
                     fg_color=self._C_SURFACE, border_color=self._C_BORDER,
                     text_color=self._C_TEXT,
                     placeholder_text="e.g. 8hR5acKTCx2s").pack(side=tk.LEFT, padx=(0, 8))
        ctk.CTkLabel(self._xtoys_opts,
                     text="Get your Webhook ID at xtoys.app/me \u2192 Private Webhook. In xToys: load the VisualStimEdger script \u2192 Connections \u2192 add your toy \u2192 enable Private Webhook.",
                     font=ctk.CTkFont(size=9), text_color=self._C_TEXT_DIM,
                     wraplength=380, justify="left").pack(anchor="w", padx=12, pady=(0, 6))
        self.xtoys_id_var.trace_add("write", self._on_xtoys_id_change)

        # ── Heart Rate (Pulsoid) options panel ────────────────────────────────
        self._hr_opts = ctk.CTkFrame(root, fg_color=self._C_SURFACE2, corner_radius=6)

        hr_row1 = ctk.CTkFrame(self._hr_opts, fg_color="transparent")
        hr_row1.pack(fill=tk.X, padx=12, pady=(7, 2))
        ctk.CTkLabel(hr_row1, text="Pulsoid Token:", font=lbl,
                     text_color=self._C_TEXT).pack(side=tk.LEFT, padx=(0, 6))
        ctk.CTkEntry(hr_row1, textvariable=self.hr_token_var, width=210,
                     fg_color=self._C_SURFACE, border_color=self._C_BORDER,
                     text_color=self._C_TEXT, show="\u2022",
                     placeholder_text="paste from pulsoid.net/ui/keys").pack(side=tk.LEFT, padx=(0, 8))
        self._hr_bpm_label = ctk.CTkLabel(hr_row1, text="\u2665 -- bpm",
                                          font=ctk.CTkFont(size=fs + 1, weight="bold"),
                                          text_color="#e91e63")
        self._hr_bpm_label.pack(side=tk.LEFT)

        hr_row2 = ctk.CTkFrame(self._hr_opts, fg_color="transparent")
        hr_row2.pack(fill=tk.X, padx=12, pady=(0, 2))

        def _hr_slider_group(parent, label, var, lo, hi):
            ctk.CTkLabel(parent, text=label, font=lbl,
                         text_color=self._C_TEXT).pack(side=tk.LEFT, padx=(0, 4))
            lbl_val = ctk.CTkLabel(parent, text=f"{var.get()} bpm",
                                   font=lbl, text_color=self._C_TEXT_DIM, width=60)
            lbl_val.pack(side=tk.LEFT, padx=(0, 6))
            def _upd(v, lv=lbl_val, sv=var):
                sv.set(int(float(v)))
                lv.configure(text=f"{int(float(v))} bpm")
                self.hr_client.resting_bpm = self.hr_resting_var.get()
                self.hr_client.peak_bpm    = self.hr_peak_var.get()
            ctk.CTkSlider(parent, from_=lo, to=hi, variable=var,
                          command=_upd, width=110,
                          button_color="#e91e63", button_hover_color="#c2185b",
                          progress_color="#e91e63").pack(side=tk.LEFT)

        _hr_slider_group(hr_row2, "Resting:", self.hr_resting_var, 40, 90)
        ctk.CTkLabel(hr_row2, text="   ", text_color="transparent").pack(side=tk.LEFT)  # spacer
        _hr_slider_group(hr_row2, "Peak:", self.hr_peak_var, 80, 170)

        ctk.CTkLabel(self._hr_opts,
                     text=("Higher HR \u2192 stronger denial, slower rewards.  "
                           "Get a free token at pulsoid.net/ui/keys \u2014 works with Polar, Garmin, "
                           "Apple Watch, most BLE chest straps via the Pulsoid app."),
                     font=ctk.CTkFont(size=9), text_color=self._C_TEXT_DIM,
                     wraplength=390, justify="left").pack(anchor="w", padx=12, pady=(2, 6))
        self.hr_token_var.trace_add("write", self._on_hr_token_change)

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
                      fg_color=self._C_SURFACE2, hover_color="#4a4a4a",
                      text_color=self._C_TEXT, border_width=1, border_color=self._C_BORDER,
                      font=ctk.CTkFont(size=10)).pack(side=tk.LEFT)

        self._restim_opts.pack(fill=tk.X, padx=P, pady=(0, 4))  # default; _on_output_change re-packs after load

        # ── MP3 player options panel ───────────────────────────────────────────
        self._mp3_opts = ctk.CTkFrame(root, fg_color=self._C_SURFACE2, corner_radius=6)
        if _MINIAUDIO_OK:
            mo = ctk.CTkFrame(self._mp3_opts, fg_color="transparent")
            mo.pack(fill=tk.X, padx=12, pady=(8, 4))

            # File / folder load buttons
            load_row = ctk.CTkFrame(mo, fg_color="transparent")
            load_row.pack(fill=tk.X, pady=(0, 4))
            ctk.CTkButton(load_row, text="📁 Load File", width=110, height=28,
                          fg_color=self._C_SURFACE, hover_color="#4a4a4a",
                          text_color=self._C_TEXT, border_width=1, border_color=self._C_BORDER,
                          font=ctk.CTkFont(size=10),
                          command=self._mp3_load_file).pack(side=tk.LEFT, padx=(0, 6))
            ctk.CTkButton(load_row, text="📂 Load Folder", width=120, height=28,
                          fg_color=self._C_SURFACE, hover_color="#4a4a4a",
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
                          hover_color="#4a4a4a", text_color=self._C_TEXT,
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
                unselected_color=self._C_SURFACE, unselected_hover_color="#4a4a4a",
                font=ctk.CTkFont(size=10), text_color=self._C_TEXT,
            ).pack(side=tk.LEFT)

        _divider()

        # ── Controls row ──────────────────────────────────────────────────────
        ctrl = ctk.CTkFrame(root, fg_color="transparent")
        ctrl.pack(fill=tk.X, padx=P, pady=3)

        def _ghost_btn(parent, text, cmd, **kw):
            kw.setdefault("font", ctk.CTkFont(size=10))
            return ctk.CTkButton(
                parent, text=text, command=cmd, height=34,
                fg_color=self._C_SURFACE2, hover_color="#4a4a4a",
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
        Tooltip(self._hold_btn, "Freeze volume at current level — tracking continues but volume won't change")
        self._about_btn = _ghost_btn(ctrl, "ⓘ", self._show_about_menu, width=34)
        self._about_btn.pack(side=tk.LEFT, padx=(4, 0))

        _divider()

        # ── Status labels ─────────────────────────────────────────────────────
        self.info_label = ctk.CTkLabel(
            root, text="State: --  |  Vol: --  |  WS: Disconnected",
            font=ctk.CTkFont(size=15, weight="bold"), text_color=self._C_TEXT,
        )
        self.info_label.pack(pady=(4, 1))

        self.stats_label = ctk.CTkLabel(
            root, text="Session: 00:00  |  Edges: 0",
            font=ctk.CTkFont(size=10), text_color=self._C_TEXT_DIM,
        )
        self.stats_label.pack(pady=(0, 4))

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
                "edge_sens":   int(self.edge_sens_var.get()),
                "tcode_axis":  self.tcode_axis_var.get(),
                "xtoys_id":    self.xtoys_id_var.get(),
                "outputs": {
                    "restim":  self.restim_on.get(),
                    "xtoys":   self.xtoys_on.get(),
                    "windows": self.audio_on.get(),
                    "mp3":     self.mp3_on.get(),
                    "hr":      self.hr_on.get(),
                },
                "hr_token":   self.hr_token_var.get(),
                "hr_resting": self.hr_resting_var.get(),
                "hr_peak":    self.hr_peak_var.get(),
                "port":        self.port_var.get(),
                "device_name": self._device_combo.get(),
                "mp3_path":    getattr(self, '_mp3_last_path', ""),
                "mp3_path_type": getattr(self, '_mp3_last_type', "file"),
                "mp3_loop":    self._mp3_loop_var.get() if _MINIAUDIO_OK else "folder",
                "cum_odds":    self._cum_odds,
                "denial_phrases": self._denial_phrases,
                "cum_override_range": self._cum_override_range,
                "ui_font_size": self._ui_font_size,
                "theme": self._theme_name,
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
            if "edge_sens" in data:
                try:
                    self.edge_sens_var.set(max(0, min(500, int(data["edge_sens"]))))
                    if hasattr(self, "_sens_lbl"):
                        self._sens_lbl.configure(text=f"{int(self.edge_sens_var.get())} px")
                except Exception:
                    pass
            if "xtoys_id" in data:
                self.xtoys_id_var.set(str(data.get("xtoys_id", "")))
                self.xtoys.webhook_id = self.xtoys_id_var.get()
            if "tcode_axis" in data:
                try:
                    ax = str(data["tcode_axis"]).strip().upper()
                    if len(ax) >= 2 and ax[0].isalpha() and ax[1:].isdigit():
                        self.tcode_axis_var.set(ax)
                        self.restim.axis = ax
                except Exception:
                    pass
            if "outputs" in data:
                outs = data["outputs"]
                self.restim_on.set(bool(outs.get("restim", True)))
                self.xtoys_on.set(bool(outs.get("xtoys", False)))
                self.audio_on.set(bool(outs.get("windows", False)))
                self.mp3_on.set(bool(outs.get("mp3", False)) and _MINIAUDIO_OK)
                self.hr_on.set(bool(outs.get("hr", False)))
                self._on_output_change()
                self._update_output_btns()
            elif "mode" in data:
                # Migrate old single-mode config
                mode = data["mode"]
                self.restim_on.set(mode == "restim")
                self.xtoys_on.set(mode == "xtoys")
                self.audio_on.set(mode == "windows")
                self.mp3_on.set(mode == "mp3" and _MINIAUDIO_OK)
                self._on_output_change()
                self._update_output_btns()
            # Heart rate config
            if "hr_token" in data:
                tok = str(data["hr_token"])
                self.hr_token_var.set(tok)
                self.hr_client.token = tok.strip()
            if "hr_resting" in data:
                try:
                    v = max(40, min(90, int(data["hr_resting"])))
                    self.hr_resting_var.set(v)
                    self.hr_client.resting_bpm = v
                except Exception:
                    pass
            if "hr_peak" in data:
                try:
                    v = max(80, min(170, int(data["hr_peak"])))
                    self.hr_peak_var.set(v)
                    self.hr_client.peak_bpm = v
                except Exception:
                    pass
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
            if "cum_odds" in data and isinstance(data["cum_odds"], dict):
                self._cum_odds.update(data["cum_odds"])
            if "denial_phrases" in data and isinstance(data["denial_phrases"], list):
                self._denial_phrases = data["denial_phrases"]
            if "cum_override_range" in data:
                self._cum_override_range = bool(data["cum_override_range"])
            # Legacy keys cum_silence / cum_ramp are silently ignored.
            if "ui_font_size" in data:
                self._ui_font_size = int(data["ui_font_size"])
            if "theme" in data and data["theme"] in THEMES:
                self._theme_name = data["theme"]
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

        # Stop HR client
        try:
            self.hr_client.stop()
        except Exception:
            pass

        # Stop overlay server
        if hasattr(self, '_overlay'):
            try:
                self._overlay.stop()
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
                time.sleep(0.03)   # ~30 fps cap — no point capturing faster
            else:
                time.sleep(0.25)

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


    _AUTO_INTERVAL      = 2.0   # seconds between re-applies
    _AUTO_GOOD_RANGE    = 40    # pixels of range before setting Edging/Erect

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
                                     fg_color=self._C_SURFACE2, hover_color="#4a4a4a")

    def _auto_feed(self, y: float):
        """Called every heavy frame while AUTO is on and tracking is good.

        Strategy: Flaccid is set immediately to the lowest observed position
        (highest Y). Edging and Erect are only set once we've seen enough
        vertical range to be confident — you won't get a full stroke in
        the first few seconds.
        """
        now = time.time()
        if self._auto_obs_start is None:
            self._auto_obs_start = now

        # Expand observed range
        changed = False
        if self._auto_min_y is None or y < self._auto_min_y:
            self._auto_min_y = y
            changed = True
        if self._auto_max_y is None or y > self._auto_max_y:
            self._auto_max_y = y
            changed = True

        if not changed and now - self._auto_last_apply < self._AUTO_INTERVAL:
            return
        self._auto_last_apply = now

        # Always set Flaccid to lowest observed position (highest Y)
        self.heights["Flaccid"] = self._auto_max_y

        rng = self._auto_max_y - self._auto_min_y
        if rng >= self._AUTO_GOOD_RANGE:
            # Enough range — set Edging and Erect
            self.heights["Edging"] = self._auto_min_y + 0.05 * rng
            self.heights["Erect"]  = (self._auto_min_y + self._auto_max_y) / 2
            if getattr(self, '_auto_btn_state', None) != "active":
                self._auto_btn_state = "active"
                self._auto_btn.configure(text="AUTO  \u2713", fg_color="#3EC941", hover_color="#32a435")
            log.debug(f"AUTO heights: edging={self.heights['Edging']:.0f} "
                      f"erect={self.heights['Erect']:.0f} flaccid={self.heights['Flaccid']:.0f}")
        else:
            self._auto_btn.configure(text=f"AUTO  (flaccid set, need more range)")

    def _disable_auto(self):
        """Turn off AUTO so manual height settings aren't overwritten."""
        if self._auto_mode:
            self._auto_mode = False
            self._auto_btn.configure(text="AUTO  (off)",
                                     fg_color=self._C_SURFACE2, hover_color="#4a4a4a")

    def _cancel_pick(self):
        """Cancel any active Manual pick mode."""
        if self._pick_height:
            self._pick_height = None
            self.video_label.configure(cursor="")
            self.video_label.unbind("<Button-1>")

    def _set_edging(self):
        self._disable_auto()
        self._cancel_pick()
        self.heights["Edging"] = self.head_y
        log.info(f"Edging height set at Y={self.head_y}")
        self._save_config()

    def _set_erect(self):
        self._disable_auto()
        self._cancel_pick()
        self.heights["Erect"] = self.head_y
        log.info(f"Erect height set at Y={self.head_y}")
        self._save_config()

    def _set_flaccid(self):
        self._disable_auto()
        self._cancel_pick()
        self.heights["Flaccid"] = self.head_y
        log.info(f"Flaccid height set at Y={self.head_y}")
        self._save_config()

    def _start_pick(self, which: str):
        """Enter click-to-set mode: next click on video sets that height."""
        self._disable_auto()
        self._cancel_pick()
        self._pick_height = which
        self.video_label.configure(cursor="crosshair")
        self.video_label.bind("<Button-1>", self._on_video_click)
        log.info(f"Pick mode: click on video to set {which} height")

    def _on_video_click(self, event):
        """Handle click on video label to set the height being picked.

        event.y is in widget-space (label pixels, with letterbox padding around
        the scaled image). We need to invert the display transform so the stored
        height lives in the SAME coordinate system as self.head_y and the frame
        that _draw_height_lines paints on — raw source-frame pixels."""
        which = self._pick_height
        if not which:
            return
        scale    = getattr(self, "_disp_scale", 1.0) or 1.0
        offset_y = getattr(self, "_disp_offset_y", 0)
        frame_h  = getattr(self, "_disp_frame_h", None)
        y_in_img = event.y - offset_y
        frame_y  = int(round(y_in_img / scale))
        if frame_h:
            frame_y = max(0, min(frame_h - 1, frame_y))
        self.heights[which] = frame_y
        log.info(f"{which} height set at Y={frame_y} via click "
                 f"(label y={event.y}, scale={scale:.3f}, offset_y={offset_y})")
        self._save_config()
        self._pick_height = None
        self.video_label.configure(cursor="")
        self.video_label.unbind("<Button-1>")

    def _reset_heights(self):
        """Clear calibrated heights — called whenever the feed changes."""
        self.heights = {"Edging": None, "Erect": None, "Flaccid": None}
        self._auto_min_y = None
        self._auto_max_y = None
        self._auto_obs_start = None
        self._auto_last_apply = 0.0
        self._head_y_history.clear()
        if hasattr(self, '_auto_btn_state'):
            self._auto_btn_state = None
        if hasattr(self, '_auto_btn'):
            self._auto_btn.configure(text="AUTO", fg_color=self._C_SURFACE2,
                                     hover_color="#4a4a4a")
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
                self._head_y_history.clear()
                self.tracking_ok = True
        finally:
            self.tracking_paused = False

    _AGGR_COLORS = {
        "Easy":   ("#7c3aed", "#6525d0"),
        "Middle": ("#F5A623", "#d48e1a"),
        "Hard":   ("#cc6600", "#994c00"),
        "Expert": ("#FF4444", "#cc3636"),
    }

    def _on_aggr_change(self, val=None):
        val = val or self.aggr_var.get()
        col, hov = self._AGGR_COLORS.get(val, (self._C_RED, self._C_RED_HOV))
        self._aggr_seg.configure(selected_color=col, selected_hover_color=hov)
        self._save_config()

    # Odds of "Let me cum?" being granted per aggressiveness level
    _CUM_ODDS_DEFAULT = {"Easy": 2, "Middle": 4, "Hard": 6, "Expert": 30}
    _CUM_DENY_COOLDOWN = {"Easy": 30, "Middle": 30, "Hard": 60, "Expert": 120}
    _CUM_GRANT_TIME = {"Easy": 300, "Middle": 300, "Hard": 180, "Expert": 60}
    _DENIAL_PHRASES_DEFAULT = [
        "Not this time.", "Keep trying, gooner.",
        "Denied. Back to edging.", "Nope. Suffer.",
        "The answer is no.", "Maybe next time ;)",
        "Absolutely not.", "You wish.",
        "Earn it.", "Not even close.",
        "Haha, no.", "Try again later, perv.",
        "Permission denied.", "Stay on the edge.",
    ]

    def _on_letmecum(self):
        """Roll the dice — grant or deny permission to cum."""
        import random

        # Check denial cooldown
        cooldown_left = getattr(self, '_letmecum_cooldown_until', 0.0) - time.time()
        if cooldown_left > 0:
            self._letmecum_btn.configure(text=f"Wait {int(cooldown_left)}s...")
            return

        aggr = self.aggr_var.get()
        denominator = self._cum_odds.get(aggr, 4)
        granted = random.randint(1, denominator) == 1
        self._last_letmecum_result = "granted" if granted else "denied"
        self._last_letmecum_time = time.time()

        if granted:
            self._cum_allowed = True
            grant_secs = self._CUM_GRANT_TIME.get(aggr, 300)
            self._cum_grant_expires = time.time() + grant_secs
            mins = grant_secs // 60
            self._snark_label.configure(text=f"You've been a good boy. You have {mins} min.",
                                        text_color="#3EC941")
            self._letmecum_btn.configure(text=f"CUM NOW! {mins}:00", fg_color="#3EC941",
                                         hover_color="#32a435")
            if self._cum_override_range:
                target_vol = 1.0
            else:
                target_vol = self.max_vol_var.get() / 100.0
            if self.restim_on.get():
                self.restim.set_volume(target_vol, instant=True, floor=0.0, ceiling=1.0)
            if self.xtoys_on.get():
                self.xtoys.set_volume(target_vol, instant=True, floor=0.0, ceiling=1.0)
            if self.audio_on.get() and self.win_audio and self.win_audio.connected:
                self.win_audio.set_volume(target_vol, 0.0, 1.0)
            if self.mp3_on.get() and self.music_player:
                self.music_player.volume = target_vol
            self.root.after(1000, self._tick_cum_grant)
            log.info(f"Cum GRANTED (1/{denominator} on {aggr}) — {mins} min window")
        else:
            self._cum_allowed = False
            self._denial_count += 1
            cooldown = self._CUM_DENY_COOLDOWN.get(aggr, 30)
            self._letmecum_cooldown_until = time.time() + cooldown
            self._letmecum_btn.configure(text="DENIED!", fg_color="#FF4444",
                                         hover_color="#cc3636")
            snark = random.choice(self._denial_phrases) if self._denial_phrases else "Denied."
            self._snark_label.configure(text=snark)
            self.root.after(2000, lambda: self._tick_letmecum_cooldown())
            log.info(f"Cum DENIED (1/{denominator} on {aggr}) — {cooldown}s cooldown")

    def _tick_letmecum_cooldown(self):
        """Update the button text with remaining cooldown, then restore."""
        remaining = getattr(self, '_letmecum_cooldown_until', 0.0) - time.time()
        if remaining > 0:
            self._letmecum_btn.configure(text=f"Wait {int(remaining)}s...",
                                         fg_color="#FF4444", hover_color="#cc3636")
            self.root.after(1000, self._tick_letmecum_cooldown)
        else:
            self._letmecum_btn.configure(text="Let me cum?",
                                         fg_color="#3EC941", hover_color="#32a435")
            self._snark_label.configure(text="")

    def _tick_cum_grant(self):
        """Countdown the cum grant window. When expired, revoke permission."""
        if not self._cum_allowed:
            return
        remaining = getattr(self, '_cum_grant_expires', 0) - time.time()
        if remaining > 0:
            m = int(remaining) // 60
            s = int(remaining) % 60
            self._letmecum_btn.configure(text=f"CUM NOW! {m}:{s:02d}")
            self._snark_label.configure(
                text=f"You've been a good boy. {m}:{s:02d} remaining.",
                text_color="#3EC941")
            self.root.after(1000, self._tick_cum_grant)
        else:
            # Time's up — revoke permission
            self._cum_allowed = False
            self._last_letmecum_result = "expired"
            self._letmecum_btn.configure(text="Too slow!",
                                         fg_color="#FF4444", hover_color="#cc3636")
            self._snark_label.configure(text="Time's up. Back to edging.",
                                        text_color="#ff4444")
            self.root.after(3000, lambda: (
                self._letmecum_btn.configure(text="Let me cum?",
                                             fg_color="#3EC941", hover_color="#32a435"),
                self._snark_label.configure(text="")))
            log.info("Cum grant expired — permission revoked")

    def _on_cum(self):
        """Hard stop — volume to 0 and pinned there until user presses Resume."""
        self._cum_count += 1
        self._cum_time = time.time()
        self._cum_stopped = True
        self._cum_allowed = False
        self._cum_grant_expires = 0
        self._letmecum_btn.configure(text="Let me cum?", fg_color="#3EC941",
                                     hover_color="#32a435")
        # Disable "Let me cum?" while stopped — no point rolling the dice
        # when the session has been marked finished.
        try:
            self._letmecum_btn.configure(state="disabled")
        except Exception:
            pass
        self._snark_label.configure(text="")
        # Swap "I've CUM" → "Resume" — single button, same slot, no timer games
        self._cum_btn.configure(text="Resume", command=self._on_resume,
                                fg_color=self._C_GREEN,
                                hover_color=self._C_GREEN_H,
                                text_color="white")
        ceil_val = self.max_vol_var.get() / 100.0
        if self.restim_on.get():
            self.restim.set_volume(0.0, instant=True, floor=0.0, ceiling=ceil_val)
        if self.xtoys_on.get():
            self.xtoys.set_volume(0.0, instant=True, floor=0.0, ceiling=ceil_val)
        if self.audio_on.get() and self.win_audio and self.win_audio.connected:
            self.win_audio.set_volume(0.0, 0.0, 1.0)
        if self.mp3_on.get() and self.music_player:
            self.music_player.volume = 0.0
        log.info("Cum triggered — session stopped, awaiting Resume")

    def _on_resume(self):
        """Restart the edging loop after an 'I've CUM' press."""
        self._cum_stopped = False
        self._cum_time = None
        # Swap "Resume" → "I've CUM"
        self._cum_btn.configure(text="I've CUM", command=self._on_cum,
                                fg_color="#e0e0e8",
                                hover_color="#c8c8d0",
                                text_color=self._C_BG)
        # Re-enable "Let me cum?"
        try:
            self._letmecum_btn.configure(state="normal")
        except Exception:
            pass
        log.info("Session resumed after I've CUM")

    def _open_settings(self):
        """Open the settings dialog."""
        win = ctk.CTkToplevel(self.root)
        win.title("Settings")
        win.geometry("480x740")
        win.configure(fg_color=self._C_BG)
        win.transient(self.root)
        win.grab_set()

        lbl = ctk.CTkFont(size=11, weight="bold")

        # ── Appearance (Theme + Font) ─────────────────────────────────────────
        ctk.CTkLabel(win, text="Appearance",
                     font=lbl, text_color=self._C_TEXT).pack(padx=16, pady=(12, 4), anchor="w")
        appear_frame = ctk.CTkFrame(win, fg_color=self._C_SURFACE, corner_radius=8)
        appear_frame.pack(fill=tk.X, padx=16, pady=(0, 8))

        # Theme selector
        ctk.CTkLabel(appear_frame, text="Theme", font=ctk.CTkFont(size=10, weight="bold"),
                     text_color=self._C_TEXT_DIM).pack(padx=12, pady=(8, 2), anchor="w")
        theme_var = tk.StringVar(value=self._theme_name)
        for tname, tcolors in THEMES.items():
            row = ctk.CTkFrame(appear_frame, fg_color="transparent")
            row.pack(fill=tk.X, padx=12, pady=2)
            ctk.CTkRadioButton(row, text=tname, variable=theme_var, value=tname,
                               font=ctk.CTkFont(size=11), text_color=self._C_TEXT,
                               fg_color=self._C_ACCENT, hover_color=self._C_ACCENT_H,
                               border_color=self._C_BORDER, width=100
                               ).pack(side=tk.LEFT)
            swatch_cv = tk.Canvas(row, width=120, height=16, bg=self._C_SURFACE,
                                  highlightthickness=0)
            swatch_cv.pack(side=tk.LEFT, padx=(8, 0))
            preview_colors = [tcolors["BG"], tcolors["SURFACE"], tcolors["ACCENT"],
                              tcolors["RED"], tcolors["GREEN"], tcolors["BLUE"]]
            for i, col in enumerate(preview_colors):
                swatch_cv.create_rectangle(i * 20, 1, i * 20 + 18, 15,
                                           fill=col, outline=tcolors["BORDER"], width=1)

        # Font size
        ctk.CTkFrame(appear_frame, height=1, fg_color=self._C_BORDER
                     ).pack(fill=tk.X, padx=12, pady=6)
        ctk.CTkLabel(appear_frame, text="Font Size", font=ctk.CTkFont(size=10, weight="bold"),
                     text_color=self._C_TEXT_DIM).pack(padx=12, pady=(0, 2), anchor="w")
        font_row = ctk.CTkFrame(appear_frame, fg_color="transparent")
        font_row.pack(fill=tk.X, padx=12, pady=(0, 4))
        font_var = tk.IntVar(value=self._ui_font_size)
        font_slider = ctk.CTkSlider(font_row, from_=8, to=18, number_of_steps=10,
                                     variable=font_var,
                                     button_color=self._C_ACCENT,
                                     button_hover_color=self._C_ACCENT_H,
                                     progress_color=self._C_ACCENT,
                                     fg_color=self._C_SURFACE2, width=180)
        font_slider.pack(side=tk.LEFT, padx=(0, 8))
        font_val_lbl = ctk.CTkLabel(font_row, text=f"{self._ui_font_size}px", font=lbl,
                                     text_color=self._C_ACCENT, width=36)
        font_val_lbl.pack(side=tk.LEFT)

        # Live font preview
        font_preview = ctk.CTkLabel(appear_frame, text="Sample Text Abc 123",
                                     font=ctk.CTkFont(size=self._ui_font_size, weight="bold"),
                                     text_color=self._C_TEXT)
        font_preview.pack(padx=12, pady=(0, 4), anchor="w")

        def _on_font_slide(v):
            sz = int(float(v))
            font_val_lbl.configure(text=f"{sz}px")
            font_preview.configure(font=ctk.CTkFont(size=sz, weight="bold"))
        font_slider.configure(command=_on_font_slide)

        # Apply button
        def _apply_appearance():
            self._theme_name = theme_var.get()
            self._ui_font_size = int(font_var.get())
            self._save_config()
            win.destroy()
            self._on_close()
            import sys, os
            os.execv(sys.executable, [sys.executable] + sys.argv)

        ctk.CTkButton(appear_frame, text="Apply (restarts app)", command=_apply_appearance,
                      font=ctk.CTkFont(size=10, weight="bold"), height=28, corner_radius=4,
                      fg_color=self._C_ACCENT, hover_color=self._C_ACCENT_H,
                      text_color="white").pack(padx=12, pady=(4, 10), anchor="e")

        # ── Calibration ───────────────────────────────────────────────────────
        ctk.CTkLabel(win, text="Calibration",
                     font=lbl, text_color=self._C_TEXT).pack(padx=16, pady=(8, 4), anchor="w")
        calib_frame = ctk.CTkFrame(win, fg_color=self._C_SURFACE, corner_radius=8)
        calib_frame.pack(fill=tk.X, padx=16, pady=(0, 8))

        sens_hdr = ctk.CTkFrame(calib_frame, fg_color="transparent")
        sens_hdr.pack(fill=tk.X, padx=12, pady=(8, 2))
        ctk.CTkLabel(sens_hdr, text="Edging Sensitivity",
                     font=ctk.CTkFont(size=10, weight="bold"),
                     text_color=self._C_TEXT_DIM, anchor="w").pack(side=tk.LEFT)
        sens_val_lbl = ctk.CTkLabel(sens_hdr, text=f"{int(self.edge_sens_var.get())} px",
                                    font=lbl, text_color=self._C_ACCENT, anchor="e")
        sens_val_lbl.pack(side=tk.RIGHT)

        def _on_sens(val):
            sens_val_lbl.configure(text=f"{int(float(val))} px")
            self._save_config()

        ctk.CTkSlider(
            calib_frame, from_=0, to=500, number_of_steps=500,
            variable=self.edge_sens_var, command=_on_sens,
            button_color=self._C_ACCENT, button_hover_color=self._C_ACCENT_H,
            progress_color=self._C_ACCENT, fg_color=self._C_SURFACE2,
        ).pack(fill=tk.X, padx=12, pady=(0, 4))
        ctk.CTkLabel(
            calib_frame,
            text="Pulls the Edging state-display toward Erect by N pixels so "
                 "the OBS overlay lights up earlier. "
                 "Affects only the State label / OBS broadcast — the volume / "
                 "denial math always uses the strict calibrated line.",
            font=ctk.CTkFont(size=9), text_color=self._C_TEXT_DIM,
            wraplength=420, justify="left", anchor="w",
        ).pack(fill=tk.X, padx=12, pady=(0, 10))

        # ── Restim ────────────────────────────────────────────────────────────
        ctk.CTkLabel(win, text="Restim",
                     font=lbl, text_color=self._C_TEXT).pack(padx=16, pady=(8, 4), anchor="w")
        restim_frame = ctk.CTkFrame(win, fg_color=self._C_SURFACE, corner_radius=8)
        restim_frame.pack(fill=tk.X, padx=16, pady=(0, 8))

        ax_row = ctk.CTkFrame(restim_frame, fg_color="transparent")
        ax_row.pack(fill=tk.X, padx=12, pady=(8, 2))
        ctk.CTkLabel(ax_row, text="T-code Axis",
                     font=ctk.CTkFont(size=10, weight="bold"),
                     text_color=self._C_TEXT_DIM, anchor="w").pack(side=tk.LEFT)

        AXIS_PRESETS = ["V0", "V1", "V2", "L0", "L1", "L2",
                        "R0", "R1", "R2", "A0", "A1", "A2"]
        current_axis = self.tcode_axis_var.get()
        if current_axis not in AXIS_PRESETS:
            AXIS_PRESETS.insert(0, current_axis)

        axis_combo = ctk.CTkComboBox(
            ax_row, values=AXIS_PRESETS, variable=self.tcode_axis_var,
            width=90, command=lambda _=None: self._on_tcode_axis_change(),
            fg_color=self._C_SURFACE2, border_color=self._C_BORDER,
            button_color=self._C_ACCENT, button_hover_color=self._C_ACCENT_H,
            dropdown_fg_color=self._C_SURFACE2, text_color=self._C_TEXT,
        )
        axis_combo.pack(side=tk.RIGHT)

        ctk.CTkLabel(
            restim_frame,
            text="Which T-code axis VSE sends volume commands on. This must "
                 "match the axis your Restim session has bound to the "
                 "parameter you want to control (usually 'Volume'). Check "
                 "Restim's Websocket / T-code panel if you're unsure.",
            font=ctk.CTkFont(size=9), text_color=self._C_TEXT_DIM,
            wraplength=420, justify="left", anchor="w",
        ).pack(fill=tk.X, padx=12, pady=(4, 10))

        # ── xToys ─────────────────────────────────────────────────────────────
        ctk.CTkLabel(win, text="xToys",
                     font=lbl, text_color=self._C_TEXT).pack(padx=16, pady=(8, 4), anchor="w")
        xtoys_frame = ctk.CTkFrame(win, fg_color=self._C_SURFACE, corner_radius=8)
        xtoys_frame.pack(fill=tk.X, padx=16, pady=(0, 8))

        xtoys_help_row = ctk.CTkFrame(xtoys_frame, fg_color="transparent")
        xtoys_help_row.pack(fill=tk.X, padx=12, pady=(8, 4))
        ctk.CTkLabel(xtoys_help_row, text="First time setup?",
                     font=ctk.CTkFont(size=9), text_color=self._C_TEXT_DIM).pack(side=tk.LEFT)

        def _show_xtoys_help():
            hw = ctk.CTkToplevel(win)
            hw.title("xToys Setup")
            hw.geometry("420x280")
            hw.resizable(False, False)
            hw.grab_set()
            ctk.CTkLabel(hw,
                text="Setup steps:\n\n"
                     "1. Open xtoys.app in a browser and sign in\n"
                     "2. Scripts \u2192 search \u201cVisualStimEdger\u201d \u2192 Load Script\n"
                     "3. Connections \u2192 add your toy under Generic Output\n"
                     "4. Go to xtoys.app/me \u2192 Private Webhook\n"
                     "5. Copy the Webhook ID \u2192 paste it into VSE\n"
                     "6. In xToys Connections \u2192 enable Private Webhook \u2192 Save\n"
                     "7. Keep the xToys tab open while using VSE",
                font=ctk.CTkFont(size=11), text_color=self._C_TEXT,
                justify="left", anchor="w", wraplength=380,
            ).pack(padx=20, pady=20, anchor="w")
            ctk.CTkButton(hw, text="Close", command=hw.destroy, width=80,
                          fg_color=self._C_SURFACE2, hover_color="#4a4a4a",
                          text_color=self._C_TEXT).pack(pady=(0, 16))

        ctk.CTkButton(xtoys_help_row, text="? Setup Guide", command=_show_xtoys_help,
                      width=100, height=22, font=ctk.CTkFont(size=9),
                      fg_color=self._C_SURFACE2, hover_color="#4a4a4a",
                      text_color=self._C_TEXT_DIM, border_width=1,
                      border_color=self._C_BORDER).pack(side=tk.LEFT, padx=(8, 0))


        # ── Cum Volume Override ───────────────────────────────────────────────
        ctk.CTkLabel(win, text="Cum Volume Behavior",
                     font=lbl, text_color=self._C_TEXT).pack(padx=16, pady=(8, 4), anchor="w")
        cum_vol_frame = ctk.CTkFrame(win, fg_color=self._C_SURFACE, corner_radius=8)
        cum_vol_frame.pack(fill=tk.X, padx=16, pady=(0, 8))
        override_var = tk.BooleanVar(value=self._cum_override_range)
        ctk.CTkRadioButton(cum_vol_frame, text="Override to 100% volume (ignore range)",
                           variable=override_var, value=True,
                           font=ctk.CTkFont(size=11), text_color=self._C_TEXT,
                           fg_color=self._C_ACCENT, hover_color=self._C_ACCENT_H,
                           border_color=self._C_BORDER
                           ).pack(padx=16, pady=(8, 2), anchor="w")
        ctk.CTkRadioButton(cum_vol_frame, text="Respect volume ceiling (stay within range)",
                           variable=override_var, value=False,
                           font=ctk.CTkFont(size=11), text_color=self._C_TEXT,
                           fg_color=self._C_ACCENT, hover_color=self._C_ACCENT_H,
                           border_color=self._C_BORDER
                           ).pack(padx=16, pady=(2, 8), anchor="w")

        # ── Cum Odds ──────────────────────────────────────────────────────────
        ctk.CTkLabel(win, text="\"Let me cum?\" Odds  (1 in N chance)",
                     font=lbl, text_color=self._C_TEXT).pack(padx=16, pady=(12, 4), anchor="w")
        odds_frame = ctk.CTkFrame(win, fg_color=self._C_SURFACE, corner_radius=8)
        odds_frame.pack(fill=tk.X, padx=16, pady=(0, 8))

        odds_vars = {}
        for level in ("Easy", "Middle", "Hard", "Expert"):
            row = ctk.CTkFrame(odds_frame, fg_color="transparent")
            row.pack(fill=tk.X, padx=12, pady=4)
            ctk.CTkLabel(row, text=level, font=lbl, text_color=self._C_TEXT,
                         width=70, anchor="w").pack(side=tk.LEFT)
            ctk.CTkLabel(row, text="1 in", font=ctk.CTkFont(size=10),
                         text_color=self._C_TEXT_DIM).pack(side=tk.LEFT, padx=(0, 4))
            var = tk.IntVar(value=self._cum_odds.get(level, 4))
            odds_vars[level] = var
            ctk.CTkEntry(row, textvariable=var, width=60,
                         fg_color=self._C_SURFACE2, border_color=self._C_BORDER,
                         text_color=self._C_TEXT).pack(side=tk.LEFT)

        # ── Denial Phrases ────────────────────────────────────────────────────
        ctk.CTkLabel(win, text="Denial Phrases  (one per line)",
                     font=lbl, text_color=self._C_TEXT).pack(padx=16, pady=(8, 4), anchor="w")
        phrases_box = ctk.CTkTextbox(win, height=220,
                                     fg_color=self._C_SURFACE, border_color=self._C_BORDER,
                                     text_color=self._C_TEXT, border_width=1,
                                     font=ctk.CTkFont(size=11))
        phrases_box.pack(fill=tk.BOTH, expand=True, padx=16, pady=(0, 8))
        phrases_box.insert("1.0", "\n".join(self._denial_phrases))

        # ── Buttons ───────────────────────────────────────────────────────────
        btn_row = ctk.CTkFrame(win, fg_color="transparent")
        btn_row.pack(fill=tk.X, padx=16, pady=(0, 12))

        def _save():
            for level, var in odds_vars.items():
                val = var.get()
                if val >= 1:
                    self._cum_odds[level] = val
            text = phrases_box.get("1.0", "end").strip()
            self._denial_phrases = [l.strip() for l in text.split("\n") if l.strip()]
            self._cum_override_range = override_var.get()
            self._ui_font_size = int(font_var.get())
            self._theme_name = theme_var.get()
            self._save_config()
            win.destroy()

        def _reset():
            self._cum_odds = dict(self._CUM_ODDS_DEFAULT)
            self._denial_phrases = list(self._DENIAL_PHRASES_DEFAULT)
            for level, var in odds_vars.items():
                var.set(self._CUM_ODDS_DEFAULT.get(level, 4))
            phrases_box.delete("1.0", "end")
            phrases_box.insert("1.0", "\n".join(self._denial_phrases))

        ctk.CTkButton(btn_row, text="Reset Defaults", command=_reset, width=120,
                      fg_color=self._C_SURFACE2, hover_color="#4a4a4a",
                      text_color=self._C_TEXT, border_width=1, border_color=self._C_BORDER,
                      font=ctk.CTkFont(size=10)).pack(side=tk.LEFT)
        ctk.CTkButton(btn_row, text="Save", command=_save, width=100,
                      fg_color=self._C_ACCENT, hover_color=self._C_ACCENT_H,
                      text_color="white", font=ctk.CTkFont(size=12, weight="bold")
                      ).pack(side=tk.RIGHT)

    def _on_floor_change(self, val):
        lo = int(self.min_vol_var.get())
        hi = int(self.max_vol_var.get())
        self._range_lbl.configure(text=f"{lo}% – {hi}%")
        if hasattr(self, '_range_draw'):
            self._range_draw()
        self._save_config()

    def _on_ceil_change(self, val):
        lo = int(self.min_vol_var.get())
        hi = int(self.max_vol_var.get())
        self._range_lbl.configure(text=f"{lo}% – {hi}%")
        if hasattr(self, '_range_draw'):
            self._range_draw()
        self._save_config()

    def _on_output_change(self):
        P = 12
        # Hide all option panels then re-pack in order for active outputs
        self._restim_opts.pack_forget()
        self._xtoys_opts.pack_forget()
        self._hr_opts.pack_forget()
        self._windows_opts.pack_forget()
        self._mp3_opts.pack_forget()

        after = self._mode_card
        if self.restim_on.get():
            self._restim_opts.pack(fill=tk.X, padx=P, pady=(0, 4), after=after)
            after = self._restim_opts
        if self.xtoys_on.get():
            self._xtoys_opts.pack(fill=tk.X, padx=P, pady=(0, 4), after=after)
            after = self._xtoys_opts
        if self.hr_on.get():
            self._hr_opts.pack(fill=tk.X, padx=P, pady=(0, 4), after=after)
            after = self._hr_opts
            self.hr_client.resting_bpm = self.hr_resting_var.get()
            self.hr_client.peak_bpm    = self.hr_peak_var.get()
            self.hr_client.token       = self.hr_token_var.get().strip()
            self.hr_client.start()
        else:
            self.hr_client.stop()
        if self.audio_on.get():
            self._windows_opts.pack(fill=tk.X, padx=P, pady=(0, 4), after=after)
            after = self._windows_opts
            if not self.win_devices:
                self._refresh_devices()
        if self.mp3_on.get() and _MINIAUDIO_OK:
            self._mp3_opts.pack(fill=tk.X, padx=P, pady=(0, 4), after=after)
        self._save_config()

    def _on_tcode_axis_change(self, *_):
        """Swap the axis VSE sends volume on. No reconnect needed — the next
        send will use the new prefix."""
        val = (self.tcode_axis_var.get() or "").strip().upper()
        if len(val) < 2 or not val[0].isalpha() or not val[1:].isdigit():
            return
        if val != self.restim.axis:
            self.restim.axis = val
            log.info(f"Restim: T-code axis switched to {val}")
            # Push the current volume on the new axis so Restim can pick it up
            # without waiting for the next tick.
            try:
                self.restim.set_volume(self.restim.volume, instant=True)
            except Exception:
                pass
        self._save_config()

    def _on_xtoys_id_change(self, *_):
        new_id = self.xtoys_id_var.get().strip()
        if new_id != self.xtoys.webhook_id:
            self.xtoys.disconnect()
            self.xtoys.webhook_id = new_id
            log.info(f"xToys: webhook ID={new_id!r}")
        self._save_config()

    def _on_hr_token_change(self, *_):
        new_token = self.hr_token_var.get().strip()
        if new_token != self.hr_client.token and self.hr_on.get():
            self.hr_client.restart(token=new_token)
        elif new_token != self.hr_client.token:
            self.hr_client.token = new_token
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
                    v = self.win_audio.get_volume()
                    self._orig_win_volume = v if v is not None else 0.5
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
                                     fg_color=self._C_SURFACE2, hover_color="#4a4a4a",
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

            # OBS overlay broadcast — 4 Hz
            if now - self._last_overlay_broadcast >= 0.25:
                self._last_overlay_broadcast = now
                self._broadcast_overlay(state)

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
            interval = 100 if self.tracking_paused else 33  # ~30fps poll
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

    def _edging_y(self):
        """Effective Edging line Y, with sensitivity offset baked in.

        head_y is bbox top-Y — smaller = more erect. To make Edging trip
        earlier (on less firm position) we nudge the line DOWN in frame
        space (larger Y) toward Erect. Clamped so we can never cross Erect."""
        base = self.heights.get("Edging")
        if base is None:
            return None
        try:
            offset = int(self.edge_sens_var.get())
        except Exception:
            offset = 0
        eff = base + max(0, offset)
        erect = self.heights.get("Erect")
        if erect is not None:
            # Never go past Erect (need a 1-pixel gap for the math to survive).
            eff = min(eff, erect - 1)
        return eff

    def _determine_state(self, y_pos):
        if any(self.heights.get(k) is None for k in ("Edging", "Erect", "Flaccid")):
            return "Erect (Needs Calibration)"
        edging_y = self._edging_y()
        dist_edging  = abs(y_pos - edging_y)
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

        if state == "Edging":
            if self._prev_state != "Edging":
                self._edge_enter_time = now
            elif (now - getattr(self, '_edge_enter_time', now) >= 1.0
                  and not getattr(self, '_edge_counted', False)
                  and now - getattr(self, '_last_edge_time', 0.0) >= 10.0):
                self.edge_count += 1
                self._last_edge_time = now
                self._edge_counted = True
        else:
            self._edge_counted = False

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
        if self.heights["Edging"]  is not None:
            # Solid line at the calibrated Edging height, plus a dashed line
            # at the effective (sensitivity-offset) position if it differs.
            base = int(self.heights["Edging"])
            cv2.line(frame, (0, base), (fw, base), (0, 0, 255), 2)
            eff = self._edging_y()
            if eff is not None and eff != base:
                eff = int(eff)
                # Dashed orange line at the effective (sensitivity-offset) trigger position.
                dash = 14
                for x in range(0, fw, dash * 2):
                    cv2.line(frame, (x, eff), (min(x + dash, fw), eff), (0, 165, 255), 2)
                cv2.putText(frame, "SENS", (4, eff - 4),
                            cv2.FONT_HERSHEY_SIMPLEX, 0.45, (0, 165, 255), 1, cv2.LINE_AA)
        if self.heights["Erect"]   is not None: cv2.line(frame, (0, int(self.heights["Erect"])),   (fw, int(self.heights["Erect"])),   (0, 255, 0), 2)
        if self.heights["Flaccid"] is not None: cv2.line(frame, (0, int(self.heights["Flaccid"])), (fw, int(self.heights["Flaccid"])), (255, 0, 0), 2)

    def _compute_volume_delta(self):
        if any(self.heights.get(k) is None for k in ("Edging", "Erect", "Flaccid")):
            return 0.0
        # Volume math uses the RAW calibrated Edging line, not the sensitivity-
        # adjusted one. The sensitivity slider is purely an OBS / state-display
        # knob so the overlay can light up earlier without also making the
        # denial curve hair-trigger.
        edging_y = self.heights["Edging"]
        full_range = self.heights["Flaccid"] - edging_y
        if abs(full_range) < 1:
            return 0.0

        history = self._head_y_history
        smoothed_y = sum(history) / len(history) if history else self.head_y

        position   = (smoothed_y             - edging_y) / full_range
        erect_norm = (self.heights["Erect"]   - edging_y) / full_range
        aggr_level = self.aggr_var.get()
        aggr_mult  = AGGR_LEVELS.get(aggr_level, 1.0)

        # Velocity: normalised rate of change across the history window.
        # Positive = moving toward flaccid, negative = moving toward edging.
        if len(history) >= 4:
            velocity = (history[-1] - history[0]) / (len(history) * abs(full_range))
        else:
            velocity = 0.0

        if 0.0 <= position <= erect_norm:
            # Sweet zone: drift flaccid → nudge volume up;
            # climb fast toward edge → pre-emptive denial, but only once the
            # session's edge count has crossed the per-level unlock threshold.
            # Early edges are "free" (pure reactive behavior from the branches
            # below); later edges trigger predictive denial inside the ~1–10s
            # PONR window documented in the Restim edging literature. Strength
            # then escalates mildly per edge past the threshold, so the tool
            # tightens progressively across the session.
            if velocity >= 0.0:
                vel_nudge = velocity * 0.4
                return VOLUME_STEP * vel_nudge * aggr_mult
            unlock = PREEMPT_UNLOCK.get(aggr_level)
            if unlock is not None and self.edge_count >= unlock:
                # 1.0 at the unlock point, +0.1 per edge past it, caps at 2.0
                escalation     = min(2.0, 1.0 + 0.1 * (self.edge_count - unlock))
                edge_proximity = 1.0 - (position / max(erect_norm, 1e-3))  # 0=at erect, 1=at edge
                vel_preempt    = (-velocity) * (0.3 + 0.5 * edge_proximity) * escalation
                return -VOLUME_STEP * vel_preempt * aggr_mult
            return 0.0

        if position < 0.0:
            # Past edging — ease off; dampen further if still moving toward edging
            vel_damp = max(0.0, -velocity) * 0.3
            return -VOLUME_STEP * (0.5 + vel_damp) * aggr_mult

        # Past erect, drifting toward flaccid
        dist      = (position - erect_norm) / max(1.0 - erect_norm, 0.01)
        vel_boost = max(0.0, velocity) * 0.5   # moving fast toward flaccid = respond harder
        return VOLUME_STEP * (min(dist, 1.0) + vel_boost) * aggr_mult


    def _tick_volume(self):
        cur_time = time.time()
        if cur_time - self.last_vol_time < VOLUME_UPDATE_INTERVAL:
            return
        self.last_vol_time = cur_time

        # ── Cum allowed — volume locked at 100% until "I've CUM" ─────────────
        if getattr(self, '_cum_allowed', False):
            return

        # ── Session stopped after "I've CUM" — volume pinned at 0 ───────────
        # (_on_cum set the volume to 0 already; just keep it there every tick
        # in case something else nudged it.)
        if self._cum_stopped:
            ceil_val = self.max_vol_var.get() / 100.0
            if self.restim_on.get():
                self.restim.set_volume(0.0, floor=0.0, ceiling=ceil_val)
            if self.xtoys_on.get():
                self.xtoys.set_volume(0.0, floor=0.0, ceiling=ceil_val)
            if self.audio_on.get() and self.win_audio and self.win_audio.connected:
                self.win_audio.set_volume(0.0, 0.0, 1.0)
            if self.mp3_on.get() and self.music_player:
                self.music_player.volume = 0.0
            return

        if self.hold_active:
            return  # Volume frozen by user

        history    = self._head_y_history
        smoothed_y = sum(history) / len(history) if history else self.head_y


        floor_val = min(self.min_vol_var.get(), self.max_vol_var.get()) / 100.0
        ceil_val  = self.max_vol_var.get() / 100.0
        delta     = self._compute_volume_delta()

        # ── HR modifier: tighten denial / slow rewards when heart rate is high ─
        if self.hr_on.get() and self.hr_client.connected:
            hr_mod = self.hr_client.modifier()   # 1.0 (resting) → 2.0 (peak)
            if delta < 0.0:
                delta *= hr_mod                  # harder denial
            elif delta > 0.0:
                delta *= max(0.4, 2.0 - hr_mod)  # stingier reward (halved at peak)

        if self.restim_on.get():
            if delta != 0.0:
                self.restim.adjust_volume(delta, floor=floor_val, ceiling=ceil_val)
            self.restim.maybe_reconnect()
        if self.xtoys_on.get():
            if delta != 0.0:
                self.xtoys.adjust_volume(delta, floor=floor_val, ceiling=ceil_val)
            self.xtoys.maybe_reconnect()
        if self.audio_on.get() and self.win_audio and self.win_audio.connected and delta != 0.0:
            self.win_audio.adjust_volume(delta, floor=floor_val, ceiling=ceil_val)
        if self.mp3_on.get() and self.music_player and delta != 0.0:
            self.music_player.adjust_volume(delta, floor=floor_val, ceiling=ceil_val)

    def _update_status_label(self, state):
        quality_str = "Track: OK" if self.tracking_ok else "Track: LOST"
        parts = []
        conn_color = "#ffaa00"

        if self.restim_on.get():
            ok = bool(self.restim.ws)
            if ok: conn_color = "#00ff00"
            elif conn_color != "#00ff00": conn_color = "#ff0000"
            parts.append(f"WS: {'OK' if ok else 'Disconnected'}")
        if self.xtoys_on.get():
            ok = self.xtoys.connected
            if ok: conn_color = "#00ff00"
            elif conn_color != "#00ff00":
                conn_color = "#ffaa00" if not self.xtoys.enabled else "#ff0000"
            parts.append(f"xToys: {'OK' if ok else ('No ID' if not self.xtoys.enabled else 'Connecting...')}")
        if self.audio_on.get():
            if self.win_audio and self.win_audio.connected:
                conn_color = "#00ff00"
                parts.append("Audio: OK")
            else:
                parts.append("Audio: No Device")
        if self.mp3_on.get() and self.music_player:
            self._mp3_update_track_label()
            if self.music_player._state == "playing": conn_color = "#00ff00"
            parts.append(f"MP3: {self.music_player.track_name or '—'}")
        if self.hr_on.get():
            sbpm = self.hr_client.smooth_bpm()
            if self.hr_client.connected and sbpm is not None:
                parts.append(f"\u2665 {sbpm:.0f} bpm")
                if hasattr(self, '_hr_bpm_label'):
                    self._hr_bpm_label.configure(
                        text=f"\u2665 {sbpm:.0f} bpm",
                        text_color="#e91e63")
            else:
                parts.append("\u2665 No HR")
                if hasattr(self, '_hr_bpm_label'):
                    self._hr_bpm_label.configure(
                        text="\u2665 -- bpm",
                        text_color=self._C_TEXT_DIM)

        src_str = " | ".join(parts) if parts else "No output"

        # Volume — highest active output
        _vols = []
        if self.restim_on.get():  _vols.append(self.restim.volume)
        if self.xtoys_on.get():   _vols.append(self.xtoys.volume)
        if self.audio_on.get() and self.win_audio and self.win_audio.connected:
            _vols.append(self.win_audio.get_volume() or 0)
        if self.mp3_on.get() and self.music_player:
            _vols.append(self.music_player.volume)
        vol_str = f"{max(_vols) * 100:.0f}%" if _vols else "--"

        ft = self._frame_times
        fps = (len(ft) - 1) / max(ft[-1] - ft[0], 1e-9) if len(ft) >= 2 else 0.0
        yolo_str = (f"YOLO: {self.detector.last_conf:.0%}"
                    if self.detector.available and self.detector.last_conf > 0 else "YOLO: --")

        self.info_label.configure(
            text=(f"State: {state}  |  Vol: {vol_str}  |  {quality_str}"
                  f"  |  {src_str}  |  {fps:.0f} fps  |  {yolo_str}"),
            text_color=conn_color,
        )

    def _broadcast_overlay(self, state):
        # Use highest active volume so OBS reflects what's actually happening
        vols = []
        if self.restim_on.get():  vols.append(self.restim.volume)
        if self.xtoys_on.get():   vols.append(self.xtoys.volume)
        if self.audio_on.get() and self.win_audio and self.win_audio.connected:
            vols.append(self.win_audio.get_volume() or 0)
        if self.mp3_on.get() and self.music_player:
            vols.append(self.music_player.volume)
        vol = max(vols) if vols else 0

        # "Let me cum?" tattle-tale
        letmecum = None
        result = getattr(self, '_last_letmecum_result', None)
        if result:
            elapsed_lmc = time.time() - getattr(self, '_last_letmecum_time', 0)
            if result == "granted" and self._cum_allowed:
                grant_left = getattr(self, '_cum_grant_expires', 0) - time.time()
                letmecum = {"result": "granted", "time_left": round(max(grant_left, 0))}
            elif result == "expired" and elapsed_lmc < 5:
                letmecum = {"result": "expired"}
            elif result == "denied":
                cooldown_left = getattr(self, '_letmecum_cooldown_until', 0) - time.time()
                if cooldown_left > 0:
                    letmecum = {"result": "denied", "retry_in": round(cooldown_left)}
                elif elapsed_lmc < 5:
                    letmecum = {"result": "denied", "retry_in": 0}

        # Heart rate data for overlay
        hr_data = None
        if self.hr_on.get():
            sbpm = self.hr_client.smooth_bpm()
            hr_data = {
                "bpm": round(sbpm) if sbpm is not None else None,
                "connected": self.hr_client.connected,
                "modifier": round(self.hr_client.modifier(), 2),
            }

        obs_state = "-" if "Calibration" in state else state
        self._overlay.broadcast(json.dumps({
            "state": obs_state,
            "volume": round(vol, 3),
            "edge_count": self.edge_count,
            "session_seconds": round(time.time() - self.session_start),
            "cum_stopped": self._cum_stopped,
            "aggressiveness": self.aggr_var.get(),
            "letmecum": letmecum,
            "cum_count": self._cum_count,
            "denial_count": self._denial_count,
            "hr": hr_data,
        }))

    def _display_frame(self, frame):
        img = Image.fromarray(cv2.cvtColor(frame, cv2.COLOR_BGR2RGB))
        # Fit image to container, preserving aspect ratio
        iw, ih = img.size
        cw = self.video_label.winfo_width()
        ch = self.video_label.winfo_height()
        scale = 1.0
        nw, nh = iw, ih
        if cw > 1 and ch > 1:
            scale = min(cw / iw, ch / ih)
            nw, nh = int(iw * scale), int(ih * scale)
            if nw > 0 and nh > 0:
                img = img.resize((nw, nh), Image.LANCZOS)
        # Record the transform so clicks on the label can be mapped back to
        # source-frame coordinates (tk.Label centers the image → letterbox offsets).
        self._disp_scale    = scale
        self._disp_offset_x = max(0, (cw - nw) // 2) if cw > 1 else 0
        self._disp_offset_y = max(0, (ch - nh) // 2) if ch > 1 else 0
        self._disp_frame_h  = ih
        self._disp_frame_w  = iw
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

            # Stamp version number in the gap below the subtitle
            draw = ImageDraw.Draw(img)
            try:
                f_ver = ImageFont.truetype("C:/Windows/Fonts/segoeui.ttf", 14)
            except Exception:
                f_ver = ImageFont.load_default()
            ver_txt = f"v{VERSION}"
            bb = draw.textbbox((0, 0), ver_txt, font=f_ver)
            tx = (W - (bb[2] - bb[0])) // 2
            draw.text((tx, 160), ver_txt, fill=(100, 100, 100, 255), font=f_ver)

            # Flatten RGBA onto the dark background colour (#0d0d0d = 13,13,13)
            bg_flat = Image.new("RGB", (W, H), (13, 13, 13))
            bg_flat.paste(img, mask=img.split()[3])

            photo = ImageTk.PhotoImage(bg_flat)

            lbl = tk.Label(root, image=photo, bg="#0d0d0d",
                           borderwidth=0, highlightthickness=0)
            lbl.image = photo  # keep reference
            lbl.pack()

            # Click-zone over the baked-in button in the PNG.
            # Fractional positions from splash.png pixel scan.
            _btn_y1 = int(0.873 * H)   # 511/585
            _btn_y2 = int(0.971 * H)   # 568/585
            _btn_x1 = int(0.059 * W)   # 35/591
            _btn_x2 = int(0.941 * W)   # 556/591

            def _in_btn(e):
                return _btn_x1 <= e.x <= _btn_x2 and _btn_y1 <= e.y <= _btn_y2

            lbl.bind("<Button-1>", lambda e: _start() if _in_btn(e) else None)
            lbl.configure(cursor="arrow")

            # Blink: alternate between normal image and a bright-flash
            # version to draw attention to the Start button.
            flash = Image.new("RGBA", img.size, (0, 0, 0, 0))
            flash_draw = ImageDraw.Draw(flash)
            flash_draw.rounded_rectangle(
                [_btn_x1, _btn_y1, _btn_x2, _btn_y2],
                radius=int(12 * W / 700), fill=(255, 255, 255, 60))
            img_bright = Image.alpha_composite(img, flash)
            bg_bright = Image.new("RGB", (W, H), (13, 13, 13))
            bg_bright.paste(img_bright, mask=img_bright.split()[3])
            photo_bright = ImageTk.PhotoImage(bg_bright)

            _blink_id = [None]
            blink_state = [False]
            def _blink():
                blink_state[0] = not blink_state[0]
                lbl.configure(image=photo_bright if blink_state[0] else photo)
                _blink_id[0] = root.after(700, _blink)
            _blink_id[0] = root.after(700, _blink)

            # Hand cursor only when entering/leaving the button area
            _cursor = ["arrow"]
            def _on_motion(e):
                want = "hand2" if _in_btn(e) else "arrow"
                if want != _cursor[0]:
                    _cursor[0] = want
                    lbl.configure(cursor=want)
            lbl.bind("<Motion>", _on_motion)

            def _close():
                if _blink_id[0]:
                    root.after_cancel(_blink_id[0])
                root.destroy()
            root.protocol("WM_DELETE_WINDOW", _close)
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
    RED      = "#FF4444"
    YELLOW   = "#ffcc00"
    TEXT     = "#eeeeee"
    DIM      = "#666666"
    SUBDIM   = "#888888"
    P        = 24

    f_title  = tkfont.Font(family="Segoe UI", size=32, weight="bold")
    f_sub    = tkfont.Font(family="Segoe UI", size=14)
    f_head   = tkfont.Font(family="Segoe UI", size=16, weight="bold")
    f_body   = tkfont.Font(family="Segoe UI", size=13)
    f_num    = tkfont.Font(family="Segoe UI", size=13, weight="bold")
    f_btn    = tkfont.Font(family="Segoe UI", size=17, weight="bold")

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
                 fg=TEXT, anchor="w", wraplength=420).pack(anchor="w", padx=(8, 0))
        tk.Label(col, text=body, font=f_body, bg=CARD,
                 fg=SUBDIM, anchor="w", justify="left", wraplength=420).pack(anchor="w", padx=(8, 0))

    tk.Label(root,
             text="Controls volume only — does not generate e-stim signals.\n"
                  "You need Restim, xToys, electron-redrive, an .mp3, etc. already running.",
             font=f_body, bg=BG, fg=DIM, justify="center", wraplength=500).pack(pady=(4, 12))

    btn = tk.Button(root, text="I'm ready — select my camera feed  \u2192",
                    font=f_btn, bg=RED, fg="white", activebackground="#cc3636",
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


_SINGLE_INSTANCE_MUTEX = None  # held for process lifetime


def _acquire_single_instance() -> bool:
    """Create a named Windows mutex so only one VSE can run at once.
    Returns True if we got the lock, False if another instance owns it."""
    global _SINGLE_INSTANCE_MUTEX
    try:
        import ctypes
        from ctypes import wintypes
        ERROR_ALREADY_EXISTS = 183
        CreateMutexW = ctypes.windll.kernel32.CreateMutexW
        CreateMutexW.argtypes = [wintypes.LPVOID, wintypes.BOOL, wintypes.LPCWSTR]
        CreateMutexW.restype = wintypes.HANDLE
        handle = CreateMutexW(None, False, "Global\\VisualStimEdgerSingleton")
        if ctypes.windll.kernel32.GetLastError() == ERROR_ALREADY_EXISTS:
            return False
        _SINGLE_INSTANCE_MUTEX = handle  # keep alive for process lifetime
        return True
    except Exception as e:
        log.warning(f"Single-instance check failed, allowing launch: {e}")
        return True


def main():
    parser = argparse.ArgumentParser(description="VisualStimEdger")
    parser.add_argument("--debug", action="store_true",
                        help="Enable verbose debug logging to console")
    args = parser.parse_args()

    # Always log to a file, because the windowed exe has no console.
    log_path = CONFIG_PATH.parent / "vse.log"
    try:
        log_path.parent.mkdir(parents=True, exist_ok=True)
    except Exception:
        pass
    handlers = [logging.StreamHandler()]
    try:
        fh = logging.FileHandler(log_path, mode="w", encoding="utf-8")
        handlers.append(fh)
    except Exception:
        pass
    logging.basicConfig(
        level=logging.DEBUG if args.debug else logging.INFO,
        format="%(asctime)s %(levelname)-8s %(name)s: %(message)s",
        datefmt="%H:%M:%S",
        handlers=handlers,
        force=True,
    )

    log.info(f"VisualStimEdger v{VERSION} starting")
    log.info(f"Log file: {log_path}")

    if not _acquire_single_instance():
        log.warning("Another VisualStimEdger instance is already running — exiting")
        try:
            import ctypes
            ctypes.windll.user32.MessageBoxW(
                None,
                "VisualStimEdger is already running.\n\nClose the existing window first.",
                "VisualStimEdger",
                0x40,  # MB_ICONINFORMATION
            )
        except Exception:
            pass
        return

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
