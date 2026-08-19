from __future__ import annotations  # lazy annotations: PEP-604 (X | None) unions
                                    # must not evaluate on Python 3.9 (default on
                                    # many Macs) — see SessionLogger._log crash.

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
# OpenCV shares one internal parallel_for thread pool across the whole process.
# VSE calls cv2 from two threads at once — cv2.cvtColor on the background capture
# thread (capture_window_region) and cv2.dnn/TrackerCSRT on the main thread. That
# concurrency corrupts the shared pool's internals and crashes with an access
# violation in MSVCP140.dll (no Python traceback). Disabling OpenCV's internal
# threading makes every cv2 call run synchronously on its calling thread with no
# shared mutable pool — eliminating the race. Negligible perf cost at our frame sizes.
try:
    cv2.setNumThreads(0)
except Exception:
    pass
import numpy as np
import time
import threading

# All OpenCV calls funnel through this lock. cv2 shares process-global state (memory
# allocator, error state) that is NOT safe for concurrent calls from different threads.
# VSE calls cv2 from the background capture thread (cv2.cvtColor in capture_window_region)
# AND the main thread (dnn/TrackerCSRT/drawing) at ~30fps each — the overlap corrupts the
# shared allocator and crashes with an access violation in MSVCP140.dll (no Python trace).
# Holding this lock around every cv2 call makes capture and processing mutually exclusive.
_CV2_LOCK = threading.Lock()
import webbrowser
import requests
import ssl
import websocket
import random
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
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
import datetime
import pathlib
import queue
from collections import deque
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
from comtypes import CLSCTX_ALL

import atexit
import gc

log = logging.getLogger("VisualStimEdger")

_sct_tls = threading.local()

def _get_sct():
    """Per-THREAD mss instance.

    mss is NOT thread-safe — it stashes its GDI bitmap/DC handles in thread-local
    storage. A single shared instance used from both the capture thread (the mss
    fallback in capture_window_region, hit every frame for GPU-composited source
    windows where PrintWindow returns black) and the main thread (region select,
    re-select, grid) corrupts those handles cross-thread, which can hard-kill the
    process below Python — no traceback, no WER. Giving each thread its own mss
    keeps the handles thread-confined.
    """
    sct = getattr(_sct_tls, "sct", None)
    if sct is None:
        sct = mss()
        _sct_tls.sct = sct
        atexit.register(_close_sct, sct)
    return sct

def _close_sct(sct):
    try:
        sct.close()
    except Exception:
        pass


def resource_path(relative):
    """Resolve path to bundled resource — works both in dev and PyInstaller .exe."""
    base = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, relative)


# ── localisation ──────────────────────────────────────────────────────────────
_CATALOG: dict = {}   # active-language translations: English source string -> target


def tr(s: str) -> str:
    """Translate a UI string via the active-language catalog. Falls back to the
    English key on any miss, so partial coverage renders English, never blank."""
    return _CATALOG.get(s, s)


def _load_catalog(lang: str) -> None:
    """Load i18n/vse_ui_strings_<lang>.tsv into the active catalog (en -> target).
    lang 'en' or a missing/unreadable file leaves the catalog empty = English UI.
    The TSV is two tab-separated columns; a literal \\n in a cell becomes a real
    newline so keys match the multi-line Python source strings."""
    _CATALOG.clear()
    if not lang or lang == "en":
        return
    path = resource_path(os.path.join("i18n", f"vse_ui_strings_{lang}.tsv"))
    try:
        with open(path, encoding="utf-8") as fh:
            for i, line in enumerate(fh):
                if i == 0:                        # header row (en / <lang>)
                    continue
                line = line.rstrip("\r\n")
                if "\t" not in line:
                    continue
                en, tgt = line.split("\t", 1)
                en, tgt = en.replace("\\n", "\n"), tgt.replace("\\n", "\n")
                if en and tgt:
                    _CATALOG[en] = tgt
        log.info(f"i18n: loaded {len(_CATALOG)} strings for '{lang}'")
    except FileNotFoundError:
        log.warning(f"i18n: no catalog for '{lang}' — staying English")
    except Exception as e:
        log.warning(f"i18n: failed to load '{lang}' ({e}) — staying English")


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
        Three-tier confidence cascade (single forward pass):
          Tier 1 — class 1 (head) conf >= 0.35                  anywhere in frame
          Tier 2 — class 1 (head) conf >= 0.20                  lower 60% of frame only
          Tier 3 — class 0 (shaft) conf >= 0.25                 infer head at top of shaft bbox
        Returns (x, y, w, h) in frame pixels or None.
        """
        if not self.available:
            return None

        fh, fw = frame.shape[:2]
        blob = cv2.dnn.blobFromImage(frame, 1 / 255.0, self.INPUT_SIZE,
                                     swapRB=True, crop=False)
        self._net.setInput(blob)
        outputs = self._net.forward(self._output_layers)

        head_boxes, head_confs   = [], []
        shaft_boxes, shaft_confs = [], []
        for output in outputs:
            for det in output:
                scores   = det[5:]
                class_id = int(np.argmax(scores))
                conf     = float(scores[class_id])
                cx = int(det[0] * fw);  cy = int(det[1] * fh)
                bw = int(det[2] * fw);  bh = int(det[3] * fh)
                box = [cx - bw // 2, cy - bh // 2, bw, bh]
                if class_id == 1:
                    head_boxes.append(box);  head_confs.append(conf)
                elif class_id == 0:
                    shaft_boxes.append(box); shaft_confs.append(conf)

        def _best_nms(boxes, confs, min_conf):
            filt = [(b, c) for b, c in zip(boxes, confs) if c >= min_conf]
            if not filt:
                return None, 0.0
            fb, fc = zip(*filt)
            idxs = cv2.dnn.NMSBoxes(list(fb), list(fc), min_conf, self.nms_threshold)
            if len(idxs) == 0:
                return None, 0.0
            best = max(idxs.flatten(), key=lambda i: fc[i])
            return tuple(fb[best]), fc[best]

        # Tier 1: head class, high confidence, full frame
        bbox, conf = _best_nms(head_boxes, head_confs, 0.35)
        if bbox:
            self.last_conf = conf
            return bbox

        # Tier 2: head class, lower confidence, lower 60% of frame only
        lower_idx   = [i for i, b in enumerate(head_boxes)
                       if b[1] + b[3] // 2 >= fh * 0.40]
        lower_boxes = [head_boxes[i] for i in lower_idx]
        lower_confs = [head_confs[i] for i in lower_idx]
        bbox, conf  = _best_nms(lower_boxes, lower_confs, 0.20)
        if bbox:
            self.last_conf = conf
            return bbox

        # Tier 3: shaft class fallback — infer head position at top of shaft bbox
        shaft_bbox, conf = _best_nms(shaft_boxes, shaft_confs, 0.25)
        if shaft_bbox:
            sx, sy, sw, sh = shaft_bbox
            head_h = max(20, sh // 4)
            self.last_conf = conf
            return (sx, sy, sw, head_h)

        return None

# --- CONFIGURATION ---
VERSION = "1.9.4"
GITHUB_REPO = "blucrew/VisualStimEdger"
RESTIM_HOST = '127.0.0.1'
RESTIM_PORT = 12346
TCODE_AXIS = 'V0'
VOLUME_STEP = 0.05
VOLUME_UPDATE_INTERVAL = 0.5

# Default main-window client width on first boot (before the user resizes it).
# The visible content fits ~656px wide; reqwidth over-reports because a widget
# below over-requests, which used to boot the window too wide. See _fit_window.
_DEFAULT_WIN_W = 660
# Hard cap on the remembered window width. Content only needs ~560px across, so a
# stored width wider than this is the old over-request artifact (or dead empty
# space) — refuse to load or persist it, and fall back to the skinny default.
_MAX_WIN_W = 900

# Erect-zone recovery: without this the sweet-zone only nudges volume UP in
# proportion to how fast you're drifting toward flaccid, so holding steady near
# the erect line parks the volume forever. This is a flat, gentle per-tick climb
# (≈1.6%/s at a 0.5s tick) scaled by distance from the edge — 0 at the edge line,
# full at the erect line — so it recovers when you hold off the edge but never
# fights you while you're actually edging.
# Per-difficulty recovery/ramp tuning. "Relentless" scaling: harder difficulty
# works you harder — faster recovery, shorter holds, faster ramp = more edges
# per minute. Recovery and ramp are in %/second (recovery measured at the erect
# line, tapering to 0 at the edge); hold is in seconds.
RECOVERY_PCT_DEFAULT = {"Easy": 0.8, "Middle": 1.6, "Hard": 3.0,  "Expert": 5.0}
AE_HOLD_SECS_DEFAULT = {"Easy": 25,  "Middle": 15,  "Hard": 8,    "Expert": 4}
AE_RAMP_PCT_DEFAULT  = {"Easy": 4.0, "Middle": 8.0, "Hard": 14.0, "Expert": 22.0}


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
CUM_DETECT_MAXLEN  = 300  # ~10 s of head-Y samples at 30 fps for cum detection

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
        self.label = tk.Label(self.root, text=tr("Step 1 — Draw a box around your video feed, then release to lock.  [Esc to cancel]"), font=("Arial", 18, "bold"), bg="#111111", fg="#F5A623")
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
        if self.start_x is None:
            return
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
        hwndDC = saveDC = mfcDC = saveBitMap = oldBitMap = None
        frame = None
        try:
            hwndDC    = win32gui.GetWindowDC(hwnd)
            mfcDC     = win32ui.CreateDCFromHandle(hwndDC)
            saveDC    = mfcDC.CreateCompatibleDC()
            saveBitMap = win32ui.CreateBitmap()
            saveBitMap.CreateCompatibleBitmap(mfcDC, w, h)
            # Keep the DC's original bitmap so we can deselect saveBitMap before
            # deleting it — DeleteObject on a STILL-SELECTED bitmap silently fails
            # and leaks the GDI handle. At 30 fps that exhausts the per-process GDI
            # limit in minutes, after which DWM throws wil::ResultException storms
            # and the process dies. (Prior "fix" 9b2dc74 missed the deselect.)
            oldBitMap = saveDC.SelectObject(saveBitMap)

            result = ctypes.windll.user32.PrintWindow(hwnd, saveDC.GetSafeHdc(), 2)
            if result == 1:
                bmpinfo = saveBitMap.GetInfo()
                bmpstr  = saveBitMap.GetBitmapBits(True)
                img     = np.frombuffer(bmpstr, dtype=np.uint8).reshape(
                              (bmpinfo['bmHeight'], bmpinfo['bmWidth'], 4))
                with _CV2_LOCK:
                    frame = cv2.cvtColor(img, cv2.COLOR_BGRA2BGR)
        except Exception as e:
            log.debug(f"PrintWindow failed: {e}")
            frame = None
        finally:
            # Deselect saveBitMap (restore original) BEFORE deleting it, then tear
            # down in reverse order of creation so every GDI object actually frees.
            if saveDC and oldBitMap:
                try: saveDC.SelectObject(oldBitMap)
                except Exception: pass
            if saveBitMap:
                try: win32gui.DeleteObject(saveBitMap.GetHandle())
                except Exception: pass
            if saveDC:
                try: saveDC.DeleteDC()
                except Exception: pass
            if mfcDC:
                try: mfcDC.DeleteDC()
                except Exception: pass
            if hwndDC:
                try: win32gui.ReleaseDC(hwnd, hwndDC)
                except Exception: pass

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
        with _CV2_LOCK:
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
    
    root.title(tr("Step 2: Select Cock Head"))
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

    btn_frame = tk.Frame(root, bg="#111111")
    btn_frame.pack(fill=tk.X)
    tk.Button(btn_frame, text=tr("Confirm Head Area ✅"), command=root.destroy, font=("Arial", 11, "bold"), bg="#3EC941", fg="#111111", activebackground="#32a435", activeforeground="#111111").pack(pady=5)
    
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
    """Sends intensity to xToys via the account Private Webhook HTTP endpoint.
    Endpoint : POST https://webhook.xtoys.app/<webhook_id>
    Body     : {"action": "setintensity", "intensity": <0-100>}   (JSON)
    The webhook_id is your Private Webhook ID from xtoys.app/me → Private Webhook.
    Cloud-relayed, so it works cross-device (VSE and the xToys tab can be on
    different machines) with no local port or helper app. Stateless HTTP POST.
    """
    _WEBHOOK_BASE = "https://webhook.xtoys.app/"
    _WS_BASE      = "wss://webhook.xtoys.app/"   # VSE listens here for the script's ack
                                                 # (proof it is running + which toys are
                                                 # bound) so the status can be truthful
    _WS_STALE_S   = 20.0     # script counts as "responding" if heard from within this
    _KEEPALIVE_SECS = 10.0   # re-send an unchanged level at least this often so the
                             # webhook connection stays warm (status stays "OK") and
                             # the toy keeps being driven while the user holds steady

    def __init__(self, webhook_id="", port=None):
        self.webhook_id      = webhook_id
        self.volume          = 0.0
        self._lock           = threading.Lock()
        self._last_sent_int  = None
        self._last_error     = None
        self._last_success_t = 0.0
        self._last_send_t    = 0.0
        self._session        = requests.Session()
        # ── WSS return channel (script → VSE): ack + bound-toy list ─────────────
        self._toys           = []      # toy/channel names bound under Generic Output
        self._last_ws_rx     = 0.0     # last time the script answered over the socket
        self._ws             = None
        self._ws_thread      = None
        self._ws_stop        = threading.Event()
        self._ws_id          = ""      # webhook id the current listener is bound to

    @property
    def enabled(self):
        return bool(self.webhook_id.strip())

    @property
    def connected(self):
        """True if a successful send happened in the last 30 s."""
        with self._lock:
            return (time.time() - self._last_success_t) < 30.0

    @property
    def listening(self):
        """True if the script answered over the socket recently — i.e. it is actually
        running and receiving our commands, not merely that the cloud returned 200."""
        with self._lock:
            return (time.time() - self._last_ws_rx) < self._WS_STALE_S

    @property
    def toys(self):
        """Names of the toys/channels bound under the script's Generic Output, as the
        script reports them over the socket. Empty until the script answers."""
        with self._lock:
            return list(self._toys)

    def maybe_reconnect(self):
        # HTTP send is stateless; this keeps the WSS return-channel listener alive.
        self.start_listener()

    def set_volume(self, vol, instant=False, floor=0.0, ceiling=1.0):
        with self._lock:
            self.volume = max(floor, min(ceiling, vol))
            if not self.enabled:
                return
            val_int = int(round(self.volume * 100))
            now = time.time()
            # Dedupe identical values, but re-send every _KEEPALIVE_SECS anyway so
            # the connection doesn't decay to "Connecting..." on a volume plateau
            # and the toy keeps being driven while the user holds steady.
            if (not instant and val_int == self._last_sent_int
                    and (now - self._last_send_t) < self._KEEPALIVE_SECS):
                return
            self._last_send_t = now
            url  = self._WEBHOOK_BASE + self.webhook_id.strip()
            body = {"action": "setintensity", "intensity": val_int}
            try:
                r = self._session.post(url, json=body, timeout=5.0)
                if r.status_code == 200:
                    self._last_sent_int  = val_int
                    self._last_error     = None
                    self._last_success_t = time.time()
                    log.info(f"xToys: SEND intensity={val_int}  (vol={self.volume:.3f})")
                else:
                    self._last_error = f"HTTP {r.status_code}"
                    log.warning(f"xToys: SEND failed — {self._last_error}")
            except Exception as e:
                # Log only the exception type — the full requests error text embeds
                # the webhook URL (with the id in the query string).
                self._last_error = type(e).__name__
                log.warning(f"xToys: SEND failed — {type(e).__name__}")

    def adjust_volume(self, delta, floor=0.0, ceiling=1.0):
        self.set_volume(self.volume + delta, floor=floor, ceiling=ceiling)

    def keepalive(self, floor=0.0, ceiling=1.0):
        # Re-assert the current level. set_volume only re-sends if the value is
        # unchanged AND _KEEPALIVE_SECS has elapsed, so calling this every tick is
        # cheap: it fires about once per interval on a plateau, keeping the webhook
        # warm (status stays OK) and the toy driven while the user holds steady.
        self.set_volume(self.volume, floor=floor, ceiling=ceiling)

    # ── WSS return channel ─────────────────────────────────────────────────────
    def start_listener(self):
        """Open (or re-open) the WSS listener for the current webhook id. The script
        answers commands here with an ack + its bound-toy list, which is how VSE knows
        it is genuinely connected rather than the cloud simply returning 200."""
        if not self.enabled:
            return
        wid = self.webhook_id.strip()
        if self._ws_thread and self._ws_thread.is_alive() and self._ws_id == wid:
            return
        self._stop_listener()
        self._ws_id = wid
        self._ws_stop.clear()
        with self._lock:
            self._last_ws_rx = 0.0
            self._toys = []
        self._ws_thread = threading.Thread(target=self._ws_run, daemon=True, name="XToysWS")
        self._ws_thread.start()

    def _stop_listener(self):
        self._ws_stop.set()
        ws = self._ws
        if ws:
            try:
                ws.close()
            except Exception:
                pass
        self._ws = None

    def _ws_run(self):
        while not self._ws_stop.is_set():
            wid = self._ws_id
            if not wid:
                time.sleep(2.0)
                continue
            try:
                ws = websocket.WebSocketApp(
                    self._WS_BASE + wid,
                    on_open=self._ws_on_open,
                    on_message=self._ws_on_message,
                    on_error=lambda ws, e: log.debug(f"xToys WS: {type(e).__name__}"),
                    on_close=lambda ws, c, m: None,
                )
                self._ws = ws
                # TLS verified by default (the webhook id is a device-control
                # capability in the URL) — do NOT pass cert_reqs=CERT_NONE.
                ws.run_forever(ping_interval=20, ping_timeout=10)
            except Exception as e:
                log.debug(f"xToys WS connect: {type(e).__name__}")
            finally:
                self._ws = None
            if not self._ws_stop.is_set():
                time.sleep(5.0)

    def _ws_on_open(self, ws):
        log.info("xToys: return-channel connected — waiting for the script to ack")
        # Ask which toys are bound; the script answers over this same socket.
        try:
            ws.send(json.dumps({"action": "getToys"}))
        except Exception:
            pass

    def _ws_on_message(self, ws, raw):
        try:
            data = json.loads(raw)
        except Exception:
            return
        if not isinstance(data, dict):
            return
        # Only the SCRIPT's own frames prove it is running — its ack/reply carries an
        # "action" and/or a "toys" list. xToys' relay frames ({"success":true},
        # {"event":"join",...}) arrive on every connect and must NOT be mistaken for
        # the script responding, or "listening" would be true with no script at all.
        if "action" not in data and "toys" not in data:
            return
        now  = time.time()
        toys = data.get("toys")
        with self._lock:
            was_listening = (now - self._last_ws_rx) < self._WS_STALE_S
            self._last_ws_rx = now
            if isinstance(toys, list):
                self._toys = [str(t) for t in toys]
        if not was_listening:
            log.info("xToys: script confirmed running (ack received)")

    def disconnect(self):
        self._stop_listener()
        with self._lock:
            self._last_sent_int  = None
            self._last_success_t = 0.0
            self._last_ws_rx     = 0.0
            self._toys           = []
            try:
                self._session.close()
            except Exception:
                pass
            self._session = requests.Session()


class SessionLogger:
    """
    Writes per-session event log to APPDATA/VisualStimEdger/sessions/.
    File is flushed atomically after every event, so it survives crashes.
    """

    def __init__(self, sessions_dir: pathlib.Path, version: str):
        today = datetime.date.today()
        date_str = f"{today.month}-{today.day}-{str(today.year)[2:]}"
        base = sessions_dir / f"VSE_SESSION_{date_str}"
        path = base.with_suffix(".json")
        n = 2
        while path.exists():
            path = sessions_dir / f"VSE_SESSION_{date_str}_{n}.json"
            n += 1
        self.path = path
        sessions_dir.mkdir(parents=True, exist_ok=True)
        self._events: list = []
        self._log("session_start", {"version": version})

    # ── public api ────────────────────────────────────────────────────────────

    def log_state_change(self, new_state: str):
        self._log("state_change", {"state": new_state})

    def log_edge_counted(self, edge_count: int):
        self._log("edge_counted", {"total_edges": edge_count})

    def log_letmecum(self, result: str, aggressiveness: str, odds_denominator: int):
        self._log("letmecum", {
            "result": result,
            "aggressiveness": aggressiveness,
            "odds": f"1/{odds_denominator}" if result == "granted" else f"{odds_denominator-1}/{odds_denominator}",
        })

    def log_cum(self):
        self._log("cum")

    def log_heart_rate(self, bpm, modifier: float):
        self._log("heart_rate", {"bpm": bpm, "modifier": round(modifier, 3)})

    def log_session_end(self, edge_count: int, cum_count: int, denial_count: int, elapsed_s: float):
        m, s = divmod(int(elapsed_s), 60)
        self._log("session_end", {
            "duration": f"{m:02d}:{s:02d}",
            "edges": edge_count,
            "orgasms": cum_count,
            "denials": denial_count,
        })

    # ── internals ─────────────────────────────────────────────────────────────

    @staticmethod
    def _now() -> str:
        return datetime.datetime.now().isoformat(timespec="seconds")

    def _log(self, event_type: str, data: dict | None = None):
        entry: dict = {"t": self._now(), "event": event_type}
        if data:
            entry.update(data)
        self._events.append(entry)
        self._flush()

    def _flush(self):
        tmp = self.path.with_suffix(".tmp")
        try:
            with open(tmp, "w", encoding="utf-8") as f:
                json.dump({"events": self._events}, f, indent=2, ensure_ascii=False)
            tmp.replace(self.path)
        except Exception:
            pass  # best effort — don't crash the app over logging


class HRClient:
    """
    Receives live heart-rate data from Pulsoid via WebSocket.
    Token: https://pulsoid.net/ui/keys  (requires Pulsoid paid plan as of 2025)
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
                # Verify TLS. The Pulsoid token rides in the wss URL query string,
                # so a MITM with a self-signed cert could otherwise read it (and
                # feed fake BPM, which drives the denial multiplier). websocket-client
                # verifies by default; do NOT pass cert_reqs=CERT_NONE.
                ws.run_forever(
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


class BLEHRClient:
    """
    Direct BLE heart rate client — reads from any BLE device advertising
    the standard Heart Rate Service (UUID 0x180D). Free, no subscription.
    Requires bleak: pip install bleak
    """
    HR_SERVICE = "0000180d-0000-1000-8000-00805f9b34fb"
    HR_CHAR    = "00002a37-0000-1000-8000-00805f9b34fb"
    _STALE_S   = 8.0

    def __init__(self, resting_bpm: int = 70, peak_bpm: int = 100):
        self.resting_bpm    = resting_bpm
        self.peak_bpm       = peak_bpm
        self.device_address: str | None = None
        self.device_name:    str | None = None
        self._bpm:    int | None = None
        self._bpm_hist       = deque(maxlen=6)
        self._last_rx        = 0.0
        self._lock           = threading.Lock()
        self._stop           = threading.Event()
        self._thread: threading.Thread | None = None
        self._loop            = None

    # ── interface (matches HRClient) ─────────────────────────────────────────

    @property
    def connected(self) -> bool:
        with self._lock:
            return self._bpm is not None and (time.time() - self._last_rx) < self._STALE_S

    @property
    def bpm(self) -> int | None:
        with self._lock:
            return self._bpm

    def smooth_bpm(self) -> float | None:
        with self._lock:
            if not self._bpm_hist:
                return None
            return sum(self._bpm_hist) / len(self._bpm_hist)

    def modifier(self) -> float:
        if not self.connected:
            return 1.0
        sbpm = self.smooth_bpm()
        if sbpm is None:
            return 1.0
        resting = float(self.resting_bpm)
        peak    = float(self.peak_bpm)
        if peak <= resting:
            return 1.0
        return 1.0 + max(0.0, min(1.0, (sbpm - resting) / (peak - resting)))

    # ── lifecycle ─────────────────────────────────────────────────────────────

    def start(self, address: str, name: str = ""):
        self.device_address = address
        self.device_name    = name or address
        self._stop.clear()
        self._thread = threading.Thread(target=self._run, daemon=True, name="BLEHRClient")
        self._thread.start()
        log.info(f"BLE HR: connecting to {self.device_name} ({address})")

    def stop(self):
        self._stop.set()
        with self._lock:
            self._bpm = None
        lp = self._loop
        if lp and lp.is_running():
            lp.call_soon_threadsafe(lp.stop)

    # ── internals ─────────────────────────────────────────────────────────────

    def _on_notify(self, _sender, data: bytearray):
        flags = data[0]
        bpm   = int.from_bytes(data[1:3], "little") if (flags & 0x01) else data[1]
        with self._lock:
            self._bpm = bpm
            self._bpm_hist.append(bpm)
            self._last_rx = time.time()

    def _run(self):
        import asyncio
        self._loop = asyncio.new_event_loop()
        asyncio.set_event_loop(self._loop)
        try:
            self._loop.run_until_complete(self._connect_loop())
        finally:
            self._loop.close()
            self._loop = None

    async def _connect_loop(self):
        try:
            from bleak import BleakClient
        except ImportError:
            log.error("BLE HR: bleak not installed — run: pip install bleak")
            return
        import asyncio
        while not self._stop.is_set():
            try:
                async with BleakClient(self.device_address, timeout=10.0) as client:
                    log.info(f"BLE HR: connected to {self.device_name}")
                    await client.start_notify(self.HR_CHAR, self._on_notify)
                    while not self._stop.is_set() and client.is_connected:
                        await asyncio.sleep(1.0)
                    if client.is_connected:
                        await client.stop_notify(self.HR_CHAR)
            except Exception as e:
                log.warning(f"BLE HR: {e} — retry in 5 s")
                with self._lock:
                    self._bpm = None
                if self._stop.is_set():
                    break
                await asyncio.sleep(5.0)

    @staticmethod
    def scan_sync(timeout: float = 6.0) -> list[tuple[str, str]]:
        """Synchronous BLE scan. Blocks for `timeout` seconds. Returns [(address, name), ...]."""
        import asyncio, sys
        async def _scan():
            try:
                from bleak import BleakScanner
            except ImportError:
                log.warning("BLE scan: bleak not installed — pip install bleak")
                return []
            try:
                devices = await BleakScanner.discover(
                    timeout=timeout,
                    service_uuids=[BLEHRClient.HR_SERVICE],
                )
                return [(d.address, d.name or d.address) for d in devices]
            except Exception as e:
                log.warning(f"BLE scan discover error: {e}")
                return []
        try:
            # On Windows, bleak uses WinRT which requires a ProactorEventLoop.
            # Setting the policy explicitly before creating the loop fixes
            # crashes when scan_sync is called from a background thread.
            if sys.platform == "win32":
                asyncio.set_event_loop_policy(asyncio.WindowsProactorEventLoopPolicy())
            loop = asyncio.new_event_loop()
            asyncio.set_event_loop(loop)
            try:
                return loop.run_until_complete(_scan())
            finally:
                loop.close()
                asyncio.set_event_loop(None)
        except Exception as e:
            log.warning(f"BLE scan failed: {e}")
            return []


class VoiceEngine:
    """Offline keyword/phrase recogniser using Vosk + sounddevice.

    Runs a background thread that feeds mic audio through a KaldiRecognizer
    with a restricted grammar.  When a keyword is matched the registered
    callback is called from the audio thread — the caller is responsible for
    posting to the UI thread (root.after) if needed.

    Model: vosk-model-small-en-us (~40 MB) placed at
        <resource_path>/models/vosk-model-small-en-us
    Install: pip install vosk sounddevice
    """

    # Phrase vocabulary (order matters: phrases before single words in matching)
    _PHRASES = [
        "erect up", "erect down",
        "flaccid up", "flaccid down",
        "edging up", "edging down",
        "find head", "set lines",
        "switch source", "resume session",
        "clear exclude",
        "let me cum",
    ]
    _WORDS = [
        "came", "cumming", "pause", "resume",
        "select", "here", "confirm", "cancel", "again", "back",
        "exclude", "please",
        "up", "down", "left", "right",
        "one", "two", "three", "four", "five",
        "six", "seven", "eight", "nine",
    ]
    _NUM_MAP = {
        "one": 1, "two": 2, "three": 3, "four": 4, "five": 5,
        "six": 6, "seven": 7, "eight": 8, "nine": 9,
    }
    MODEL_DIR = "vosk-model-small-en-us"

    def __init__(self, model_path: str, device_name: str = "", samplerate: int = 16000):
        self.model_path  = model_path
        self.device_name = device_name
        self.samplerate  = samplerate
        self._safeword   = "red"
        self._cb         = None          # fn(keyword: str), called from audio thread
        self._stop       = threading.Event()
        self._thread: threading.Thread | None = None
        self._level      = 0.0           # RMS 0–1 for level meter UI
        self._lock       = threading.Lock()

    # ── public api ────────────────────────────────────────────────────────────

    @staticmethod
    def model_path_default() -> str:
        base = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
        return os.path.join(base, "models", VoiceEngine.MODEL_DIR)

    @staticmethod
    def model_available(path: str) -> bool:
        return bool(path) and os.path.isdir(path)

    @staticmethod
    def list_input_devices() -> list[str]:
        try:
            import sounddevice as sd
            return [d['name'] for d in sd.query_devices() if d['max_input_channels'] > 0]
        except Exception:
            return []

    def set_safeword(self, word: str):
        self._safeword = word.strip().lower() or "red"

    def set_callback(self, fn):
        """Register the single keyword callback (replaces previous one)."""
        with self._lock:
            self._cb = fn

    @property
    def level(self) -> float:
        return self._level

    @property
    def running(self) -> bool:
        return bool(self._thread and self._thread.is_alive())

    def start(self):
        if self.running:
            return
        self._stop.clear()
        self._thread = threading.Thread(target=self._run, daemon=True, name="VoiceEngine")
        self._thread.start()
        log.info(f"VoiceEngine: starting (device={self.device_name or 'default'})")

    def stop(self):
        self._stop.set()
        t = self._thread
        self._thread = None
        if t and t.is_alive():
            t.join(timeout=1.0)

    # ── internals ─────────────────────────────────────────────────────────────

    def _build_grammar(self) -> str:
        words: set[str] = set()
        for ph in self._PHRASES:
            words.update(ph.split())
        words.update(self._WORDS)
        for part in self._safeword.split():
            words.add(part)
        words.add("[unk]")
        return json.dumps(sorted(words))

    def _dispatch(self, kw: str):
        with self._lock:
            cb = self._cb
        if cb:
            try:
                cb(kw)
            except Exception as e:
                log.warning(f"VoiceEngine callback error: {e}")

    def _match(self, text: str):
        """Match recognised text against vocabulary; dispatch first hit."""
        # Safeword always wins
        if self._safeword and self._safeword in text.split():
            self._dispatch("safeword")
            return
        # Multi-word phrases before single words
        for phrase in self._PHRASES:
            if phrase in text:
                self._dispatch(phrase)
                return
        # Single words (exact word boundary)
        words_in = set(text.split())
        for kw in self._WORDS:
            if kw in words_in:
                self._dispatch(kw)
                return

    def _run(self):
        try:
            from vosk import Model, KaldiRecognizer, SetLogLevel
            SetLogLevel(-1)
        except ImportError:
            log.error("VoiceEngine: vosk not installed — pip install vosk")
            return
        try:
            import sounddevice as sd
        except ImportError:
            log.error("VoiceEngine: sounddevice not installed")
            return
        if not os.path.isdir(self.model_path):
            log.error(f"VoiceEngine: model not found at {self.model_path}")
            return
        try:
            model = Model(self.model_path)
            rec   = KaldiRecognizer(model, self.samplerate, self._build_grammar())
        except Exception as e:
            log.error(f"VoiceEngine: init failed: {e}")
            return
        # Resolve mic device index
        dev_idx = None
        if self.device_name:
            try:
                for i, d in enumerate(sd.query_devices()):
                    if d['max_input_channels'] > 0 and self.device_name in d['name']:
                        dev_idx = i
                        break
            except Exception:
                pass
        blocksize = int(self.samplerate * 0.1)   # 100ms chunks
        try:
            with sd.RawInputStream(
                samplerate=self.samplerate, blocksize=blocksize,
                device=dev_idx, dtype='int16', channels=1,
            ) as stream:
                while not self._stop.is_set():
                    data, _ = stream.read(blocksize)
                    raw = bytes(data)
                    pcm = np.frombuffer(raw, dtype=np.int16).astype(np.float32)
                    self._level = min(1.0, float(np.sqrt(np.mean(pcm ** 2))) / 32768.0 * 8.0)
                    if rec.AcceptWaveform(raw):
                        text = json.loads(rec.Result()).get("text", "").strip().lower()
                        if text and text != "[unk]":
                            log.debug(f"VoiceEngine recognised: {text!r}")
                            self._match(text)
        except Exception as e:
            if not self._stop.is_set():
                log.error(f"VoiceEngine: stream error: {e}")


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


def _pycaw_flow(d):
    """Return the flow int for a pycaw device, or -1 if inaccessible."""
    try:
        return int(d.flow)
    except Exception:
        return -1


def list_audio_devices():
    """Return all render (output) devices.

    Tries pycaw first (real IMMDevice objects = reliable volume control).
    Falls back to sounddevice if pycaw finds nothing usable.
    """
    # Level 1: pycaw — enumerate named devices, prefer render (flow=0)
    try:
        all_devs = AudioUtilities.GetAllDevices()
        named = [d for d in all_devs if d._dev is not None and d.FriendlyName]
        render = [d for d in named if _pycaw_flow(d) == 0]
        result = render if render else named   # if flow is broken, show all named
        if result:
            log.info(f"WinAudio: pycaw found {len(result)} device(s) "
                     f"({'render-only' if render else 'flow unavailable, all named'})")
            return result
    except Exception as e:
        log.error(f"WinAudio: GetAllDevices failed: {e}")

    # Level 2: sounddevice (names only — _resolve_dev fuzzy-matches to pycaw for control)
    try:
        import sounddevice as sd
        devs = [_SounddeviceAudioDevice(d['name'])
                for d in sd.query_devices()
                if d['max_output_channels'] > 0]
        if devs:
            log.info(f"WinAudio: sounddevice found {len(devs)} output device(s)")
            return devs
    except Exception as e:
        log.error(f"WinAudio: sounddevice failed: {e}")

    # Level 3: default speakers only
    try:
        default = AudioUtilities.GetSpeakers()
        if default:
            log.info("WinAudio: using default speakers fallback")
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
            all_devs = [d for d in AudioUtilities.GetAllDevices()
                        if d._dev is not None and d.FriendlyName]
            nl = name.lower()
            # Pass 1: exact match
            for d in all_devs:
                if d.FriendlyName == name:
                    return d._dev
            # Pass 2: substring — one is contained in the other
            for d in all_devs:
                fl = d.FriendlyName.lower()
                if nl in fl or fl in nl:
                    log.debug(f"WinAudio: fuzzy matched '{name}' → '{d.FriendlyName}'")
                    return d._dev
            # Pass 3: prefix match — sounddevice truncates names at ~31 chars
            prefix = nl[:31]
            for d in all_devs:
                if d.FriendlyName.lower().startswith(prefix):
                    log.debug(f"WinAudio: prefix matched '{name}' → '{d.FriendlyName}'")
                    return d._dev
            # Pass 4: first render device
            render = [d for d in all_devs if _pycaw_flow(d) == 0]
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
        try:
            self._loop.run_until_complete(self._serve())
        except RuntimeError:
            pass  # loop stopped mid-serve during shutdown — expected teardown, not a crash

    async def _serve(self):
        try:
            server = await asyncio.start_server(self._handle, '127.0.0.1', self.PORT)
        except OSError as e:
            log.error(f"Overlay: cannot bind port {self.PORT} ({e}) — OBS overlay disabled")
            return
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
            # Snapshot the client list under the lock, then do all async I/O
            # outside it — awaiting drain() while holding a threading.Lock can
            # stall other threads trying to acquire the lock.
            with self._lock:
                clients = list(self._clients)
            dead = []
            for w in clients:
                try:
                    w.write(frame)
                    await w.drain()
                except Exception:
                    dead.append(w)
            if dead:
                with self._lock:
                    for w in dead:
                        if w in self._clients:
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
        self._device      = None
        self._device_name = ""        # "" = system default; set via set_output_device()
        self._stop_flag   = False
        self._state       = "stopped"   # stopped | playing | paused
        self._skip        = 0           # +1 next, -1 prev (set from main thread)
        self._playlist    = []          # list[pathlib.Path]
        self._idx         = 0
        self.loop_mode    = "folder"    # "track" | "folder"
        self.volume       = 0.5         # 0.0–1.0, written from main thread, read from audio thread
        # UI notification — set by audio thread, read by main thread (strings are atomic)
        self.track_name = ""
        self.track_info = ""          # e.g. "3 / 12"

    @staticmethod
    def list_devices() -> list[str]:
        """Return names of available playback devices."""
        try:
            return [d['name'] for d in _miniaudio.Devices().get_playbacks()]
        except Exception:
            return []

    def set_output_device(self, name: str):
        """Switch output device — restarts playback if currently playing."""
        self._device_name = name
        if self._state == "playing":
            self._stop_device()
            self._start()

    # ── Generator ─────────────────────────────────────────────────────────────
    def _gen(self):
        required_frames = yield b""   # prime
        _consecutive_fails = 0
        while not self._stop_flag:
            if not self._playlist or self._state == "stopped":
                _consecutive_fails = 0
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
                _consecutive_fails += 1
                self._idx = (self._idx + 1) % len(self._playlist)
                # If every file in the playlist has failed, yield silence to
                # avoid a tight spin loop that pegs CPU at 100 %.
                if _consecutive_fails >= len(self._playlist):
                    log.warning("MusicPlayer: all files unreadable — yielding silence")
                    required_frames = yield bytes(required_frames * 8)
                    _consecutive_fails = 0
                continue
            _consecutive_fails = 0  # reset on successful open

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
            device_id = None
            if self._device_name:
                try:
                    for d in _miniaudio.Devices().get_playbacks():
                        if d['name'] == self._device_name:
                            device_id = d['id']
                            break
                except Exception:
                    pass
            self._device = _miniaudio.PlaybackDevice(
                output_format=_miniaudio.SampleFormat.FLOAT32,
                nchannels=2, sample_rate=44100,
                buffersize_msec=150,
                device_id=device_id,
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
    "Evil": {
        "BG": "#0d0000", "SURFACE": "#1a0000", "SURFACE2": "#250500",
        "ACCENT": "#cc0000", "ACCENT_H": "#990000",
        "RED": "#ff2200", "RED_HOV": "#cc1800",
        "GREEN": "#3a8a1a", "GREEN_H": "#2a6a10",
        "BLUE": "#440000", "BLUE_H": "#330000",
        "YELLOW": "#cc3300", "YELLOW_H": "#993300",
        "TEXT": "#ffcccc", "TEXT_DIM": "#7a3a3a", "BORDER": "#660000",
    },
    "ReThorn": {
        "BG": "#111111", "SURFACE": "#1a1a1a", "SURFACE2": "#222222",
        "ACCENT": "#F5A623", "ACCENT_H": "#d48e1a",
        "RED": "#FF4444", "RED_HOV": "#cc3636",
        "GREEN": "#3EC941", "GREEN_H": "#32a435",
        "BLUE": "#444CFC", "BLUE_H": "#363dca",
        "YELLOW": "#F5A623", "YELLOW_H": "#d48e1a",
        "TEXT": "#e8e8f0", "TEXT_DIM": "#6a6a7a", "BORDER": "#333333",
    },
    "Void": {
        "BG": "#0d0012", "SURFACE": "#160020", "SURFACE2": "#1e002c",
        "ACCENT": "#9d4edd", "ACCENT_H": "#7b2fc9",
        "RED": "#ff3366", "RED_HOV": "#cc2255",
        "GREEN": "#39d987", "GREEN_H": "#2ab36d",
        "BLUE": "#2d2b8a", "BLUE_H": "#232170",
        "YELLOW": "#ffb300", "YELLOW_H": "#cc8f00",
        "TEXT": "#e8e0ff", "TEXT_DIM": "#6a5a80", "BORDER": "#2d1a40",
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
    _C_BG        = "#111111"
    _C_SURFACE   = "#1a1a1a"
    _C_SURFACE2  = "#222222"
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
    _C_TEXT      = "#e8e8f0"
    _C_TEXT_DIM  = "#6a6a7a"
    _C_BORDER    = "#333333"

    # ── YOLO reanchoring
    _YOLO_INTERVAL_LOCKED = 30   # frames between YOLO checks when tracking is solid
    _YOLO_INTERVAL_LOST   = 5    # frames between YOLO checks when tracking is suspect/lost
    _YOLO_CONFIRM  = 2     # consecutive detections in same area before reanchoring
    _YOLO_MAX_JUMP = 2.0   # max allowed jump as multiple of current bbox diagonal
    # Tracker plausibility
    _SIZE_RATIO_MAX  = 2.5
    _MAX_JUMP_FACTOR = 2.5
    # Deep tracker (TrackerVit) — far more occlusion/motion-robust than CSRT and
    # ~constant cost; CSRT stays as a selectable fallback. A confidence score below
    # this counts as a suspect frame (routes to the appearance re-find in lock mode).
    _VIT_MIN_SCORE   = 0.30
    # Lock-on re-find: when "Lock on target" is on and CSRT drops the target, match
    # a saved appearance patch near the last position instead of freezing (YOLO
    # reanchor is disabled while locked, and can't see a plug/clip anyway).
    _LOCK_TPL_REFRESH_S = 1.0   # refresh the target's appearance patch this often while healthy
    _RELOCK_SEARCH      = 3.5   # re-find search window, as a multiple of the target size
    _RELOCK_MIN_SCORE   = 0.55  # min TM_CCOEFF_NORMED to accept a re-lock (keeps it off hands/skin)
    # Colour lock — track a vividly-coloured marker by its hue (e.g. a bright clip).
    # Cheap, rotation/scale/occlusion-proof; only works for saturated colours.
    _COLOR_HUE_TOL       = 12      # +/- hue band accepted (OpenCV hue is 0-179)
    _COLOR_MIN_AREA_FRAC = 0.0004  # ignore colour blobs smaller than this fraction of the frame
    _COLOR_SEARCH        = 2.5     # search window as a multiple of the target diagonal
    _COLOR_MAX_JUMP      = 2.0     # accept the nearest colour blob only within this * diagonal
    _COLOR_MAX_AREA      = 6.0     # reject blobs bigger than this * the box — that's skin, not a marker
    _COLOR_SAT_FRAC      = 0.20    # >= this fraction of saturated box pixels picks hue mode, else dark

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
        self.heights         = {"Edging": None, "Erect": None, "Flaccid": None, "PONR": None}
        # Point of No Return: optional yellow panic line. Crossing ABOVE it hard-
        # cuts volume to the floor. Edge-triggered — True while the head is past
        # the line so we send one instant cut, then hold. Nothing depends on it.
        self._ponr_triggered = False
        # Remembered window width (persisted). None = use the default on first
        # boot. Height always auto-fits to content, so only width is sticky.
        self._win_w          = None

        self.last_vol_time   = time.time()
        self.tracking_paused    = False
        self.last_bbox          = tuple(int(v) for v in bbox)
        self.tracking_ok        = True
        self._track_msg         = ""
        self.yolo_frame_counter = 0
        self.yolo_candidate     = None  # (bbox, hits) pending confirmation
        self._lock_template     = None  # appearance patch of the locked target (BGR)
        self._lock_tpl_time     = 0.0   # last time the patch was refreshed
        self._color_lock        = False # track by colour instead of CSRT
        self._color_hue         = None  # locked hue centre (0-179); None until derived
        self._color_tol         = self._COLOR_HUE_TOL
        self._color_s_min       = 60    # saturation / value floors for the colour mask
        self._color_v_min       = 50
        self._color_needs_derive = False # re-derive colour from the box on the next frame
        self._lock_mode         = "hue"  # "hue" (vivid marker) or "dark" (black electrode)
        self._dark_v            = 90     # dark mode: pixels darker than this are the target
        self._dark_s            = 110    # dark mode: ...and less saturated than this
        self._head_y_history    = deque(maxlen=HEAD_Y_SMOOTH)
        self._frame_times       = deque(maxlen=30)  # for FPS calculation
        self._frame_dt          = deque(maxlen=300) # per-frame Δt (ms) for p95/p99
        self._last_frame_t      = 0.0               # ts of previous processed frame
        self._last_perf_log     = 0.0               # throttle for the perf log line

        # Session stats
        self.session_start   = time.time()
        self.state_times     = {"Edging": 0.0, "Erect": 0.0, "Flaccid": 0.0}
        self.edge_count      = 0
        self._prev_state     = "Erect"
        self._last_state_time = time.time()
        self._stats_tick     = 0   # throttle label updates

        # Hold
        self.hold_active = False

        # Lazy-initialised state — declared here to keep attribute creation in one place
        self._proc_frame_count        = 0
        self._last_status_time        = 0.0
        self._last_display_time       = 0.0
        self._play_mode               = False
        self._auto_btn_state          = None
        self._letmecum_cooldown_until = 0.0
        self._cum_grant_expires       = 0.0
        self._last_letmecum_result    = None
        self._last_letmecum_time      = 0.0

        # Cum cooldown  (None = not active)
        self._cum_count = 0
        self._denial_count = 0
        self._cum_override_range = True   # True = cum goes to 100%, False = respect ceiling
        self._cum_stopped = False         # True after "I've CUM" until "Resume"
        self._refractory_mins = 5        # cooldown duration (0 = no timer, manual resume)
        self._refractory_until: float = 0.0  # epoch time when refractory ends

        # ── Volume recovery (fixes "parks forever after an edge") ──────────────
        self._erect_recovery = True      # passive gentle climb when held in erect zone
        self._active_edging  = False     # aggressive auto-ramp cycle after holding off edge
        self._recovery_pct = dict(RECOVERY_PCT_DEFAULT)  # per-level passive climb %/s
        self._ae_hold_secs = dict(AE_HOLD_SECS_DEFAULT)  # per-level hold before ramp (s)
        self._ae_ramp_pct  = dict(AE_RAMP_PCT_DEFAULT)   # per-level ramp %/s
        self._ae_below_since: float | None = None  # when we last backed off the edge

        # ── Auto-cum detection ────────────────────────────────────────────────
        self._auto_cum_enabled     = False
        self._auto_cum_delay       = 5      # seconds before _on_cum fires (0-10)
        self._auto_cum_sensitivity = 5      # 1-10, higher = easier to trigger
        self._cum_detect_buf: deque = deque(maxlen=CUM_DETECT_MAXLEN)

        # ── Hands-free cum mode (settings-only toggle) ────────────────────────
        self._hf_enabled    = False
        self._hf_min_edges  = dict(self._HF_MIN_EDGES_DEFAULT)
        self._hf_cum_chance = dict(self._HF_CUM_CHANCE_DEFAULT)

        # ── Bondage mode ──────────────────────────────────────────────────────
        self._bondage_active           = False
        self._bondage_standby          = False   # True after safeword, awaiting "resume session"
        self._bondage_safeword         = "red"
        self._bondage_mic_device       = ""      # "" = system default
        self._bondage_safeword_saved   = False   # True when safeword was loaded from config
        self._bondage_configured       = False   # True once bondage has been started this session
        self._evil_pulse_job           = None    # after() handle for snark pulse
        self._voice_engine: VoiceEngine | None = None
        # Grid navigator
        self._grid_active  = False
        self._grid_mode    = 'head'  # 'head' | 'exclude'
        self._grid_region  = None   # (x, y, w, h) in frame px; None = full frame
        self._grid_depth   = 0      # 0 = full frame, max 3
        # Source picker (voice "switch source")
        self._source_picker_active  = False
        self._source_picker_windows: list = []   # [(hwnd, title), ...]
        self._source_picker_win: object = None   # CTkToplevel reference
        self._cum_peak_activity    = 0.0    # rolling max of slow-window std
        self._cum_score            = 0.0    # accumulator 0-25
        self._cum_cd_active        = False  # True during the cancel-window countdown
        self._cum_cd_job           = None   # root.after handle
        self._cum_cd_remaining     = 0      # seconds left in countdown
        # Brightness spike detection
        self._bright_buf:            deque = deque(maxlen=900)  # ~30s at 30fps
        self._bright_baseline_ready = False
        self._bright_baseline_mean  = 0.0
        self._bright_baseline_std   = 2.0   # conservative floor
        self._ui_font_size = 11           # base font size for UI
        self._theme_name = DEFAULT_THEME
        self._language   = "en"           # UI language code; see _load_catalog / tr
        self._cum_time: float | None = None
        self._cum_undo_active = False   # UX-1: misclick protection
        self._cum_undo_job    = None    # UX-1: root.after handle for undo window
        self._cum_allowed = False
        self._cum_odds = dict(self._CUM_ODDS_DEFAULT)
        self._denial_phrases = list(self._DENIAL_PHRASES_DEFAULT)

        self._loading_config = False

        # Evil Mode state
        self._evil_mode      = False
        self._pre_evil_theme = DEFAULT_THEME
        self._ruin_odds      = dict(self._RUIN_ODDS_DEFAULT)
        self._ruin_phrases   = list(self._RUIN_PHRASES_DEFAULT)
        self._ruin_count     = 0
        self._ruin_active    = False   # True while a ruin pulse sequence is running
        self._ruin_jobs: list = []     # pending root.after ids for the ruin, cancellable
        self._exclusion_zones: list[tuple[int,int,int,int]] = []   # (x,y,w,h) in frame px
        self._reanchor_lock = False   # True = freeze tracker on user's target, no YOLO reanchor
        self._ez_drawing    = False   # True while user is dragging a new zone
        self._ez_disp_start = None   # (x,y) display coords of drag start
        self._ez_disp_end   = None   # (x,y) display coords of current drag end

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
        self._ui_queue       = queue.Queue()   # background→main-thread calls (see _call_on_main)
        self._capture_thread = threading.Thread(target=self._capture_loop, daemon=True)
        self._capture_thread.start()

        # CV — deep tracker (Vit) with a CSRT fallback, on a downscaled frame when a
        # slow machine needs it (see _make_tracker / _downscale / _tracker_init).
        self._proc_width        = 0
        self._track_scale       = 1.0
        self._tracker_backend   = "csrt"   # "csrt" (proven default) or "vit" (deep, opt-in)
        self._tracker_has_score = False    # set by _make_tracker if the backend scores
        self._tracker_dirty     = False    # rebuild + re-init the tracker next frame
        self.tracker  = self._make_tracker()
        self._tracker_init(initial_frame, bbox)
        self.detector = DickDetector()

        # Output clients
        self.restim           = RestimClient(RESTIM_HOST, RESTIM_PORT, TCODE_AXIS)
        self.xtoys            = XToysClient()
        self.win_audio        = None
        self.win_devices      = []
        self._orig_win_volume = None  # restored on exit
        self.music_player     = MusicPlayer() if _MINIAUDIO_OK else None
        self.hr_client        = HRClient()
        self.ble_hr_client  = BLEHRClient()
        self._hr_source     = "pulsoid"   # "pulsoid" | "ble"
        self._ble_addr      = ""
        self._ble_name      = ""

        # Root window
        self.root = ctk.CTk()
        self.root.title("VisualStimEdger")
        _icon = pathlib.Path(resource_path("icon.ico"))
        if _icon.exists():
            self.root.iconbitmap(str(_icon))

        # ── Tcl/Tk error handler ───────────────────────────────────────────────
        # CustomTkinter schedules cursor-blink after-callbacks that can fire
        # after the widget is destroyed (pack_forget / window close).  In a
        # compiled exe the resulting "invalid command name" Tcl error can
        # escalate to Tcl_Panic → EXCEPTION_BREAKPOINT crash instead of the
        # harmless stderr noise you get in dev.  Override bgerror to catch and
        # swallow known-benign cases; log everything else.
        def _tk_report_callback_exception(exc_type, exc_val, exc_tb):
            msg = str(exc_val)
            if "invalid command name" in msg and "_blink" in msg:
                return  # CTk cursor blink on destroyed widget — harmless
            log.error(
                "Tk callback exception: %s: %s",
                exc_type.__name__, exc_val,
                exc_info=(exc_type, exc_val, exc_tb),
            )
        self.root.report_callback_exception = _tk_report_callback_exception

        # tkinter vars — must be created after root exists
        self.min_vol_var = tk.DoubleVar(value=0.0)
        self.max_vol_var = tk.DoubleVar(value=100.0)
        self.aggr_var    = tk.StringVar(value="Middle")
        self.proc_res_var = tk.StringVar(value="Full")   # tracking-resolution preset
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
                if isinstance(_pre.get("language"), str):
                    self._language = _pre["language"]
        except Exception:
            pass
        _load_catalog(self._language)     # load translations before any UI is built
        self._apply_theme(self._theme_name)

        # Session event log
        self.session_logger = SessionLogger(CONFIG_PATH.parent / "sessions", VERSION)
        self.root.after(30_000, self._hr_log_poll)

        self._build_ui()
        self._on_aggr_change()   # set initial aggressiveness colour
        self._load_config()
        # Sync toggle to loaded config value
        self._auto_cum_var.set(self._auto_cum_enabled)
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        self._start_update_check()

    def run(self):
        self.root.after(5, self._update_frame)
        self.root.after(30, self._pump_ui_queue)   # drain background→main-thread calls
        self.root.after(self._GC_INTERVAL_MS, self._gc_collect)  # main-thread GC sweep
        self.root.mainloop()
        # If mainloop ever returns without going through _on_close (which hard-exits
        # on its own), do the same collapsed shutdown: flush the critical side effects
        # on the main thread, then hard-exit — skipping the interpreter teardown that
        # would finalise Tk objects off-thread (Tcl_AsyncDelete, exit 3).
        self._running = False
        try:
            self._flush_critical()
        except Exception:
            pass
        try:
            logging.shutdown()
        except Exception:
            pass
        os._exit(0)

    def _call_on_main(self, fn, *args):
        """Schedule fn(*args) on the Tk main thread. Background threads MUST use this
        instead of touching Tk (or self.root.after) directly — a cross-thread Tcl call
        is what trips 'Tcl_AsyncDelete: async handler deleted by the wrong thread' when
        the app shuts down. Drained by _pump_ui_queue."""
        self._ui_queue.put((fn, args))

    def _pump_ui_queue(self):
        """Run background→main calls on the Tk thread, then reschedule itself."""
        try:
            while True:
                fn, args = self._ui_queue.get_nowait()
                try:
                    fn(*args)
                except Exception:
                    log.exception("UI-queue callback failed")
        except queue.Empty:
            pass
        if self._running:
            self.root.after(30, self._pump_ui_queue)

    def _gc_collect(self):
        """Reclaim cyclic garbage on the MAIN thread. Automatic GC is disabled
        process-wide (see main) because when it fired on the capture thread it
        finalized main-thread Tk objects off-thread — the Tcl_AsyncDelete abort.
        Collecting here keeps memory bounded with every Tk finalizer on the main
        thread. Full collect every _GC_INTERVAL_MS; cheap for this heap size."""
        try:
            gc.collect()
        except Exception:
            pass
        if self._running:
            self.root.after(self._GC_INTERVAL_MS, self._gc_collect)

    # ------------------------------------------------------------------ UI

    def _build_ui(self):
        root = self.root
        root.configure(fg_color=self._C_BG)
        root.minsize(540, 200)

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
        settings_menu.add_command(label=tr("Settings..."), command=self._open_settings)
        menubar.add_cascade(label=tr("Settings"), menu=settings_menu)

        about_menu = tk.Menu(menubar, tearoff=0,
                             bg=self._C_BG, fg=self._C_TEXT,
                             activebackground=self._C_ACCENT, activeforeground="white")
        about_menu.add_command(label=f"Version  v{VERSION}", state="disabled")
        about_menu.add_command(label="Dev: Sir Thorn", state="disabled")
        about_menu.add_separator()
        about_menu.add_command(label=tr("GitHub (latest release)"),
                               command=lambda: webbrowser.open("https://github.com/blucrew/VisualStimEdger/releases/latest"))
        about_menu.add_command(label=tr("Ko-fi (support)"),
                               command=lambda: webbrowser.open("https://ko-fi.com/stimstation"))
        menubar.add_cascade(label=tr("About"), menu=about_menu)

        root.configure(menu=menubar)

        # ── Update banner (hidden until needed) ───────────────────────────────
        self._update_banner = ctk.CTkFrame(root, fg_color=self._C_SURFACE, corner_radius=0)
        self._update_label  = ctk.CTkLabel(self._update_banner, text="",
                                           text_color="white", font=lbl)
        self._update_label.pack(side=tk.LEFT, padx=P, pady=6)
        self._update_btn = ctk.CTkButton(self._update_banner, text=tr("Download"), width=100,
                                         fg_color=self._C_ACCENT, hover_color=self._C_ACCENT_H,
                                         text_color="white", font=lbl, corner_radius=4)
        self._update_btn.pack(side=tk.RIGHT, padx=P, pady=6)

        # ── Video + calibration buttons ───────────────────────────────────────
        top_frame = ctk.CTkFrame(root, fg_color="transparent")
        top_frame.pack(padx=P, pady=(P, 4), fill=tk.X)
        self._first_widget = top_frame

        # ── Settings area — no scrollbar; window auto-resizes to content ─────
        sf = ctk.CTkFrame(root, fg_color="transparent", corner_radius=0)
        sf.pack(fill=tk.X, padx=0, pady=0)
        self._sf = sf

        # video — plain tk.Label so ImageTk works without wrapping
        vid_col = ctk.CTkFrame(top_frame, fg_color="transparent")
        vid_col.pack(side=tk.LEFT)
        self._vid_shell = ctk.CTkFrame(vid_col, fg_color=self._C_SURFACE,
                                      corner_radius=8, border_width=1, border_color=self._C_BORDER,
                                      width=330, height=260)
        vid_shell = self._vid_shell
        vid_shell.pack()
        vid_shell.pack_propagate(False)
        self.video_label = tk.Label(vid_shell, bg=self._C_SURFACE)
        self.video_label.pack(padx=3, pady=(3, 0), fill=tk.BOTH, expand=True)
        self._snark_label = ctk.CTkLabel(vid_shell, text="", font=ctk.CTkFont(size=11, slant="italic"),
                                          text_color="#ff4444", height=20)
        self._snark_label.pack(padx=3, pady=(0, 3))
        # #2 — tribute button, shown ONLY during the cum-grant countdown; skippable,
        # never blocks the count. Created hidden; _grant_cum packs it, the cum-window
        # end (_tick_cum_grant expiry / cum-stopped) hides it again.
        self._tribute_btn = ctk.CTkButton(
            vid_shell, text=tr("😈 tribute your dom before you're allowed → ☕"), height=24,
            fg_color="#8a1a1a", hover_color="#6a0a0a", text_color="#F5A623",
            border_width=1, border_color="#F5A623",
            font=ctk.CTkFont(size=11, weight="bold"),
            command=lambda: webbrowser.open("https://ko-fi.com/stimstation"))
        # Re-select buttons live directly under the video preview
        _vbr = ctk.CTkFrame(vid_col, fg_color="transparent")
        _vbr.pack(fill=tk.X, pady=(4, 0))
        _vbtn_kw = dict(font=ctk.CTkFont(size=10), height=28,
                        fg_color=self._C_SURFACE2, hover_color="#4a4a4a",
                        text_color=self._C_TEXT, border_width=1, border_color=self._C_BORDER)
        ctk.CTkButton(_vbr, text=tr("Re-Select Feed"), command=self._reselect_feed,
                      **_vbtn_kw).pack(side=tk.LEFT, expand=True, fill=tk.X, padx=(0, 2))
        ctk.CTkButton(_vbr, text=tr("Re-Select Head"), command=self._reselect_head,
                      **_vbtn_kw).pack(side=tk.LEFT, expand=True, fill=tk.X, padx=(2, 0))

        # Lock toggle — freeze the tracker on the user's target (no auto-reanchor).
        # Best with a high-contrast target like a plug/electrode that the tracker
        # would otherwise get pulled toward; keeps it from "skipping" off the box.
        self._lock_track_var = tk.BooleanVar(value=self._reanchor_lock)
        _lock_sw = ctk.CTkSwitch(
            vid_col, text=tr("🔒 Lock on target (no auto-reanchor)"),
            variable=self._lock_track_var, command=self._on_lock_track_toggle,
            font=ctk.CTkFont(size=10), text_color=self._C_TEXT_DIM,
            switch_width=28, switch_height=14,
            button_color=self._C_ACCENT, fg_color=self._C_SURFACE2,
            progress_color=self._C_ACCENT)
        _lock_sw.pack(anchor="w", pady=(4, 0))
        Tooltip(_lock_sw, tr("Stop YOLO from re-anchoring the tracker. Draw your box on a "
                          "stable, high-contrast spot (a plug/electrode is ideal) and it "
                          "stays put. Turn off to let auto-reacquire find the head again."))

        # Active Edging — auto-ramp the volume back up after you hold off the edge,
        # then deny again on the next edge. The aggressive auto-edging loop.
        self._active_edging_var = tk.BooleanVar(value=self._active_edging)
        _ae_sw = ctk.CTkSwitch(
            vid_col, text=tr("🔁 Active Edging (auto-ramp)"),
            variable=self._active_edging_var, command=self._on_active_edging_toggle,
            font=ctk.CTkFont(size=10), text_color=self._C_TEXT_DIM,
            switch_width=28, switch_height=14,
            button_color=self._C_ACCENT, fg_color=self._C_SURFACE2,
            progress_color=self._C_ACCENT)
        _ae_sw.pack(anchor="w", pady=(2, 0))
        Tooltip(_ae_sw, tr("After you back off the edge and hold for the configured time "
                        "(Settings → Edging behaviour), volume ramps back toward your max "
                        "until you edge again — then the cycle repeats. Off = normal reactive mode."))

        # height buttons — stacked right of video
        hbf = ctk.CTkFrame(top_frame, fg_color="transparent", width=180, height=260)
        hbf.pack(side=tk.LEFT, fill=tk.Y, padx=(10, 0))
        hbf.pack_propagate(False)

        def _hbtn_row(parent, text, set_cmd, pick_cmd, color, hover):
            row = ctk.CTkFrame(parent, fg_color="transparent")
            row.pack(fill=tk.X, pady=1)
            row.columnconfigure(0, weight=1, uniform="hbtn")
            row.columnconfigure(1, weight=1, uniform="hbtn")
            ctk.CTkButton(row, text=f"📷 {tr(text)}", command=set_cmd, font=btn, height=28,
                          fg_color=color, hover_color=hover,
                          text_color="white", corner_radius=4
                          ).grid(row=0, column=0, sticky="ew", padx=(0, 2))
            ctk.CTkButton(row, text=tr("Manual \u271a"), command=pick_cmd, font=ctk.CTkFont(size=10),
                          height=28,
                          fg_color=color, hover_color=hover,
                          text_color="white", corner_radius=4
                          ).grid(row=0, column=1, sticky="ew")
            return row

        self._auto_btn = ctk.CTkButton(
            hbf, text=tr("📷 AUTO  (observing...)"), command=self._toggle_auto,
            font=btn, height=28, corner_radius=4,
            fg_color=self._C_ACCENT, hover_color=self._C_ACCENT_H, text_color="white")
        self._auto_btn.pack(fill=tk.X, pady=(0, 3))
        Tooltip(self._auto_btn, tr("Auto-calibrate heights by observing motion range"))

        # Exclusion zone controls
        ez_row = ctk.CTkFrame(hbf, fg_color="transparent")
        ez_row.pack(fill=tk.X, pady=(0, 4))
        self._ez_add_btn = ctk.CTkButton(
            ez_row, text=tr("＋ Exclusion Zone"), command=self._add_exclusion_zone,
            font=ctk.CTkFont(size=10), height=26, corner_radius=4,
            fg_color=self._C_ACCENT, hover_color=self._C_ACCENT_H, text_color="black")
        self._ez_add_btn.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=(0, 4))
        Tooltip(self._ez_add_btn, tr("Draw a rectangle on the video feed — YOLO reanchor detections inside it will be ignored"))
        def _clear_zones():
            self._exclusion_zones.clear()
            self._save_config()
            log.info("All exclusion zones cleared")
        ctk.CTkButton(
            ez_row, text=tr("✕ Clear"), command=_clear_zones,
            font=ctk.CTkFont(size=10), height=26, corner_radius=4,
            fg_color=self._C_SURFACE2, hover_color="#4a4a4a",
            text_color=self._C_TEXT, border_width=1, border_color=self._C_BORDER,
            width=60).pack(side=tk.LEFT)
        Tooltip(ez_row, tr("Remove all exclusion zones"))

        # ── Point of No Return (optional yellow panic line) ──────────────────
        # On screen this line sits ABOVE the red edging line, so its button sits
        # just above the red Edging button here too. Unlike the three calibration
        # lines below it is OPTIONAL — nothing depends on it being set. Crossing
        # ABOVE it instantly cuts volume to the floor: a hard "don't let me finish"
        # stop for when the reactive ramp is too slow.
        ctk.CTkLabel(hbf, text=tr("Optional cut-off"), anchor="w",
                     font=ctk.CTkFont(size=9), text_color=self._C_TEXT_DIM
                     ).pack(fill=tk.X, pady=(1, 0))
        ponr_row = ctk.CTkFrame(hbf, fg_color="transparent")
        ponr_row.pack(fill=tk.X, pady=(0, 1))
        ponr_row.columnconfigure(0, weight=3, uniform="ponr")
        ponr_row.columnconfigure(1, weight=2, uniform="ponr")
        ponr_row.columnconfigure(2, weight=1, uniform="ponr")
        ctk.CTkButton(ponr_row, text=tr("⚡ P.O.N.R."), command=self._set_ponr,
                      font=btn, height=28, fg_color=self._C_YELLOW, hover_color=self._C_YELLOW_H,
                      text_color="black", corner_radius=4
                      ).grid(row=0, column=0, sticky="ew", padx=(0, 2))
        ctk.CTkButton(ponr_row, text=tr("Manual ✚"), command=lambda: self._start_pick("PONR"),
                      font=ctk.CTkFont(size=10), height=28,
                      fg_color=self._C_YELLOW, hover_color=self._C_YELLOW_H,
                      text_color="black", corner_radius=4
                      ).grid(row=0, column=1, sticky="ew", padx=(0, 2))
        ctk.CTkButton(ponr_row, text="✕", command=self._clear_ponr,
                      font=ctk.CTkFont(size=12), height=28,
                      fg_color=self._C_SURFACE2, hover_color="#4a4a4a",
                      text_color=self._C_TEXT, border_width=1, border_color=self._C_BORDER,
                      corner_radius=4
                      ).grid(row=0, column=2, sticky="ew")
        Tooltip(ponr_row,
                tr("OPTIONAL hard cut-off — you don't need this set. Place a yellow line "
                "above the red edging line; cross ABOVE it and volume instantly drops to "
                "your floor, so a sudden spike can't finish you before the normal ramp "
                "reacts. Leave it unset (or push it to the very top) to disable. ✕ clears it."))

        # ── The three calibration lines — ALL REQUIRED for the tool to run ───
        ctk.CTkLabel(hbf, text=tr("Calibration — set all 3 (required)"), anchor="w",
                     font=ctk.CTkFont(size=9, weight="bold"), text_color=self._C_TEXT_DIM
                     ).pack(fill=tk.X, pady=(4, 0))
        _hbtn_row(hbf, "Edging",  self._set_edging,  lambda: self._start_pick("Edging"),  self._C_RED,    self._C_RED_HOV)
        _hbtn_row(hbf, "Erect",   self._set_erect,   lambda: self._start_pick("Erect"),   self._C_GREEN,  self._C_GREEN_H)
        _hbtn_row(hbf, "Flaccid", self._set_flaccid, lambda: self._start_pick("Flaccid"), self._C_BLUE,   self._C_BLUE_H)

        # cum buttons anchored at bottom
        # Auto-cum toggle (packs at very bottom, below buttons)
        _acd_row = ctk.CTkFrame(hbf, fg_color="transparent")
        _acd_row.pack(side=tk.BOTTOM, fill=tk.X, pady=(0, 1))
        self._auto_cum_var = tk.BooleanVar(value=False)
        ctk.CTkSwitch(
            _acd_row, text=tr("Auto-detect cum"),
            variable=self._auto_cum_var,
            command=self._on_auto_cum_toggle,
            font=ctk.CTkFont(size=9), text_color=self._C_TEXT_DIM,
            switch_width=28, switch_height=14,
            button_color=self._C_ACCENT, fg_color=self._C_SURFACE2,
        ).pack(side=tk.LEFT, padx=(4, 0))

        cum_row = ctk.CTkFrame(hbf, fg_color="transparent")
        cum_row.pack(side=tk.BOTTOM, fill=tk.X, pady=1)
        cum_row.columnconfigure(0, weight=1, uniform="cumbtn")
        cum_row.columnconfigure(1, weight=1, uniform="cumbtn")
        self._letmecum_btn = ctk.CTkButton(
            cum_row, text=tr("Let me cum?"), command=self._on_letmecum,
            font=btn, height=50, corner_radius=4,
            fg_color=self._C_GREEN, hover_color=self._C_GREEN_H, text_color="white")
        self._letmecum_btn.grid(row=0, column=0, sticky="ew", padx=(0, 2))
        Tooltip(self._letmecum_btn, tr("Roll the dice — odds depend on aggressiveness. Win = temporary full volume permission"))
        self._cum_btn = ctk.CTkButton(cum_row, text=tr("I've CUM"), command=self._on_cum,
                      font=btn, height=50, corner_radius=4,
                      fg_color="#e0e0e8", hover_color="#c8c8d0",
                      text_color=self._C_BG)
        self._cum_btn.grid(row=0, column=1, sticky="ew")
        Tooltip(self._cum_btn, tr("Press after finishing — volume drops to 0 and stays there. Press again (as 'Resume') if you want to go another round."))

        # ── Vertical volume range slider — right of hbf ───────────────────────
        _tiny = ctk.CTkFont(size=9)
        vol_col = ctk.CTkFrame(top_frame, fg_color=self._C_SURFACE, corner_radius=8, width=40)
        vol_col.pack(side=tk.LEFT, fill=tk.Y, padx=(4, 0))
        vol_col.pack_propagate(False)

        ctk.CTkLabel(vol_col, text=tr("VOL\nRANGE"), font=_tiny,
                     text_color=self._C_TEXT_DIM, justify="center").pack(pady=(6, 0))
        ctk.CTkLabel(vol_col, text=tr("ceil"), font=_tiny,
                     text_color=self._C_TEXT_DIM).pack()

        _track_w_v = 4
        _handle_r  = 7
        _cv_w      = _handle_r * 2 + 16   # canvas width (handle + triangle marker)
        self._range_cv = tk.Canvas(vol_col, width=_cv_w, bg=self._C_SURFACE,
                                   highlightthickness=0)
        self._range_cv.pack(fill=tk.Y, expand=True, padx=4, pady=2)

        ctk.CTkLabel(vol_col, text=tr("floor"), font=_tiny,
                     text_color=self._C_TEXT_DIM).pack(pady=(0, 2))
        self._range_lbl = ctk.CTkLabel(vol_col, text="0%\n–\n100%", font=_tiny,
                                       text_color=self._C_YELLOW, wraplength=40, justify="center")
        self._range_lbl.pack(pady=(0, 6))

        self._range_drag = None  # 'lo' or 'hi'
        _drag_tip = [None]   # [Toplevel | None]

        def _show_drag_tip(x_root, y_root, text):
            if _drag_tip[0] is None or not _drag_tip[0].winfo_exists():
                tip = tk.Toplevel(self.root)
                tip.overrideredirect(True)
                tip.attributes("-topmost", True)
                lbl = tk.Label(tip, text=text,
                               bg="#222222", fg="#ffffff",
                               font=("Segoe UI", 10, "bold"),
                               padx=8, pady=4, relief="flat")
                lbl.pack()
                _drag_tip[0] = tip
            else:
                _drag_tip[0].winfo_children()[0].configure(text=text)
            _drag_tip[0].geometry(f"+{x_root + 18}+{y_root - 14}")

        def _hide_drag_tip():
            if _drag_tip[0] is not None and _drag_tip[0].winfo_exists():
                _drag_tip[0].destroy()
            _drag_tip[0] = None

        def _range_draw(event=None):
            c = self._range_cv
            c.delete("all")
            h = c.winfo_height()
            if h < 20:
                return
            pad = _handle_r + 2
            track_h = h - pad * 2
            cx = c.winfo_width() // 2
            lo = self.min_vol_var.get() / 100.0
            hi = self.max_vol_var.get() / 100.0
            # Y=pad → 100%, Y=pad+track_h → 0%
            lo_y = pad + (1.0 - lo) * track_h
            hi_y = pad + (1.0 - hi) * track_h
            # background track
            c.create_line(cx, pad, cx, pad + track_h, fill=self._C_SURFACE2,
                          width=_track_w_v, capstyle="round")
            # active range
            c.create_line(cx, hi_y, cx, lo_y, fill=self._C_ACCENT,
                          width=_track_w_v, capstyle="round")
            # handles
            for y in (lo_y, hi_y):
                c.create_oval(cx - _handle_r, y - _handle_r,
                              cx + _handle_r, y + _handle_r,
                              fill=self._C_ACCENT, outline="")
            # cum release marker — right-pointing triangle beside the track
            cum_frac = 1.0 if self._cum_override_range else hi
            cum_y = pad + (1.0 - cum_frac) * track_h
            ts = 5
            c.create_polygon(
                cx + _handle_r + 2, cum_y - ts,
                cx + _handle_r + 2, cum_y + ts,
                cx + _handle_r + ts + 3, cum_y,
                fill="#3EC941", outline="",
            )

        def _range_press(e):
            h = self._range_cv.winfo_height()
            pad = _handle_r + 2
            track_h = h - pad * 2
            if track_h < 1:
                return
            frac = max(0.0, min(1.0, 1.0 - (e.y - pad) / track_h))
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
            h = self._range_cv.winfo_height()
            pad = _handle_r + 2
            track_h = h - pad * 2
            if track_h < 1:
                return
            frac = max(0.0, min(1.0, 1.0 - (e.y - pad) / track_h))
            val = round(frac * 100)
            if self._range_drag == 'lo':
                val = min(val, int(self.max_vol_var.get()))
                self.min_vol_var.set(val)
            else:
                val = max(val, int(self.min_vol_var.get()))
                self.max_vol_var.set(val)
            lo = int(self.min_vol_var.get())
            hi = int(self.max_vol_var.get())
            self._range_lbl.configure(text=f"{lo}%\n–\n{hi}%")
            _range_draw()
            if self._range_drag == 'lo':
                _show_drag_tip(e.x_root, e.y_root, f"Min Vol ({lo}%)")
            else:
                _show_drag_tip(e.x_root, e.y_root, f"Max Vol ({hi}%)")

        def _range_release(e):
            self._range_drag = None
            _hide_drag_tip()
            self._save_config()

        self._range_cv.bind("<ButtonPress-1>", _range_press)
        self._range_cv.bind("<B1-Motion>", _range_move)
        self._range_cv.bind("<ButtonRelease-1>", _range_release)
        self._range_cv.bind("<Configure>", _range_draw)
        self._range_draw = _range_draw
        Tooltip(self._range_cv,
               tr("Drag handles to set volume floor (bottom) and ceiling (top).\n"
               "Green triangle = where 'Let me cum?' sends volume.\n"
               "Change override behavior in Settings > Cum Volume Behavior."))

        def _divider():
            ctk.CTkFrame(sf, height=1, fg_color=self._C_BORDER).pack(fill=tk.X, padx=P, pady=1)

        # ── Aggressiveness ────────────────────────────────────────────────────
        aggr_card = ctk.CTkFrame(sf, fg_color=self._C_SURFACE, corner_radius=8)
        aggr_card.pack(fill=tk.X, padx=P, pady=2)
        Tooltip(aggr_card,
               tr("Controls how fast volume ramps and 'Let me cum?' odds.\n"
               "Easy: gentle ramps, 1-in-2 odds, 5 min grant\n"
               "Middle: moderate ramps, 1-in-4 odds, 5 min grant\n"
               "Hard: fast ramps, 1-in-6 odds, 3 min grant\n"
               "Expert: aggressive ramps, 1-in-30 odds, 1 min grant"))
        aggr_row = ctk.CTkFrame(aggr_card, fg_color="transparent")
        aggr_row.pack(fill=tk.X, padx=12, pady=6)
        ctk.CTkLabel(aggr_row, text=tr("Aggressiveness"), font=lbl,
                     text_color=self._C_TEXT).pack(side=tk.LEFT)
        self._evil_btn = ctk.CTkButton(
            aggr_row, text=tr("EVIL MODE"), command=self._toggle_evil_mode,
            font=ctk.CTkFont(size=12), height=28, width=36, corner_radius=4,
            fg_color="transparent", hover_color="#3a0000",
            border_width=2, border_color="#cc0000", text_color="#cc0000",
        )
        self._evil_btn.pack(side=tk.RIGHT, padx=(4, 0))
        Tooltip(self._evil_btn, "Evil Mode — adds ruin outcome, crimson theme, and devil.png overlay")
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
        mode_card = ctk.CTkFrame(sf, fg_color=self._C_SURFACE, corner_radius=8)
        self._mode_card = mode_card
        mode_card.pack(fill=tk.X, padx=P, pady=2)
        mode_row = ctk.CTkFrame(mode_card, fg_color="transparent")
        mode_row.pack(fill=tk.X, padx=12, pady=6)
        ctk.CTkLabel(mode_row, text=tr("Output"), font=lbl,
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
        self._restim_opts = ctk.CTkFrame(sf, fg_color=self._C_SURFACE2, corner_radius=6)
        ro = ctk.CTkFrame(self._restim_opts, fg_color="transparent")
        ro.pack(padx=12, pady=7)
        ctk.CTkLabel(ro, text=tr("Restim Port:"), font=lbl, text_color=self._C_TEXT).pack(side=tk.LEFT, padx=(0, 6))
        ctk.CTkEntry(ro, textvariable=self.port_var, width=72,
                     fg_color=self._C_SURFACE, border_color=self._C_BORDER,
                     text_color=self._C_TEXT).pack(side=tk.LEFT)
        self.port_var.trace_add("write", self._on_port_change)
        # OBS overlay URL on same row
        _overlay_url = f"http://127.0.0.1:{OverlayServer.PORT}"
        ctk.CTkLabel(ro, text=tr("OBS Overlay:"), font=lbl,
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
            _copy_btn.configure(text=tr("Copied!"))
            self.root.after(1500, lambda: _copy_btn.configure(text=tr("Copy")))
        _copy_btn = ctk.CTkButton(ro, text=tr("Copy"), command=_copy_overlay_url, width=48,
                                  fg_color=self._C_SURFACE, hover_color="#4a4a4a",
                                  text_color=self._C_TEXT, border_width=1, border_color=self._C_BORDER,
                                  font=ctk.CTkFont(size=10))
        _copy_btn.pack(side=tk.LEFT)

        # xToys options panel
        self._xtoys_opts = ctk.CTkFrame(sf, fg_color=self._C_SURFACE2, corner_radius=6)
        xo = ctk.CTkFrame(self._xtoys_opts, fg_color="transparent")
        xo.pack(fill=tk.X, padx=12, pady=(7, 2))
        ctk.CTkLabel(xo, text=tr("Webhook ID:"), font=lbl,
                     text_color=self._C_TEXT).pack(side=tk.LEFT, padx=(0, 6))
        self._xtoys_entry = ctk.CTkEntry(xo, textvariable=self.xtoys_id_var, width=240,
                     fg_color=self._C_SURFACE, border_color=self._C_BORDER,
                     text_color=self._C_TEXT,
                     placeholder_text=tr("e.g. 8hR5acKTCx2s"))
        self._xtoys_entry.pack(side=tk.LEFT, padx=(0, 8))
        # One-click shortcut straight to the Private Webhook page \u2014 saves users hunting
        # for xtoys.app/me. Mirrors the About-menu link buttons (webbrowser.open).
        ctk.CTkButton(xo, text=tr("Get ID \u2197"), width=64, height=24,
                      command=lambda: webbrowser.open("https://xtoys.app/me"),
                      fg_color=self._C_SURFACE, hover_color="#4a4a4a",
                      text_color=self._C_TEXT, border_width=1, border_color=self._C_BORDER,
                      font=ctk.CTkFont(size=10)).pack(side=tk.LEFT)
        ctk.CTkLabel(self._xtoys_opts,
                     text=tr("In xToys: load the VisualStimEdger script \u2192 add your toy under Connections \u2192 Generic Output, then press \u25b6 to run it (nothing moves until it runs). Tap \u201cGet ID\u201d to open xtoys.app/me \u2192 Private Webhook, then paste the ID here."),
                     font=ctk.CTkFont(size=9), text_color=self._C_TEXT_DIM,
                     wraplength=380, justify="left").pack(anchor="w", padx=12, pady=(0, 6))
        self._xtoys_warn = ctk.CTkLabel(self._xtoys_opts, text="",
                     font=ctk.CTkFont(size=9, weight="bold"), text_color="#ff5555",
                     wraplength=380, justify="left")
        self._xtoys_warn_shown = False   # shown by _on_xtoys_id_change when the ID looks wrong
        self.xtoys_id_var.trace_add("write", self._on_xtoys_id_change)

        # ── Heart Rate (Pulsoid) options panel ────────────────────────────────
        self._hr_opts = ctk.CTkFrame(sf, fg_color=self._C_SURFACE2, corner_radius=6)

        # HR source selector
        hr_src_row = ctk.CTkFrame(self._hr_opts, fg_color="transparent")
        hr_src_row.pack(fill=tk.X, padx=12, pady=(7, 2))
        ctk.CTkLabel(hr_src_row, text=tr("HR Source:"), font=lbl,
                     text_color=self._C_TEXT).pack(side=tk.LEFT, padx=(0, 8))
        self._hr_src_seg = ctk.CTkSegmentedButton(
            hr_src_row, values=["Pulsoid", "BLE Direct"],
            command=self._on_hr_source_change,
            selected_color=self._C_ACCENT, selected_hover_color=self._C_ACCENT_H,
            unselected_color=self._C_SURFACE, unselected_hover_color="#4a4a4a",
            font=ctk.CTkFont(size=11), text_color=self._C_TEXT,
        )
        self._hr_src_seg.set("Pulsoid")
        self._hr_src_seg.pack(side=tk.LEFT)

        hr_row1 = ctk.CTkFrame(self._hr_opts, fg_color="transparent")
        hr_row1.pack(fill=tk.X, padx=12, pady=(7, 2))
        ctk.CTkLabel(hr_row1, text=tr("Pulsoid Token:"), font=lbl,
                     text_color=self._C_TEXT).pack(side=tk.LEFT, padx=(0, 6))
        ctk.CTkEntry(hr_row1, textvariable=self.hr_token_var, width=210,
                     fg_color=self._C_SURFACE, border_color=self._C_BORDER,
                     text_color=self._C_TEXT, show="\u2022",
                     placeholder_text=tr("pulsoid.net/ui/keys (paid plan required)")).pack(side=tk.LEFT, padx=(0, 8))
        self._hr_bpm_label = ctk.CTkLabel(hr_row1, text=tr("\u2665 -- bpm"),
                                          font=ctk.CTkFont(size=fs + 1, weight="bold"),
                                          text_color="#e91e63")
        self._hr_bpm_label.pack(side=tk.LEFT)

        hr_row2 = ctk.CTkFrame(self._hr_opts, fg_color="transparent")
        hr_row2.pack(fill=tk.X, padx=12, pady=(0, 2))

        def _hr_slider_group(parent, label, var, lo, hi):
            ctk.CTkLabel(parent, text=label, font=lbl,
                         text_color=self._C_TEXT).pack(side=tk.LEFT, padx=(0, 4))
            lbl_val = ctk.CTkLabel(parent, text=tr("{} bpm").format(var.get()),
                                   font=lbl, text_color=self._C_TEXT_DIM, width=60)
            lbl_val.pack(side=tk.LEFT, padx=(0, 6))
            def _upd(v, lv=lbl_val, sv=var):
                sv.set(int(float(v)))
                lv.configure(text=tr("{} bpm").format(int(float(v))))
                self.hr_client.resting_bpm = self.hr_resting_var.get()
                self.hr_client.peak_bpm    = self.hr_peak_var.get()
            ctk.CTkSlider(parent, from_=lo, to=hi, variable=var,
                          command=_upd, width=110,
                          button_color="#e91e63", button_hover_color="#c2185b",
                          progress_color="#e91e63").pack(side=tk.LEFT)

        _hr_slider_group(hr_row2, "Resting:", self.hr_resting_var, 40, 90)
        ctk.CTkFrame(hr_row2, width=16, fg_color="transparent").pack(side=tk.LEFT)
        _hr_slider_group(hr_row2, "Peak:", self.hr_peak_var, 80, 170)

        # BLE Direct row — shown when BLE source selected
        self._ble_row = ctk.CTkFrame(self._hr_opts, fg_color="transparent")
        # (packed/unpacked by _on_hr_source_change)
        ble_inner = ctk.CTkFrame(self._ble_row, fg_color="transparent")
        ble_inner.pack(fill=tk.X, padx=12, pady=(2, 7))
        self._ble_name_lbl = ctk.CTkLabel(ble_inner, text=tr("No device selected"),
                                           font=lbl, text_color=self._C_TEXT_DIM)
        self._ble_name_lbl.pack(side=tk.LEFT, expand=True, anchor="w")
        ctk.CTkButton(ble_inner, text=tr("Scan…"), width=70, height=26,
                      fg_color=self._C_ACCENT, hover_color=self._C_ACCENT_H,
                      text_color="white", font=ctk.CTkFont(size=11),
                      command=self._ble_scan_dialog).pack(side=tk.RIGHT)

        ctk.CTkLabel(self._hr_opts,
                     text=tr("Higher HR \u2192 stronger denial, slower rewards.  "
                           "Get a free token at pulsoid.net/ui/keys \u2014 works with Polar, Garmin, "
                           "Apple Watch, most BLE chest straps via the Pulsoid app."),
                     font=ctk.CTkFont(size=9), text_color=self._C_TEXT_DIM,
                     wraplength=390, justify="left").pack(anchor="w", padx=12, pady=(2, 6))
        self.hr_token_var.trace_add("write", self._on_hr_token_change)

        # Windows Audio options panel
        self._windows_opts = ctk.CTkFrame(sf, fg_color=self._C_SURFACE2, corner_radius=6)
        wo = ctk.CTkFrame(self._windows_opts, fg_color="transparent")
        wo.pack(padx=12, pady=7, fill=tk.X)
        ctk.CTkLabel(wo, text=tr("Device:"), font=lbl, text_color=self._C_TEXT).pack(side=tk.LEFT, padx=(0, 6))
        self._device_combo = ctk.CTkComboBox(
            wo, values=[], command=self._on_device_select, width=280,
            fg_color=self._C_SURFACE, border_color=self._C_BORDER,
            button_color=self._C_RED, button_hover_color=self._C_RED_HOV,
            dropdown_fg_color=self._C_SURFACE2, text_color=self._C_TEXT,
        )
        self._device_combo.pack(side=tk.LEFT, padx=(0, 6))
        ctk.CTkButton(wo, text=tr("Refresh"), command=self._refresh_devices, width=72,
                      fg_color=self._C_SURFACE2, hover_color="#4a4a4a",
                      text_color=self._C_TEXT, border_width=1, border_color=self._C_BORDER,
                      font=ctk.CTkFont(size=10)).pack(side=tk.LEFT)
        # Soft promo: users on audio output (electro) are exactly who'd want the pattern packs.
        _js_btn = ctk.CTkButton(wo, text=tr("⚡ Free stim patterns ↗"), width=168, height=24,
                      fg_color=self._C_SURFACE, hover_color="#4a4a4a",
                      text_color="#F5A623", border_width=1, border_color="#F5A623",
                      font=ctk.CTkFont(size=10, weight="bold"),
                      command=lambda: webbrowser.open("https://ko-fi.com/s/db90a470a9"))
        _js_btn.pack(side=tk.RIGHT)
        Tooltip(_js_btn, tr("Sir Thorn's Sunday Drive sessions as JSON stim patterns — free / "
                         "pay-what-you-want. Made for electro over your audio output. Opens Ko-fi."))

        self._restim_opts.pack(fill=tk.X, padx=P, pady=(0, 4))  # default; _on_output_change re-packs after load

        # ── MP3 player options panel ───────────────────────────────────────────
        self._mp3_opts = ctk.CTkFrame(sf, fg_color=self._C_SURFACE2, corner_radius=6)
        if _MINIAUDIO_OK:
            mo = ctk.CTkFrame(self._mp3_opts, fg_color="transparent")
            mo.pack(fill=tk.X, padx=12, pady=(8, 4))

            # File / folder load buttons
            load_row = ctk.CTkFrame(mo, fg_color="transparent")
            load_row.pack(fill=tk.X, pady=(0, 4))
            ctk.CTkButton(load_row, text=tr("📁 Load File"), width=110, height=28,
                          fg_color=self._C_SURFACE, hover_color="#4a4a4a",
                          text_color=self._C_TEXT, border_width=1, border_color=self._C_BORDER,
                          font=ctk.CTkFont(size=10),
                          command=self._mp3_load_file).pack(side=tk.LEFT, padx=(0, 6))
            ctk.CTkButton(load_row, text=tr("📂 Load Folder"), width=120, height=28,
                          fg_color=self._C_SURFACE, hover_color="#4a4a4a",
                          text_color=self._C_TEXT, border_width=1, border_color=self._C_BORDER,
                          font=ctk.CTkFont(size=10),
                          command=self._mp3_load_folder).pack(side=tk.LEFT, padx=(0, 10))
            # Soft promo: they're sat here looking for something to load.
            _pv_btn = ctk.CTkButton(load_row, text=tr("🎧 21 tracks to edge to ↗"), width=190, height=28,
                          fg_color=self._C_SURFACE, hover_color="#4a4a4a",
                          text_color="#F5A623", border_width=1, border_color="#F5A623",
                          font=ctk.CTkFont(size=10, weight="bold"),
                          command=lambda: webbrowser.open("https://ko-fi.com/s/ab28a9a250"))
            _pv_btn.pack(side=tk.RIGHT)
            Tooltip(_pv_btn, tr("Sir Thorn's Private Sessions Vol. 2 — 21 MP3 tracks made for edging. "
                             "Load them straight into this player. Opens Ko-fi."))
            self._mp3_track_lbl = ctk.CTkLabel(load_row, text=tr("No file loaded"),
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

            # Output device
            dev_row = ctk.CTkFrame(mo, fg_color="transparent")
            dev_row.pack(fill=tk.X, pady=(0, 4))
            ctk.CTkLabel(dev_row, text=tr("Out:"), font=ctk.CTkFont(size=10),
                         text_color=self._C_TEXT_DIM).pack(side=tk.LEFT, padx=(0, 6))
            _mp3_devs = ["Default"] + (MusicPlayer.list_devices() if _MINIAUDIO_OK else [])
            self._mp3_dev_var = tk.StringVar(value="Default")
            self._mp3_dev_combo = ctk.CTkComboBox(
                dev_row, values=_mp3_devs, variable=self._mp3_dev_var,
                width=220, font=ctk.CTkFont(size=10),
                fg_color=self._C_SURFACE, border_color=self._C_BORDER,
                button_color=self._C_ACCENT, button_hover_color=self._C_ACCENT_H,
                dropdown_fg_color=self._C_SURFACE2, text_color=self._C_TEXT,
                command=self._mp3_on_device_change,
            )
            self._mp3_dev_combo.pack(side=tk.LEFT, fill=tk.X, expand=True)

            # Loop mode
            loop_row = ctk.CTkFrame(mo, fg_color="transparent")
            loop_row.pack(pady=(0, 6))
            ctk.CTkLabel(loop_row, text=tr("Loop:"), font=ctk.CTkFont(size=10),
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
        ctrl = ctk.CTkFrame(sf, fg_color="transparent")
        ctrl.pack(fill=tk.X, padx=P, pady=3)

        def _ghost_btn(parent, text, cmd, **kw):
            kw.setdefault("font", ctk.CTkFont(size=10))
            return ctk.CTkButton(
                parent, text=text, command=cmd, height=34,
                fg_color=self._C_SURFACE2, hover_color="#4a4a4a",
                text_color=self._C_TEXT, border_width=1, border_color=self._C_BORDER,
                **kw,
            )

        self._hold_btn = _ghost_btn(ctrl, tr("Hold Volume"), self._toggle_hold,
                                    font=ctk.CTkFont(size=10, weight="bold"), width=120)
        self._hold_btn.pack(side=tk.LEFT, padx=(4, 0))
        Tooltip(self._hold_btn, tr("Freeze volume at current level — tracking continues but volume won't change"))
        self._about_btn = _ghost_btn(ctrl, "ⓘ", self._show_about_menu, width=34)
        self._about_btn.pack(side=tk.LEFT, padx=(4, 0))
        self._play_btn = _ghost_btn(ctrl, tr("▶ Play"), self._toggle_play_mode, width=60)
        self._play_btn.pack(side=tk.LEFT, padx=(4, 0))
        Tooltip(self._play_btn, tr("Play Mode — minimal immersive view; hides settings"))
        self._bondage_btn = ctk.CTkButton(
            ctrl, text=tr("🎙 BONDAGE"), command=self._open_bondage_splash,
            font=ctk.CTkFont(size=10, weight="bold"), height=34, corner_radius=4,
            fg_color="transparent", hover_color="#1a0028",
            border_width=2, border_color="#8a2a9a", text_color="#c080ff", width=90,
        )
        self._bondage_btn.pack(side=tk.LEFT, padx=(4, 0))
        Tooltip(self._bondage_btn, tr("Bondage Mode — hands-free voice control (requires Vosk model)"))

        _divider()

        # ── Status labels ─────────────────────────────────────────────────────
        self.info_label = ctk.CTkLabel(
            sf, text=tr("State: --  |  Vol: --  |  WS: Disconnected"),
            font=ctk.CTkFont(size=15, weight="bold"), text_color=self._C_TEXT,
        )
        self.info_label.pack(pady=(4, 1))

        # Contextual "what to do" line — blank when healthy, speaks up on a connection
        # problem (populated in _update_status_label). Kept separate from the busy
        # italic snark line below the video so the two don't fight over one widget.
        self._help_label = ctk.CTkLabel(
            sf, text="", font=ctk.CTkFont(size=11), text_color="#F5A623",
            wraplength=560, justify="center",
        )
        self._help_label.pack(pady=(0, 1))

        self.stats_label = ctk.CTkLabel(
            sf, text=tr("Session: 00:00  |  Edges: 0"),
            font=ctk.CTkFont(size=10), text_color=self._C_TEXT_DIM,
        )
        self.stats_label.pack(pady=(0, 4))

        self._build_play_panel(root)

    def _build_play_panel(self, parent):
        """Minimal immersive play-mode panel — replaces the scrollable settings."""
        P = 12
        pp = ctk.CTkFrame(parent, fg_color=self._C_BG, corner_radius=0)
        # NOT packed here — _toggle_play_mode handles show/hide
        self._play_panel = pp

        # ── Big state label ───────────────────────────────────────────────────
        self._play_state_lbl = ctk.CTkLabel(
            pp, text="—",
            font=ctk.CTkFont(size=52, weight="bold"),
            text_color=self._C_TEXT)
        self._play_state_lbl.pack(pady=(20, 2))

        # ── Volume ────────────────────────────────────────────────────────────
        self._play_vol_lbl = ctk.CTkLabel(
            pp, text=tr("Vol: —"),
            font=ctk.CTkFont(size=20, weight="bold"),
            text_color=self._C_TEXT_DIM)
        self._play_vol_lbl.pack(pady=(0, 8))

        # ── Stats card: edges | time | HR ────────────────────────────────────
        stats_card = ctk.CTkFrame(pp, fg_color=self._C_SURFACE, corner_radius=8)
        stats_card.pack(fill=tk.X, padx=P, pady=4)
        self._play_edges_lbl = ctk.CTkLabel(
            stats_card, text=tr("0 edges"),
            font=ctk.CTkFont(size=13, weight="bold"), text_color=self._C_TEXT)
        self._play_edges_lbl.pack(side=tk.LEFT, padx=16, pady=10)
        self._play_time_lbl = ctk.CTkLabel(
            stats_card, text="0:00",
            font=ctk.CTkFont(size=13, weight="bold"), text_color=self._C_TEXT)
        self._play_time_lbl.pack(side=tk.LEFT, padx=16, pady=10)
        self._play_hr_lbl = ctk.CTkLabel(
            stats_card, text="",
            font=ctk.CTkFont(size=13), text_color="#e91e63")
        self._play_hr_lbl.pack(side=tk.RIGHT, padx=16, pady=10)

        # ── Snark / status line ───────────────────────────────────────────────
        self._play_snark_lbl = ctk.CTkLabel(
            pp, text="",
            font=ctk.CTkFont(size=11, slant="italic"),
            text_color=self._C_TEXT_DIM)
        self._play_snark_lbl.pack(pady=(2, 6))

        # ── Action buttons ────────────────────────────────────────────────────
        def _pbtn(parent, text, cmd, fg, hov):
            return ctk.CTkButton(
                parent, text=text, command=cmd,
                font=ctk.CTkFont(size=13, weight="bold"),
                height=50, corner_radius=6,
                fg_color=fg, hover_color=hov, text_color="white")

        row1 = ctk.CTkFrame(pp, fg_color="transparent")
        row1.pack(fill=tk.X, padx=P, pady=(0, 4))
        self._play_hold_btn = _pbtn(row1, tr("Hold Volume"), self._toggle_hold,
                                    self._C_SURFACE2, "#4a4a4a")
        self._play_hold_btn.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=(0, 4))
        self._play_letmecum_btn = _pbtn(row1, "Let me cum?", self._on_letmecum,
                                        self._C_GREEN, self._C_GREEN_H)
        self._play_letmecum_btn.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=(4, 0))

        row2 = ctk.CTkFrame(pp, fg_color="transparent")
        row2.pack(fill=tk.X, padx=P, pady=(0, 4))
        self._play_cum_btn = _pbtn(row2, "I've CUM", self._on_cum, "#6a6a7a", "#555565")
        self._play_cum_btn.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=(0, 4))
        _pbtn(row2, "⚙ Settings", self._toggle_play_mode,
              self._C_SURFACE2, "#4a4a4a").pack(
            side=tk.LEFT, expand=True, fill=tk.X, padx=(4, 0))

        # Auto-detect toggle in play mode
        _play_acd_row = ctk.CTkFrame(pp, fg_color="transparent")
        _play_acd_row.pack(fill=tk.X, padx=P, pady=(0, 8))
        ctk.CTkSwitch(
            _play_acd_row, text=tr("Auto-detect cum"),
            variable=self._auto_cum_var,
            command=self._on_auto_cum_toggle,
            font=ctk.CTkFont(size=9), text_color=self._C_TEXT_DIM,
            switch_width=28, switch_height=14,
            button_color=self._C_ACCENT, fg_color=self._C_SURFACE2,
        ).pack(side=tk.LEFT)

        # Size window to content once everything is laid out
        parent.after(100, self._fit_window)

    def _start_update_check(self):
        def callback(latest, url):
            self._call_on_main(self._show_update_banner, latest, url)
        threading.Thread(target=check_for_update, args=(callback,), daemon=True).start()

    # ------------------------------------------------------------------ config persistence

    def _save_config(self):
        try:
            CONFIG_PATH.parent.mkdir(parents=True, exist_ok=True)
            data = {
                "heights":     self.heights,
                "win_w":       self._win_w,
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
                "hr_source":   self._hr_source,
                "ble_addr":    self._ble_addr,
                "ble_name":    self._ble_name,
                "port":        self.port_var.get(),
                "device_name": self._device_combo.get(),
                "mp3_path":    getattr(self, '_mp3_last_path', ""),
                "mp3_path_type": getattr(self, '_mp3_last_type', "file"),
                "mp3_loop":    self._mp3_loop_var.get() if _MINIAUDIO_OK else "folder",
                "cum_odds":    self._cum_odds,
                "denial_phrases": self._denial_phrases,
                "cum_override_range": self._cum_override_range,
                "refractory_mins":   self._refractory_mins,
                "ui_font_size": self._ui_font_size,
                "theme": self._theme_name,
                "language": self._language,
                "ruin_odds": self._ruin_odds,
                "ruin_phrases": self._ruin_phrases,
                "exclusion_zones":      self._exclusion_zones,
                "reanchor_lock":        self._reanchor_lock,
                "color_lock":           self._color_lock,
                "tracker_backend":      self._tracker_backend,
                "proc_width":           self._proc_width,
                "tribute_last_shown":    getattr(self, '_tribute_last_shown', 0.0),
                "tribute_session_count": getattr(self, '_tribute_session_count', 0),
                "erect_recovery":       self._erect_recovery,
                "active_edging":        self._active_edging,
                "recovery_pct":         self._recovery_pct,
                "ae_hold_secs":         self._ae_hold_secs,
                "ae_ramp_pct":          self._ae_ramp_pct,
                "auto_cum_enabled":     self._auto_cum_enabled,
                "auto_cum_delay":       self._auto_cum_delay,
                "auto_cum_sensitivity": self._auto_cum_sensitivity,
                "hf_enabled":           self._hf_enabled,
                "hf_min_edges":         self._hf_min_edges,
                "hf_cum_chance":        self._hf_cum_chance,
                "bondage_safeword":     self._bondage_safeword,
                "bondage_mic_device":   self._bondage_mic_device,
            }
            # Atomic write: a crash mid-write (this app has hard-killed before)
            # must not truncate config.json and wipe every setting. Write a temp
            # file then os.replace (atomic on Windows and POSIX).
            tmp = CONFIG_PATH.with_suffix(".json.tmp")
            tmp.write_text(json.dumps(data, indent=2), encoding="utf-8")
            os.replace(tmp, CONFIG_PATH)
        except Exception as e:
            log.warning(f"Config: save failed: {e}")

    def _load_config(self):
        self._loading_config = True
        try:
            if not CONFIG_PATH.exists():
                return
            data = json.loads(CONFIG_PATH.read_text())
            # Heights
            for key in ("Edging", "Erect", "Flaccid", "PONR"):
                if key in data.get("heights", {}):
                    self.heights[key] = data["heights"][key]
            # Remembered window width — restore whatever the user last sized it to.
            _ww = data.get("win_w")
            if isinstance(_ww, (int, float)) and 100 < _ww <= _MAX_WIN_W:
                self._win_w = int(_ww)
            # else: junk or an absurdly wide legacy value -> ignore it, so
            # _fit_window falls back to the skinny _DEFAULT_WIN_W instead of huge.
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
                        self._sens_lbl.configure(text=tr("{} px").format(int(self.edge_sens_var.get())))
                except Exception:
                    pass
            if "xtoys_id" in data:
                self.xtoys_id_var.set(str(data.get("xtoys_id", "")))
                self.xtoys.webhook_id = self.xtoys_id_var.get()
            self._tribute_last_shown    = float(data.get("tribute_last_shown", 0) or 0)
            self._tribute_session_count = int(data.get("tribute_session_count", 0) or 0)
            if "proc_width" in data:
                try:
                    pw = int(data.get("proc_width", 0) or 0)
                    self._proc_width = pw if pw in self._PROC_PRESETS.values() else 0
                    self.proc_res_var.set(next(
                        (k for k, v in self._PROC_PRESETS.items() if v == self._proc_width), "Full"))
                except Exception:
                    pass
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
            self._hr_source = data.get("hr_source", "pulsoid")
            self._ble_addr  = data.get("ble_addr", "")
            self._ble_name  = data.get("ble_name", "")
            if "port"        in data: self.port_var.set(data["port"])
            if "device_name" in data:
                self._refresh_devices(select_name=data["device_name"])
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
            if "refractory_mins" in data:
                self._refractory_mins = max(0, int(data["refractory_mins"]))
            # Legacy keys cum_silence / cum_ramp are silently ignored.
            if "ui_font_size" in data:
                self._ui_font_size = int(data["ui_font_size"])
            if "theme" in data and data["theme"] in THEMES and data["theme"] != "Evil":
                self._theme_name = data["theme"]
            if isinstance(data.get("language"), str):
                self._language = data["language"]
            if "ruin_odds" in data and isinstance(data["ruin_odds"], dict):
                self._ruin_odds.update(data["ruin_odds"])
            if "ruin_phrases" in data and isinstance(data["ruin_phrases"], list):
                self._ruin_phrases = data["ruin_phrases"]
            if "exclusion_zones" in data and isinstance(data["exclusion_zones"], list):
                self._exclusion_zones = [tuple(z) for z in data["exclusion_zones"]]
            if "reanchor_lock" in data:
                self._reanchor_lock = bool(data["reanchor_lock"])
                if hasattr(self, "_lock_track_var"):
                    self._lock_track_var.set(self._reanchor_lock)
            if "color_lock" in data:
                self._color_lock = bool(data["color_lock"])
                if hasattr(self, "_color_lock_var"):
                    self._color_lock_var.set(self._color_lock)
                if self._color_lock:
                    self._color_needs_derive = True   # re-derive from the box once tracking starts
            if "tracker_backend" in data:
                tb = str(data.get("tracker_backend", "vit")).lower()
                self._tracker_backend = "csrt" if tb == "csrt" else "vit"
                self._tracker_dirty = True            # rebuild to the saved backend next frame
                if hasattr(self, "_tracker_var"):
                    self._tracker_var.set("Classic" if self._tracker_backend == "csrt" else "Deep")
            if "erect_recovery" in data:
                self._erect_recovery = bool(data["erect_recovery"])
            if "active_edging" in data:
                self._active_edging = bool(data["active_edging"])
                if hasattr(self, "_active_edging_var"):
                    self._active_edging_var.set(self._active_edging)
            if isinstance(data.get("recovery_pct"), dict):
                self._recovery_pct.update({k: max(0.0, float(v))
                    for k, v in data["recovery_pct"].items() if k in self._recovery_pct})
            if isinstance(data.get("ae_hold_secs"), dict):
                self._ae_hold_secs.update({k: max(1, min(120, int(v)))
                    for k, v in data["ae_hold_secs"].items() if k in self._ae_hold_secs})
            if isinstance(data.get("ae_ramp_pct"), dict):
                self._ae_ramp_pct.update({k: max(0.0, float(v))
                    for k, v in data["ae_ramp_pct"].items() if k in self._ae_ramp_pct})
            if "auto_cum_enabled" in data:
                self._auto_cum_enabled = bool(data["auto_cum_enabled"])
            if "auto_cum_delay" in data:
                self._auto_cum_delay = max(0, min(10, int(data["auto_cum_delay"])))
            if "auto_cum_sensitivity" in data:
                self._auto_cum_sensitivity = max(1, min(10, int(data["auto_cum_sensitivity"])))
            if "hf_enabled" in data:
                self._hf_enabled = bool(data["hf_enabled"])
            if "hf_min_edges" in data and isinstance(data["hf_min_edges"], dict):
                self._hf_min_edges.update(
                    {k: max(0, int(v)) for k, v in data["hf_min_edges"].items()})
            if "hf_cum_chance" in data and isinstance(data["hf_cum_chance"], dict):
                self._hf_cum_chance.update(
                    {k: max(1, int(v)) for k, v in data["hf_cum_chance"].items()})
            if "bondage_safeword" in data:
                w = str(data["bondage_safeword"]).strip().lower()
                if w:
                    self._bondage_safeword      = w
                    self._bondage_safeword_saved = True
            if "bondage_mic_device" in data:
                self._bondage_mic_device = str(data["bondage_mic_device"])
            log.info(f"Config: loaded from {CONFIG_PATH}")
        except Exception as e:
            log.warning(f"Config: load failed: {e}")
        finally:
            self._loading_config = False

    @property
    def _active_hr(self):
        """Returns whichever HR client is currently selected."""
        return self.ble_hr_client if self._hr_source == "ble" else self.hr_client

    def _cleanup(self):
        """Non-UI cleanup — safe to call from atexit or _on_close."""
        if getattr(self, '_cleaned_up', False):
            return
        self._cleaned_up = True
        self._running = False

        if hasattr(self, 'session_logger'):
            try:
                self.session_logger.log_session_end(
                    edge_count=self.edge_count,
                    cum_count=self._cum_count,
                    denial_count=self._denial_count,
                    elapsed_s=time.time() - self.session_start,
                )
            except Exception:
                pass

        # Wait for capture thread to finish (max 1.5 s)
        t = getattr(self, '_capture_thread', None)
        if t and t.is_alive():
            t.join(timeout=1.5)

        # Zero out Restim before closing socket
        try:
            self.restim.set_volume(0.0, instant=True)
        except Exception:
            pass

        # Zero xToys too — otherwise the toy keeps running at its last level
        # after the app closes. Short timeout so shutdown can't hang.
        try:
            if getattr(self, "xtoys", None):
                self.xtoys.set_volume(0.0, instant=True)
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
        try:
            self.ble_hr_client.stop()
        except Exception:
            pass

        # Stop voice engine (bondage mode)
        if getattr(self, '_voice_engine', None):
            try:
                self._voice_engine.stop()
            except Exception:
                pass
        if getattr(self, '_splash_voice_engine', None):
            try:
                self._splash_voice_engine.stop()
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

    def _flush_critical(self):
        """The fast, must-happen-on-close side effects: stop the toy, restore the
        user's system volume, log the session. Main-thread only — no blocking joins,
        no thread .stop() — so it's safe to run right before os._exit(0). Idempotent."""
        try:
            self.restim.set_volume(0.0, instant=True)
        except Exception:
            pass
        try:
            if getattr(self, "xtoys", None):
                self.xtoys.set_volume(0.0, instant=True)
        except Exception:
            pass
        if self.win_audio and self.win_audio.connected and self._orig_win_volume is not None:
            try:
                self.win_audio._volume_interface.SetMasterVolumeLevelScalar(
                    self._orig_win_volume, None)
                log.info(f"WinAudio: restored volume to {self._orig_win_volume:.2f}")
            except Exception:
                pass
        if hasattr(self, 'session_logger'):
            try:
                self.session_logger.log_session_end(
                    edge_count=self.edge_count, cum_count=self._cum_count,
                    denial_count=self._denial_count,
                    elapsed_s=time.time() - self.session_start)
            except Exception:
                pass

    def _on_close(self):
        # Collapsed close path. The old path ran the full _cleanup(), which blocks up
        # to 1.5s joining the capture thread while Tk's event loop goes unserviced — a
        # background thread touching Tk in that window trips
        #   Tcl_AsyncDelete: async handler deleted by the wrong thread
        # (abort, exit 3), and because the abort beat the zeroing the toy was left
        # running. So do ONLY the fast, must-happen side effects on the main thread,
        # then hard-exit immediately: no joins, no destroy, no teardown window. (And
        # background threads no longer touch Tk at all — see _call_on_main — so there's
        # nothing left to race even in the millisecond before we exit.)
        self._running = False
        try:
            _w = self.root.winfo_width()
            if 200 < _w <= _MAX_WIN_W:
                self._win_w = _w          # persist the window width for next launch
        except Exception:
            pass
        self._flush_critical()            # stop the toy, restore audio, log the session
        try:
            self._save_config()
        except Exception:
            pass
        try:
            logging.shutdown()
        except Exception:
            pass
        os._exit(0)

    # ------------------------------------------------------------------ capture thread

    def _capture_loop(self):
        """Runs in background thread — captures frames and drops them in the queue."""
        while self._running:
            if self.tracking_paused:
                time.sleep(0.05)
                continue
            frame = capture_window_region(self.hwnd, self.rel_box)
            if frame is not None:
                # Non-blocking put — if the queue is full the stale frame is
                # simply dropped; the tracker always gets the freshest capture.
                try:
                    self._frame_queue.put(frame, block=False)
                except queue.Full:
                    pass
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
        self._update_label.configure(text=tr("Update available: v{}").format(latest))
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
            self._auto_btn.configure(text=tr("📷 AUTO  (observing...)"),
                                     fg_color="#b8a000", hover_color="#8a7800")
        else:
            self._auto_btn.configure(text=tr("📷 AUTO  (off)"),
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
            if self._auto_btn_state != "active":
                self._auto_btn_state = "active"
                self._auto_btn.configure(text=tr("📷 AUTO  \u2713"), fg_color="#3EC941", hover_color="#32a435")
            log.debug(f"AUTO heights: edging={self.heights['Edging']:.0f} "
                      f"erect={self.heights['Erect']:.0f} flaccid={self.heights['Flaccid']:.0f}")
        else:
            self._auto_btn.configure(text=tr("📷 AUTO  (flaccid set, need more range)"))

    def _disable_auto(self):
        """Turn off AUTO so manual height settings aren't overwritten."""
        if self._auto_mode:
            self._auto_mode = False
            self._auto_btn.configure(text=tr("📷 AUTO  (off)"),
                                     fg_color=self._C_SURFACE2, hover_color="#4a4a4a")

    def _cancel_pick(self):
        """Cancel any active Manual pick mode."""
        if self._pick_height:
            self._pick_height = None
            self.video_label.configure(cursor="")
            self.video_label.unbind("<Button-1>")

    def _add_exclusion_zone(self):
        """Let the user drag a rectangle on the video feed to mark an exclusion zone.

        Binds directly on video_label so the live feed stays visible underneath.
        The in-progress rectangle is drawn on each frame by _draw_tracking_overlay.
        """
        if getattr(self, '_ez_drawing', False):
            return  # already in progress

        self._ez_drawing    = True
        self._ez_disp_start = None
        self._ez_disp_end   = None

        vl = self.video_label
        vl.configure(cursor='crosshair')

        def _press(e):
            self._ez_disp_start = (e.x, e.y)
            self._ez_disp_end   = (e.x, e.y)

        def _drag(e):
            if self._ez_disp_start is not None:
                self._ez_disp_end = (e.x, e.y)

        def _release(e):
            self._ez_drawing    = False
            start               = self._ez_disp_start
            self._ez_disp_start = None
            self._ez_disp_end   = None
            vl.configure(cursor='')
            vl.unbind('<ButtonPress-1>')
            vl.unbind('<B1-Motion>')
            vl.unbind('<ButtonRelease-1>')

            if start is None:
                return
            x0, y0 = start
            x1, y1 = e.x, e.y
            if abs(x1 - x0) < 8 or abs(y1 - y0) < 8:
                return  # too small — ignore

            # Map display coords → source-frame coords
            sc = getattr(self, '_disp_scale', 1.0) or 1.0
            ox = getattr(self, '_disp_offset_x', 0)
            oy = getattr(self, '_disp_offset_y', 0)
            fw = getattr(self, '_disp_frame_w', 9999)
            fh = getattr(self, '_disp_frame_h', 9999)
            def _clamp(v, lo, hi): return max(lo, min(hi, v))
            fx0 = _clamp(int((min(x0, x1) - ox) / sc), 0, fw)
            fy0 = _clamp(int((min(y0, y1) - oy) / sc), 0, fh)
            fx1 = _clamp(int((max(x0, x1) - ox) / sc), 0, fw)
            fy1 = _clamp(int((max(y0, y1) - oy) / sc), 0, fh)
            if fx1 > fx0 and fy1 > fy0:
                self._exclusion_zones.append((fx0, fy0, fx1 - fx0, fy1 - fy0))
                self._save_config()
                log.info(f"Exclusion zone added: {self._exclusion_zones[-1]}, total={len(self._exclusion_zones)}")

        vl.bind('<ButtonPress-1>',   _press)
        vl.bind('<B1-Motion>',        _drag)
        vl.bind('<ButtonRelease-1>', _release)

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

    def _set_ponr(self):
        # NOTE: deliberately does NOT call _disable_auto() — the Point of No
        # Return is independent of the auto-calibrated edge/erect/flaccid lines,
        # so you can keep auto running and still drop a manual panic line on top.
        self._cancel_pick()
        self.heights["PONR"] = self.head_y
        self._ponr_triggered = False
        log.info(f"Point of No Return set at Y={self.head_y}")
        self._save_config()

    def _clear_ponr(self):
        self._cancel_pick()
        self.heights["PONR"] = None
        self._ponr_triggered = False
        log.info("Point of No Return cleared")
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
        self.heights = {"Edging": None, "Erect": None, "Flaccid": None, "PONR": None}
        self._ponr_triggered = False
        self._auto_min_y = None
        self._auto_max_y = None
        self._auto_obs_start = None
        self._auto_last_apply = 0.0
        self._head_y_history.clear()
        self._auto_btn_state = None
        if hasattr(self, '_auto_btn'):
            self._auto_btn.configure(text=tr("📷 AUTO"), fg_color=self._C_SURFACE2,
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
                self._tracker_init(pause_frame, new_bbox)
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

    # Evil Mode ruin constants
    _RUIN_ODDS_DEFAULT    = {"Easy": 0, "Middle": 5, "Hard": 20, "Expert": 40}
    _RUIN_COOLDOWN        = {"Easy": 60, "Middle": 90, "Hard": 120, "Expert": 180}
    _RUIN_PHRASES_DEFAULT = [
        "You thought that was it? Cute.",
        "Ruined. Just like you deserve.",
        "Close enough. No.",
        "That's not cumming, that's suffering.",
        "Congratulations on nothing.",
        "Ruin accepted. Permission denied.",
        "That didn't count.",
        "Oh you were close. Too bad.",
        "Felt good for a second, didn't it.",
        "Gone. All of it. Gone.",
    ]

    # Odds of "Let me cum?" being granted per aggressiveness level
    _CUM_ODDS_DEFAULT = {"Easy": 2, "Middle": 4, "Hard": 6, "Expert": 30}

    # Hands-free "Let me cum?" — auto-grant after enough edges (settings-only toggle)
    # _hf_min_edges : minimum edge count before HF can ever fire (per level)
    # _hf_cum_chance: 1-in-N roll per edge once min is reached (lower N = more likely)
    _HF_MIN_EDGES_DEFAULT  = {"Easy": 2, "Middle": 4, "Hard":  6, "Expert": 10}
    _HF_CUM_CHANCE_DEFAULT = {"Easy": 2, "Middle": 4, "Hard":  6, "Expert": 10}
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
        # Guard against re-entry from states where a roll is unsafe/nonsensical.
        # The button is disabled in these states, but the VOICE path ("let me cum"
        # / "please") bypasses the widget, and pressing the green "CUM NOW!" button
        # during an active grant would otherwise re-roll and could REVOKE the grant
        # (or trigger a ruin). Refuse to roll while stopped/granted/counting down.
        if self._cum_stopped or self._cum_allowed or self._cum_cd_active:
            return
        # Check denial cooldown
        cooldown_left = self._letmecum_cooldown_until - time.time()
        if cooldown_left > 0:
            self._letmecum_btn.configure(text=tr("Wait {}s...").format(int(cooldown_left)))
            # Voice users can't see the button — give them the snark too
            self._snark_label.configure(text=tr("Not yet. Wait {}s.").format(int(cooldown_left)),
                                        text_color="#F5A623")
            return

        aggr = self.aggr_var.get()
        denominator = self._cum_odds.get(aggr, 4)
        granted = random.randint(1, denominator) == 1
        self._last_letmecum_result = "granted" if granted else "denied"
        self._last_letmecum_time = time.time()
        self.session_logger.log_letmecum(self._last_letmecum_result, aggr, denominator)

        if granted:
            self._grant_cum(aggr, source=f"manual 1/{denominator}")
        else:
            self._cum_allowed = False
            self._denial_count += 1

            # Evil Mode: roll for ruin before normal denial cooldown
            if self._evil_mode:
                ruin_pct = self._ruin_odds.get(aggr, 0)
                if ruin_pct > 0 and random.randint(1, 100) <= ruin_pct:
                    self.session_logger.log_letmecum("ruined", aggr, denominator)
                    self._do_ruin(aggr)
                    return

            cooldown = self._CUM_DENY_COOLDOWN.get(aggr, 30)
            self._letmecum_cooldown_until = time.time() + cooldown
            self._letmecum_btn.configure(text=tr("DENIED!"), fg_color="#FF4444",
                                         hover_color="#cc3636")
            snark = random.choice(self._denial_phrases) if self._denial_phrases else "Denied."
            self._snark_label.configure(text=snark)
            self.root.after(2000, lambda: self._tick_letmecum_cooldown())
            log.info(f"Cum DENIED (1/{denominator} on {aggr}) — {cooldown}s cooldown")

    def _grant_cum(self, aggr: str, source: str = "manual"):
        """Grant cum permission for the given aggressiveness level.

        Called by _on_letmecum() on a successful roll, and directly by
        _hf_check_edge() when hands-free mode auto-fires.
        """
        self._cum_allowed = True
        grant_secs = self._CUM_GRANT_TIME.get(aggr, 300)
        self._cum_grant_expires = time.time() + grant_secs
        mins = grant_secs // 60
        hf_tag = " [auto]" if source != "manual" and not source.startswith("manual") else ""
        self._snark_label.configure(
            text=tr("You've been a good boy.{} You have {} min.").format(hf_tag, mins),
            text_color="#3EC941")
        self._letmecum_btn.configure(text=tr("CUM NOW! {}:00").format(mins), fg_color="#3EC941",
                                     hover_color="#32a435")
        self._tribute_btn.pack(padx=3, pady=(0, 3))   # #2: nudge at peak, never blocks
        if self._cum_override_range:
            target_vol = 1.0
        else:
            target_vol = self.max_vol_var.get() / 100.0
        self._set_all_outputs(target_vol, instant=True)
        self.root.after(1000, self._tick_cum_grant)
        log.info(f"Cum GRANTED ({source}, {aggr}) — {mins} min window")

    def _tick_letmecum_cooldown(self):
        """Update the button text with remaining cooldown, then restore."""
        if not self._running:
            return
        remaining = self._letmecum_cooldown_until - time.time()
        if remaining > 0:
            self._letmecum_btn.configure(text=tr("Wait {}s...").format(int(remaining)),
                                         fg_color="#FF4444", hover_color="#cc3636")
            self.root.after(1000, self._tick_letmecum_cooldown)
        else:
            if self._evil_mode:
                self._letmecum_btn.configure(text=tr("Let me cum?"),
                                             fg_color="#8a1a1a", hover_color="#6a0a0a")
            else:
                self._letmecum_btn.configure(text=tr("Let me cum?"),
                                             fg_color="#3EC941", hover_color="#32a435")
            self._snark_label.configure(text="")

    def _tick_cum_grant(self):
        """Countdown the cum grant window. When expired, revoke permission."""
        if not self._running:
            return
        if not self._cum_allowed:
            return
        remaining = self._cum_grant_expires - time.time()
        if remaining > 0:
            m = int(remaining) // 60
            s = int(remaining) % 60
            self._letmecum_btn.configure(text=tr("CUM NOW! {}:{}").format(m, f"{s:02d}"))
            self._snark_label.configure(
                text=tr("You've been a good boy. {}:{} remaining.").format(m, f"{s:02d}"),
                text_color="#3EC941")
            self.root.after(1000, self._tick_cum_grant)
        else:
            # Time's up — revoke permission
            self._cum_allowed = False
            self._last_letmecum_result = "expired"
            self._tribute_btn.pack_forget()
            self._letmecum_btn.configure(text=tr("Too slow!"),
                                         fg_color="#FF4444", hover_color="#cc3636")
            self._snark_label.configure(text=tr("Time's up. Back to edging."),
                                        text_color="#ff4444")
            if getattr(self, '_evil_mode', False):
                _btn_fg  = "#8a1a1a"
                _btn_hov = "#6b0000"
            else:
                _btn_fg  = "#3EC941"
                _btn_hov = "#32a435"
            self.root.after(3000, lambda fg=_btn_fg, hov=_btn_hov: (
                self._letmecum_btn.configure(text=tr("Let me cum?"),
                                             fg_color=fg, hover_color=hov),
                self._snark_label.configure(text="")))
            log.info("Cum grant expired — permission revoked")

    def _set_all_outputs(self, vol: float, instant: bool = False):
        """Send vol to every active output. Caller is responsible for clamping vol first."""
        if self.restim_on.get():
            self.restim.set_volume(vol, instant=instant)
        if self.xtoys_on.get():
            self.xtoys.set_volume(vol, instant=instant)
        if self.audio_on.get() and self.win_audio and self.win_audio.connected:
            self.win_audio.set_volume(vol, 0.0, 1.0)
        if self.mp3_on.get() and self.music_player:
            self.music_player.volume = vol

    def _ruin_set_volume(self, vol: float):
        """Set all active outputs instantly (used by the ruin pulse sequence)."""
        # If a cum/safeword landed mid-ruin, the session is stopped — never drive
        # output from a stale pulse (belt-and-suspenders alongside _cancel_ruin).
        if self._cum_stopped:
            return
        self._set_all_outputs(vol, instant=True)

    def _cancel_ruin(self):
        """Abort an in-progress ruin pulse sequence — its after() jobs would
        otherwise keep firing (up to 6s out, driving to 100%) into a stopped/
        post-orgasm session. Called by _cum_confirm and _voice_safeword."""
        self._ruin_active = False
        for j in getattr(self, "_ruin_jobs", []):
            try:
                self.root.after_cancel(j)
            except Exception:
                pass
        self._ruin_jobs = []

    def _do_ruin(self, aggr: str):
        """Execute the ruin pulse sequence using root.after() — no blocking."""
        self._last_letmecum_result = "ruin"
        self._last_letmecum_time = time.time()
        self._ruin_count += 1
        self._ruin_active = True
        self._ruin_jobs = []
        log.info(f"RUIN triggered on {aggr}")

        self._letmecum_btn.configure(text=tr("RUINED 😈"), fg_color="#6600cc",
                                     hover_color="#440088")

        # Longer, crueler 5-pulse climb (85 → 90 → 94 → 97 → 100) with the
        # final 100% held ~1.3s before the hard cut. ~5.7s total.
        # t=0: 85%
        self._ruin_set_volume(0.85)

        def _final():
            # HARD CUT to 0, lock, show ruin phrase, start cooldown
            self._ruin_active = False
            self._ruin_jobs = []
            self._ruin_set_volume(0.0)
            cooldown = self._RUIN_COOLDOWN.get(aggr, 90)
            self._letmecum_cooldown_until = time.time() + cooldown
            phrase = random.choice(self._ruin_phrases) if self._ruin_phrases else "Ruined."
            self._snark_label.configure(text=phrase)
            self._last_letmecum_result = "ruined"
            self._last_letmecum_time = time.time()
            # Transition button into cooldown state
            self.root.after(1000, lambda: self._tick_letmecum_cooldown())

        # (delay_ms, volume) — None volume marks the terminal hard cut
        ruin_steps = [
            ( 700, 0.0),
            (1200, 0.90),
            (1900, 0.0),
            (2400, 0.94),
            (3100, 0.0),
            (3600, 0.97),
            (4300, 0.0),
            (4800, 1.00),
            (6100, None),   # held 100% for ~1.3s, then hard cut
        ]
        for delay, vol in ruin_steps:
            if vol is None:
                self._ruin_jobs.append(self.root.after(delay, _final))
            else:
                self._ruin_jobs.append(
                    self.root.after(delay, lambda v=vol: self._ruin_set_volume(v)))

    def _on_cum(self, source="click"):
        """Hard stop — volume to 0; refractory countdown then auto-resume.

        source="click"  → 3-second undo window before committing.
        source="voice"  → immediate (no undo window).
        Internal calls (auto-cum, safeword) pass source="voice" to skip undo.
        """
        # UX-1: if the undo window is active, a repeat *click* cancels the cum.
        # But voice/internal sources (safeword, auto-cum, "came") must COMMIT, not
        # cancel — otherwise saying the safeword during the 3s window silently
        # aborts the stop and leaves the toy running. Cancel the pending timer and
        # confirm immediately for those.
        if self._cum_undo_active:
            if source == "click":
                self._cum_undo()
                return
            if self._cum_undo_job:
                self.root.after_cancel(self._cum_undo_job)
                self._cum_undo_job = None
            self._cum_confirm()
            return

        if source == "click" and not self._cum_stopped:
            # Arm undo window
            self._cum_undo_active = True
            try:
                self._cum_btn.configure(
                    text=tr("↩ Undo (3s)"), command=self._on_cum,
                    fg_color="#F5A623", hover_color="#d08800", text_color="black")
            except Exception:
                pass
            btn = getattr(self, '_play_cum_btn', None)
            if btn:
                try:
                    btn.configure(
                        text=tr("↩ Undo (3s)"), command=self._on_cum,
                        fg_color="#F5A623", hover_color="#d08800", text_color="black")
                except Exception:
                    pass
            self._cum_undo_job = self.root.after(3000, self._cum_confirm)
            return

        # Voice / internal path — commit immediately
        self._cum_confirm()

    def _cum_undo(self):
        """Cancel the 3-second undo window — restores button, no cum logged."""
        if self._cum_undo_job:
            self.root.after_cancel(self._cum_undo_job)
        self._cum_undo_job    = None
        self._cum_undo_active = False
        _cum_kw = dict(text=tr("I've CUM"), command=self._on_cum,
                       fg_color="#e0e0e8", hover_color="#c8c8d0", text_color=self._C_BG)
        try: self._cum_btn.configure(**_cum_kw)
        except Exception: pass
        try: self._play_cum_btn.configure(**_cum_kw)
        except Exception: pass
        log.info("cum undone")

    def _cum_confirm(self):
        """Actual cum logic — called after undo window expires or immediately for voice."""
        if not self._cum_undo_active and self._cum_stopped:
            # Already stopped (e.g. called twice); nothing to do
            log.info("_cum_confirm: no-op — already stopped (stray cum-button click in refractory?)")
            return
        self._cum_undo_active = False
        self._cum_undo_job    = None

        # Disarm any in-progress auto-cum countdown (fired manually or by _cum_now)
        if self._cum_cd_active:
            if self._cum_cd_job:
                self.root.after_cancel(self._cum_cd_job)
            self._cum_cd_job    = None
            self._cum_cd_active = False
        # Kill any in-progress ruin pulses so they can't drive to 100% after this stop
        self._cancel_ruin()
        self._cum_score = 0.0
        self._cum_count += 1
        self.session_logger.log_cum()
        self._cum_time = time.time()
        self._cum_stopped = True
        self._cum_allowed = False
        self._tribute_btn.pack_forget()
        self._cum_grant_expires = 0
        self._letmecum_btn.configure(text=tr("Let me cum?"), fg_color="#3EC941",
                                     hover_color="#32a435", state="disabled")
        self._snark_label.configure(text="")

        # Zero all outputs
        self._set_all_outputs(0.0, instant=True)

        if self._refractory_mins > 0:
            # Refractory countdown — clicking the button skips it early
            self._refractory_until = time.time() + self._refractory_mins * 60
            self._cum_btn.configure(
                text=tr("Refractory: {}:00\n(click to skip)").format(self._refractory_mins),
                command=self._on_resume, state="normal",
                fg_color=self._C_SURFACE2, hover_color="#4a4a4a",
                text_color=self._C_TEXT_DIM)
            btn = getattr(self, '_play_cum_btn', None)
            if btn:
                btn.configure(
                    text=tr("Refractory: {}:00").format(self._refractory_mins),
                    command=self._on_resume, state="normal",
                    fg_color=self._C_SURFACE2, hover_color="#4a4a4a",
                    text_color=self._C_TEXT_DIM)
            self.root.after(1000, self._tick_refractory)
            log.info(f"Refractory started — {self._refractory_mins} min")
        else:
            # No refractory — immediate manual resume
            self._cum_btn.configure(text=tr("Resume"), command=self._on_resume,
                                    fg_color=self._C_GREEN, hover_color=self._C_GREEN_H,
                                    text_color="white")
            btn = getattr(self, '_play_cum_btn', None)
            if btn:
                btn.configure(text=tr("Resume"), command=self._on_resume,
                               fg_color=self._C_GREEN, hover_color=self._C_GREEN_H,
                               text_color="white")
            log.info("Cum triggered — session stopped, awaiting Resume")

    def _tick_refractory(self):
        """Tick the refractory countdown once per second; auto-resume at 0."""
        if not self._running or not self._cum_stopped:
            return
        remaining = self._refractory_until - time.time()
        if remaining <= 0:
            self._on_resume()
            return
        m = int(remaining) // 60
        s = int(remaining) % 60
        label = tr("Refractory: {}:{}\n(click to skip)").format(m, f"{s:02d}")
        short  = f"Refractory: {m}:{s:02d}"
        self._cum_btn.configure(text=label)
        btn = getattr(self, '_play_cum_btn', None)
        if btn:
            btn.configure(text=short)
        self.root.after(1000, self._tick_refractory)

    def _on_resume(self):
        """Restart the edging loop — called by refractory auto-expire or manual skip."""
        log.info("_on_resume: skip/resume invoked (cum_stopped=%s)", self._cum_stopped)
        self._cum_stopped = False
        self._cum_time = None
        self._refractory_until = 0.0
        # Reset detection state so next round has a clean slate
        self._cum_detect_buf.clear()
        self._cum_peak_activity = 0.0
        self._cum_score = 0.0
        _cum_kw = dict(text=tr("I've CUM"), command=self._on_cum,
                       fg_color="#e0e0e8", hover_color="#c8c8d0", text_color=self._C_BG)
        self._cum_btn.configure(**_cum_kw)
        try:
            self._play_cum_btn.configure(**_cum_kw)
        except Exception:
            pass
        try:
            self._letmecum_btn.configure(state="normal")
        except Exception:
            pass
        log.info("Session resumed after refractory")

    # ------------------------------------------------------------------ auto-cum detection

    def _on_auto_cum_toggle(self):
        self._auto_cum_enabled = bool(self._auto_cum_var.get())
        if not self._auto_cum_enabled:
            # Cancel any active countdown if user turns it off
            self._cum_cancel()
        self._cum_detect_buf.clear()
        self._cum_peak_activity = 0.0
        self._cum_score = 0.0
        self._save_config()

    def _hf_check_edge(self):
        """Called (on main thread) each time edge_count increments.

        If hands-free mode is enabled and the level-specific minimum edge count
        has been reached, rolls 1-in-N and — on a hit — grants cum permission
        automatically via _grant_cum(), just as if the user had pressed and won
        "Let me cum?".
        """
        if not self._hf_enabled:
            return
        if self._cum_stopped or self._cum_allowed or self._cum_cd_active:
            return
        # Don't fire during a denial cooldown (respect the cooldown window)
        if self._letmecum_cooldown_until > time.time():
            return
        aggr    = self.aggr_var.get()
        min_e   = self._hf_min_edges.get(aggr, 999)
        if self.edge_count < min_e:
            return
        chance  = max(1, self._hf_cum_chance.get(aggr, 10))
        if random.randint(1, chance) == 1:
            log.info(f"Hands-free cum: edge {self.edge_count}, rolled 1/{chance} on {aggr} — GRANTED")
            self.session_logger.log_letmecum("granted", aggr, chance)
            self._last_letmecum_result = "granted"
            self._last_letmecum_time   = time.time()
            self._grant_cum(aggr, source=f"hands-free 1/{chance}")
        else:
            log.debug(f"Hands-free cum: edge {self.edge_count}, rolled miss on 1/{chance} ({aggr})")

    def _tick_cum_detect(self, y: int, frame=None):
        """
        Called each heavy frame when tracking OK and auto-cum is enabled.

        Two independent signals feed a combined score (0–25):
          Motion:     slow-window std drops (stroking stopped) + fast-window jitter present
          Brightness: ROI mean brightness spikes ≥ 2.5σ above session baseline
                      (ejaculation creates a bright white visual event)

        Score accumulation:
          Motion + brightness:  +1.5 / frame   (strongest signal)
          Motion only:          +1.0 / frame   (original behaviour)
          Brightness only:      +0.5 / frame   (visual signal without motion signature)
          Neither:              −2.0 / frame
        """
        if self._cum_cd_active or self._cum_stopped:
            return

        self._cum_detect_buf.append(y)
        n = len(self._cum_detect_buf)
        if n < 30:
            return

        buf = list(self._cum_detect_buf)

        # ── Motion signal ─────────────────────────────────────────────────────
        fast_std = float(np.std(buf[-8:]))  if n >= 8  else float(np.std(buf))
        slow_std = float(np.std(buf[-60:])) if n >= 60 else float(np.std(buf))

        self._cum_peak_activity = max(self._cum_peak_activity * 0.9998, slow_std)

        sensitivity = self._auto_cum_sensitivity
        min_peak_px = max(2.0, 15.0 - sensitivity * 1.0)
        if self._cum_peak_activity < min_peak_px:
            return

        stroking_paused = slow_std < self._cum_peak_activity * 0.30
        jitter_present  = 1.5 < fast_std < 20.0
        motion_signal   = stroking_paused and jitter_present

        # ── Brightness spike signal ───────────────────────────────────────────
        bright_signal = False
        if frame is not None:
            bx, by, bw, bh = self.last_bbox
            fh, fw = frame.shape[:2]
            roi = frame[max(0, by):min(fh, by + bh), max(0, bx):min(fw, bx + bw)]
            if roi.size > 0:
                gray = cv2.cvtColor(roi, cv2.COLOR_BGR2GRAY)
                lit  = gray[gray > 30]          # ignore dark background pixels
                if lit.size > 0:
                    brightness = float(np.mean(lit))
                    self._bright_buf.append(brightness)
                    nb = len(self._bright_buf)
                    if nb >= 300 and not self._bright_baseline_ready:
                        # 10-second initial calibration window
                        self._bright_baseline_mean  = float(np.mean(self._bright_buf))
                        self._bright_baseline_std   = max(2.0, float(np.std(self._bright_buf)))
                        self._bright_baseline_ready = True
                    elif self._bright_baseline_ready:
                        z = (brightness - self._bright_baseline_mean) / self._bright_baseline_std
                        bright_signal = z >= 2.5
                        # Slowly drift baseline toward current level (resists long-term lighting shifts)
                        diff = brightness - self._bright_baseline_mean
                        self._bright_baseline_mean += diff * 0.001
                        self._bright_baseline_std   = max(2.0,
                            self._bright_baseline_std * 0.9999 + abs(diff) * 0.0001)

        # ── Score accumulation / decay ────────────────────────────────────────
        if motion_signal and bright_signal:
            self._cum_score = min(self._cum_score + 1.5, 25.0)
        elif motion_signal:
            self._cum_score = min(self._cum_score + 1.0, 25.0)
        elif bright_signal:
            self._cum_score = min(self._cum_score + 0.5, 25.0)
        else:
            self._cum_score = max(self._cum_score - 2.0, 0.0)

        # ── Trigger ───────────────────────────────────────────────────────────
        trigger = 25.0 - (sensitivity - 1) * 2.0   # sens=1 → 25, sens=10 → 7
        if self._cum_score >= trigger:
            self._cum_score = 0.0
            self.root.after(0, self._start_cum_countdown)

    def _start_cum_countdown(self):
        """Begin the cancel-window countdown that precedes auto-firing _on_cum."""
        if self._cum_cd_active or self._cum_stopped:
            return
        self._cum_cd_active    = True
        self._cum_cd_remaining = self._auto_cum_delay
        log.info(f"Auto-cum detected — {self._auto_cum_delay}s countdown started")

        # Repurpose "Let me cum?" as the Cancel button
        try:
            self._letmecum_btn.configure(
                text=tr("✕ Cancel"), command=self._cum_cancel,
                fg_color=self._C_SURFACE2, hover_color="#4a4a4a",
                text_color=self._C_TEXT, state="normal",
            )
        except Exception:
            pass
        self._cum_cd_tick()

    def _cum_cd_tick(self):
        """Decrement countdown; fire or schedule next tick."""
        if not self._cum_cd_active:
            return
        rem = self._cum_cd_remaining
        _orange = "#d07020"
        _orange_h = "#b05010"
        label = f"Cumming… {rem}" if rem > 0 else "Cumming…"
        for btn in (self._cum_btn, getattr(self, '_play_cum_btn', None)):
            if btn is None:
                continue
            try:
                btn.configure(
                    text=label,
                    fg_color=_orange, hover_color=_orange_h,
                    text_color="white",
                    command=self._cum_now,   # clicking = skip delay
                )
            except Exception:
                pass

        # UX-5: show countdown in snark label
        try:
            if rem > 0:
                self._snark_label.configure(
                    text=tr("⏳ Cum window: {}s remaining — say 'came' or click Cancel").format(rem),
                    text_color="#3EC941")
            else:
                self._snark_label.configure(text="", text_color=self._C_TEXT)
        except Exception:
            pass

        if rem <= 0:
            self._cum_now()
            return

        self._cum_cd_remaining -= 1
        self._cum_cd_job = self.root.after(1000, self._cum_cd_tick)

    def _cum_now(self):
        """Skip remaining delay and fire _on_cum immediately."""
        if self._cum_cd_job:
            self.root.after_cancel(self._cum_cd_job)
        self._cum_cd_job    = None
        self._cum_cd_active = False
        self._restore_cum_btn_normal()
        self._on_cum(source="voice")

    def _cum_cancel(self):
        """Cancel the auto-cum countdown and restore buttons."""
        if self._cum_cd_job:
            self.root.after_cancel(self._cum_cd_job)
        self._cum_cd_job    = None
        self._cum_cd_active = False
        self._cum_score     = 0.0
        # Damp peak so the detector needs renewed activity before re-triggering
        self._cum_peak_activity *= 0.5
        self._restore_cum_btn_normal()
        # UX-5: clear the countdown snark message
        try:
            self._snark_label.configure(text="", text_color=self._C_TEXT)
        except Exception:
            pass
        log.info("Auto-cum countdown cancelled by user")

    def _restore_cum_btn_normal(self):
        """Put I've CUM / Let me cum? back to their standard appearance."""
        _cum_kw = dict(
            text=tr("I've CUM"), command=self._on_cum,
            fg_color="#e0e0e8", hover_color="#c8c8d0", text_color=self._C_BG,
        )
        try: self._cum_btn.configure(**_cum_kw)
        except Exception: pass
        try: self._play_cum_btn.configure(**_cum_kw)
        except Exception: pass
        try:
            self._letmecum_btn.configure(
                text=tr("Let me cum?"), command=self._on_letmecum,
                fg_color=self._C_GREEN, hover_color=self._C_GREEN_H,
                text_color="white",
            )
        except Exception:
            pass

    def _hr_log_poll(self):
        """Log heart rate reading every 30 seconds when HR is active."""
        if self.hr_on.get() and self._active_hr.connected:
            bpm = self._active_hr.smooth_bpm()
            if bpm is not None:
                self.session_logger.log_heart_rate(round(bpm), self._active_hr.modifier())
        self.root.after(30_000, self._hr_log_poll)

    def _open_settings(self):
        """Open the settings dialog."""
        # UX-3: snapshot reversible settings before opening so Cancel can restore them
        _orig = {
            "cum_odds":          dict(self._cum_odds),
            "ruin_odds":         dict(self._ruin_odds),
            "denial_phrases":    list(self._denial_phrases),
            "ruin_phrases":      list(self._ruin_phrases),
            "cum_override_range":self._cum_override_range,
            "refractory_mins":   self._refractory_mins,
            "erect_recovery":    self._erect_recovery,
            "recovery_pct":      dict(self._recovery_pct),
            "ae_hold_secs":      dict(self._ae_hold_secs),
            "ae_ramp_pct":       dict(self._ae_ramp_pct),
            "auto_cum_delay":    self._auto_cum_delay,
            "auto_cum_sensitivity": self._auto_cum_sensitivity,
            "bondage_safeword":  self._bondage_safeword,
            "hf_enabled":        self._hf_enabled,
            "hf_min_edges":      dict(self._hf_min_edges),
            "hf_cum_chance":     dict(self._hf_cum_chance),
        }

        win = ctk.CTkToplevel(self.root)
        win.title(tr("Settings"))
        win.configure(fg_color=self._C_BG)
        win.transient(self.root)
        win.grab_set()
        # Cap height to screen so content never overflows
        _scr_h = win.winfo_screenheight()
        win.geometry(f"490x{min(860, _scr_h - 80)}")

        # Scrollable content body + pinned OK/Reset footer
        sf = ctk.CTkScrollableFrame(win, fg_color="transparent", corner_radius=0)
        sf.pack(fill=tk.BOTH, expand=True)

        lbl = ctk.CTkFont(size=11, weight="bold")

        # ── Appearance (Theme + Font) ─────────────────────────────────────────
        ctk.CTkLabel(sf, text=tr("Appearance"),
                     font=lbl, text_color=self._C_TEXT).pack(padx=16, pady=(12, 4), anchor="w")
        appear_frame = ctk.CTkFrame(sf, fg_color=self._C_SURFACE, corner_radius=8)
        appear_frame.pack(fill=tk.X, padx=16, pady=(0, 8))

        # Theme selector — collapsed by default, expand on click
        theme_var = tk.StringVar(value=self._theme_name)
        theme_hdr = ctk.CTkFrame(appear_frame, fg_color="transparent", cursor="hand2")
        theme_hdr.pack(fill=tk.X, padx=12, pady=(8, 0))
        theme_arrow_lbl = ctk.CTkLabel(theme_hdr, text="▶", font=ctk.CTkFont(size=10),
                                       text_color=self._C_TEXT_DIM, width=14)
        theme_arrow_lbl.pack(side=tk.LEFT)
        theme_title_lbl = ctk.CTkLabel(
            theme_hdr,
            text=tr("Theme: {}").format(self._theme_name),
            font=ctk.CTkFont(size=10, weight="bold"),
            text_color=self._C_TEXT_DIM, anchor="w")
        theme_title_lbl.pack(side=tk.LEFT, padx=(4, 0))

        theme_rows_frame = ctk.CTkFrame(appear_frame, fg_color="transparent")
        # not packed yet — hidden until toggled

        def _build_theme_rows():
            for tname, tcolors in THEMES.items():
                row = ctk.CTkFrame(theme_rows_frame, fg_color="transparent")
                row.pack(fill=tk.X, padx=0, pady=2)
                ctk.CTkRadioButton(row, text=tname, variable=theme_var, value=tname,
                                   font=ctk.CTkFont(size=11), text_color=self._C_TEXT,
                                   fg_color=self._C_ACCENT, hover_color=self._C_ACCENT_H,
                                   border_color=self._C_BORDER, width=100,
                                   command=lambda t=tname: theme_title_lbl.configure(
                                       text=tr("Theme: {}").format(t))
                                   ).pack(side=tk.LEFT)
                swatch_cv = tk.Canvas(row, width=120, height=16, bg=self._C_SURFACE,
                                      highlightthickness=0)
                swatch_cv.pack(side=tk.LEFT, padx=(8, 0))
                preview_colors = [tcolors["BG"], tcolors["SURFACE"], tcolors["ACCENT"],
                                  tcolors["RED"], tcolors["GREEN"], tcolors["BLUE"]]
                for i, col in enumerate(preview_colors):
                    swatch_cv.create_rectangle(i * 20, 1, i * 20 + 18, 15,
                                               fill=col, outline=tcolors["BORDER"], width=1)

        _theme_rows_built = [False]
        _theme_open = [False]

        def _toggle_theme_rows(_event=None):
            if not _theme_rows_built[0]:
                _build_theme_rows()
                _theme_rows_built[0] = True
            if _theme_open[0]:
                theme_rows_frame.pack_forget()
                theme_arrow_lbl.configure(text="▶")
            else:
                theme_rows_frame.pack(fill=tk.X, padx=12, pady=(0, 4))
                theme_arrow_lbl.configure(text="▼")
            _theme_open[0] = not _theme_open[0]

        theme_hdr.bind("<Button-1>", _toggle_theme_rows)
        theme_arrow_lbl.bind("<Button-1>", _toggle_theme_rows)
        theme_title_lbl.bind("<Button-1>", _toggle_theme_rows)

        # Font size
        ctk.CTkFrame(appear_frame, height=1, fg_color=self._C_BORDER
                     ).pack(fill=tk.X, padx=12, pady=6)
        ctk.CTkLabel(appear_frame, text=tr("Font Size"), font=ctk.CTkFont(size=10, weight="bold"),
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
        font_val_lbl = ctk.CTkLabel(font_row, text=tr("{}px").format(self._ui_font_size), font=lbl,
                                     text_color=self._C_ACCENT, width=36)
        font_val_lbl.pack(side=tk.LEFT)

        # Live font preview
        font_preview = ctk.CTkLabel(appear_frame, text=tr("Sample Text Abc 123"),
                                     font=ctk.CTkFont(size=self._ui_font_size, weight="bold"),
                                     text_color=self._C_TEXT)
        font_preview.pack(padx=12, pady=(0, 4), anchor="w")

        def _on_font_slide(v):
            sz = int(float(v))
            font_val_lbl.configure(text=tr("{}px").format(sz))
            font_preview.configure(font=ctk.CTkFont(size=sz, weight="bold"))
        font_slider.configure(command=_on_font_slide)

        # Language
        ctk.CTkFrame(appear_frame, height=1, fg_color=self._C_BORDER
                     ).pack(fill=tk.X, padx=12, pady=6)
        ctk.CTkLabel(appear_frame, text=tr("Language"),
                     font=ctk.CTkFont(size=10, weight="bold"),
                     text_color=self._C_TEXT_DIM).pack(padx=12, pady=(0, 2), anchor="w")
        _LANG_LABELS = {"en": "English", "zh": "中文"}
        _lang_code_by_label = {v: k for k, v in _LANG_LABELS.items()}
        lang_var = tk.StringVar(value=_LANG_LABELS.get(self._language, "English"))
        ctk.CTkSegmentedButton(
            appear_frame, values=list(_LANG_LABELS.values()), variable=lang_var,
            font=ctk.CTkFont(size=11),
        ).pack(fill=tk.X, padx=12, pady=(0, 4))

        # Apply button
        def _apply_appearance():
            self._theme_name = theme_var.get()
            self._ui_font_size = int(font_var.get())
            self._language = _lang_code_by_label.get(lang_var.get(), "en")
            self._save_config()
            win.destroy()
            # Use subprocess.Popen + sys.exit so Windows doesn't get confused.
            # os.execv after root.destroy() is unreliable on Windows because
            # tkinter threads may still be running when exec replaces the image.
            import sys, subprocess
            self._cleanup()
            if getattr(sys, "frozen", False):
                subprocess.Popen([sys.executable] + sys.argv[1:])
            else:
                subprocess.Popen([sys.executable] + sys.argv)
            sys.exit(0)

        ctk.CTkButton(appear_frame, text=tr("Apply (restarts app)"), command=_apply_appearance,
                      font=ctk.CTkFont(size=10, weight="bold"), height=28, corner_radius=4,
                      fg_color=self._C_ACCENT, hover_color=self._C_ACCENT_H,
                      text_color="white").pack(padx=12, pady=(4, 10), anchor="e")

        # ── Calibration ───────────────────────────────────────────────────────
        ctk.CTkLabel(sf, text=tr("Calibration"),
                     font=lbl, text_color=self._C_TEXT).pack(padx=16, pady=(8, 4), anchor="w")
        calib_frame = ctk.CTkFrame(sf, fg_color=self._C_SURFACE, corner_radius=8)
        calib_frame.pack(fill=tk.X, padx=16, pady=(0, 8))

        sens_hdr = ctk.CTkFrame(calib_frame, fg_color="transparent")
        sens_hdr.pack(fill=tk.X, padx=12, pady=(8, 2))
        ctk.CTkLabel(sens_hdr, text=tr("Edging Sensitivity"),
                     font=ctk.CTkFont(size=10, weight="bold"),
                     text_color=self._C_TEXT_DIM, anchor="w").pack(side=tk.LEFT)
        sens_val_lbl = ctk.CTkLabel(sens_hdr, text=tr("{} px").format(int(self.edge_sens_var.get())),
                                    font=lbl, text_color=self._C_ACCENT, anchor="e")
        sens_val_lbl.pack(side=tk.RIGHT)

        def _on_sens(val):
            sens_val_lbl.configure(text=tr("{} px").format(int(float(val))))
            self._save_config()

        ctk.CTkSlider(
            calib_frame, from_=0, to=500, number_of_steps=500,
            variable=self.edge_sens_var, command=_on_sens,
            button_color=self._C_ACCENT, button_hover_color=self._C_ACCENT_H,
            progress_color=self._C_ACCENT, fg_color=self._C_SURFACE2,
        ).pack(fill=tk.X, padx=12, pady=(0, 4))
        ctk.CTkLabel(
            calib_frame,
            text=tr("Pulls the Edging state-display toward Erect by N pixels so "
                 "the OBS overlay lights up earlier. "
                 "Affects only the State label / OBS broadcast — the volume / "
                 "denial math always uses the strict calibrated line."),
            font=ctk.CTkFont(size=9), text_color=self._C_TEXT_DIM,
            wraplength=420, justify="left", anchor="w",
        ).pack(fill=tk.X, padx=12, pady=(0, 10))

        # ── Restim ────────────────────────────────────────────────────────────
        ctk.CTkLabel(sf, text=tr("Restim"),
                     font=lbl, text_color=self._C_TEXT).pack(padx=16, pady=(8, 4), anchor="w")
        restim_frame = ctk.CTkFrame(sf, fg_color=self._C_SURFACE, corner_radius=8)
        restim_frame.pack(fill=tk.X, padx=16, pady=(0, 8))

        ax_row = ctk.CTkFrame(restim_frame, fg_color="transparent")
        ax_row.pack(fill=tk.X, padx=12, pady=(8, 2))
        ctk.CTkLabel(ax_row, text=tr("T-code Axis"),
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
            text=tr("Which T-code axis VSE sends volume commands on. This must "
                 "match the axis your Restim session has bound to the "
                 "parameter you want to control (usually 'Volume'). Check "
                 "Restim's Websocket / T-code panel if you're unsure."),
            font=ctk.CTkFont(size=9), text_color=self._C_TEXT_DIM,
            wraplength=420, justify="left", anchor="w",
        ).pack(fill=tk.X, padx=12, pady=(4, 10))

        # ── xToys ─────────────────────────────────────────────────────────────
        # ── Performance ───────────────────────────────────────────────────────
        ctk.CTkLabel(sf, text=tr("Performance"),
                     font=lbl, text_color=self._C_TEXT).pack(padx=16, pady=(12, 4), anchor="w")
        perf_frame = ctk.CTkFrame(sf, fg_color=self._C_SURFACE, corner_radius=8)
        perf_frame.pack(fill=tk.X, padx=16, pady=(0, 8))
        eng_row = ctk.CTkFrame(perf_frame, fg_color="transparent")
        eng_row.pack(fill=tk.X, padx=12, pady=(10, 4))
        ctk.CTkLabel(eng_row, text=tr("Tracking engine"), font=lbl,
                     text_color=self._C_TEXT).pack(side=tk.LEFT)
        self._tracker_var = tk.StringVar(
            value="Deep" if self._tracker_backend == "vit" else "Classic")
        ctk.CTkSegmentedButton(
            eng_row, values=["Deep", "Classic"],
            variable=self._tracker_var, command=self._on_tracker_backend_change,
            selected_color=self._C_ACCENT, selected_hover_color=self._C_ACCENT_H,
            unselected_color=self._C_SURFACE2, unselected_hover_color="#4a4a4a",
            font=ctk.CTkFont(size=11, weight="bold"), text_color=self._C_TEXT,
        ).pack(side=tk.RIGHT)
        ctk.CTkLabel(
            perf_frame,
            text=tr("Classic = the proven tracker (default). Deep = an experimental neural "
                 "tracker — more robust in clean tests, but it can balloon its box on "
                 "smooth targets; still being validated. Stick with Classic for now."),
            font=ctk.CTkFont(size=10), text_color=self._C_TEXT_DIM,
            wraplength=440, justify="left").pack(padx=12, pady=(0, 8), anchor="w")
        perf_row = ctk.CTkFrame(perf_frame, fg_color="transparent")
        perf_row.pack(fill=tk.X, padx=12, pady=(10, 4))
        ctk.CTkLabel(perf_row, text=tr("Tracking resolution"), font=lbl,
                     text_color=self._C_TEXT).pack(side=tk.LEFT)
        ctk.CTkSegmentedButton(
            perf_row, values=list(self._PROC_PRESETS.keys()),
            variable=self.proc_res_var, command=self._on_proc_res_change,
            selected_color=self._C_ACCENT, selected_hover_color=self._C_ACCENT_H,
            unselected_color=self._C_SURFACE2, unselected_hover_color="#4a4a4a",
            font=ctk.CTkFont(size=11, weight="bold"), text_color=self._C_TEXT,
        ).pack(side=tk.RIGHT)
        ctk.CTkLabel(
            perf_frame,
            text=tr("Lower = faster on slow PCs. Downscales the frame before tracking, "
                 "not the video you watch. If your FPS is low, try Fast or Fastest."),
            font=ctk.CTkFont(size=10), text_color=self._C_TEXT_DIM,
            wraplength=440, justify="left").pack(padx=12, pady=(0, 10), anchor="w")

        # Lock by look (experimental) — tucked here on purpose: it only works with a
        # genuinely distinct-colour marker, and can't separate a red/yellow/black
        # target from skin. Off by default.
        cl_row = ctk.CTkFrame(perf_frame, fg_color="transparent")
        cl_row.pack(fill=tk.X, padx=12, pady=(2, 0))
        self._color_lock_var = tk.BooleanVar(value=self._color_lock)
        ctk.CTkSwitch(
            cl_row, text=tr("Lock by look (experimental)"),
            variable=self._color_lock_var, command=self._on_color_lock_toggle,
            font=ctk.CTkFont(size=11), text_color=self._C_TEXT,
            button_color=self._C_ACCENT, fg_color=self._C_SURFACE2,
            progress_color=self._C_ACCENT).pack(side=tk.LEFT)
        self._color_swatch_lbl = ctk.CTkFrame(cl_row, width=16, height=16,
                                              corner_radius=3, fg_color=self._C_SURFACE2)
        self._color_swatch_lbl.pack(side=tk.LEFT, padx=(6, 0))
        ctk.CTkLabel(
            perf_frame,
            text=tr("Tracks by colour/darkness instead of the tracker. Only worth it with a "
                 "genuinely distinct marker — a matte lime / teal / hot-pink sticker. It "
                 "can't tell a red/yellow/black target from skin. Off = normal tracking."),
            font=ctk.CTkFont(size=10), text_color=self._C_TEXT_DIM,
            wraplength=440, justify="left").pack(padx=12, pady=(0, 10), anchor="w")

        ctk.CTkLabel(sf, text=tr("xToys"),
                     font=lbl, text_color=self._C_TEXT).pack(padx=16, pady=(8, 4), anchor="w")
        xtoys_frame = ctk.CTkFrame(sf, fg_color=self._C_SURFACE, corner_radius=8)
        xtoys_frame.pack(fill=tk.X, padx=16, pady=(0, 8))

        xtoys_help_row = ctk.CTkFrame(xtoys_frame, fg_color="transparent")
        xtoys_help_row.pack(fill=tk.X, padx=12, pady=(8, 4))
        ctk.CTkLabel(xtoys_help_row, text=tr("First time setup?"),
                     font=ctk.CTkFont(size=9), text_color=self._C_TEXT_DIM).pack(side=tk.LEFT)

        def _show_xtoys_help():
            hw = ctk.CTkToplevel(win)
            hw.title(tr("xToys Setup"))
            hw.geometry("420x320")
            hw.resizable(False, False)
            hw.grab_set()
            ctk.CTkLabel(hw,
                text=tr("Setup steps:\n\n"
                     "1. Open xtoys.app in a browser and sign in\n"
                     "2. Scripts \u2192 search \u201cVisualStimEdger\u201d \u2192 Load Script\n"
                     "3. Connections \u2192 add your toy under Generic Output\n"
                     "4. Go to xtoys.app/me \u2192 Private Webhook\n"
                     "5. Copy the Webhook ID \u2192 paste it into VSE\n"
                     "6. Press the green \u25b6 to RUN the script \u2014 nothing moves until you do\n"
                     "7. Keep the xToys tab open while you use VSE"),
                font=ctk.CTkFont(size=11), text_color=self._C_TEXT,
                justify="left", anchor="w", wraplength=380,
            ).pack(padx=20, pady=20, anchor="w")
            ctk.CTkButton(hw, text=tr("Close"), command=hw.destroy, width=80,
                          fg_color=self._C_SURFACE2, hover_color="#4a4a4a",
                          text_color=self._C_TEXT).pack(pady=(0, 16))

        ctk.CTkButton(xtoys_help_row, text=tr("? Setup Guide"), command=_show_xtoys_help,
                      width=100, height=22, font=ctk.CTkFont(size=9),
                      fg_color=self._C_SURFACE2, hover_color="#4a4a4a",
                      text_color=self._C_TEXT_DIM, border_width=1,
                      border_color=self._C_BORDER).pack(side=tk.LEFT, padx=(8, 0))


        # ── Cum Volume Override ───────────────────────────────────────────────
        ctk.CTkLabel(sf, text=tr("Cum Volume Behavior"),
                     font=lbl, text_color=self._C_TEXT).pack(padx=16, pady=(8, 4), anchor="w")
        cum_vol_frame = ctk.CTkFrame(sf, fg_color=self._C_SURFACE, corner_radius=8)
        cum_vol_frame.pack(fill=tk.X, padx=16, pady=(0, 8))
        override_var = tk.BooleanVar(value=self._cum_override_range)
        ctk.CTkRadioButton(cum_vol_frame, text=tr("Override to 100% volume (ignore range)"),
                           variable=override_var, value=True,
                           font=ctk.CTkFont(size=11), text_color=self._C_TEXT,
                           fg_color=self._C_ACCENT, hover_color=self._C_ACCENT_H,
                           border_color=self._C_BORDER
                           ).pack(padx=16, pady=(8, 2), anchor="w")
        ctk.CTkRadioButton(cum_vol_frame, text=tr("Respect volume ceiling (stay within range)"),
                           variable=override_var, value=False,
                           font=ctk.CTkFont(size=11), text_color=self._C_TEXT,
                           fg_color=self._C_ACCENT, hover_color=self._C_ACCENT_H,
                           border_color=self._C_BORDER
                           ).pack(padx=16, pady=(2, 8), anchor="w")

        # ── Refractory Period ─────────────────────────────────────────────────
        ctk.CTkLabel(sf, text=tr("Refractory Period"),
                     font=lbl, text_color=self._C_TEXT).pack(padx=16, pady=(8, 4), anchor="w")
        refrac_frame = ctk.CTkFrame(sf, fg_color=self._C_SURFACE, corner_radius=8)
        refrac_frame.pack(fill=tk.X, padx=16, pady=(0, 8))

        refrac_hdr = ctk.CTkFrame(refrac_frame, fg_color="transparent")
        refrac_hdr.pack(fill=tk.X, padx=12, pady=(8, 2))
        ctk.CTkLabel(refrac_hdr, text=tr("Cooldown after I've CUM"),
                     font=ctk.CTkFont(size=10, weight="bold"),
                     text_color=self._C_TEXT_DIM, anchor="w").pack(side=tk.LEFT)
        refrac_val_lbl = ctk.CTkLabel(
            refrac_hdr,
            text="Off" if self._refractory_mins == 0 else f"{self._refractory_mins} min",
            font=lbl, text_color=self._C_ACCENT, anchor="e")
        refrac_val_lbl.pack(side=tk.RIGHT)

        refrac_var = tk.IntVar(value=self._refractory_mins)

        def _on_refrac_slide(v):
            val = int(float(v))
            refrac_val_lbl.configure(text="Off" if val == 0 else f"{val} min")

        ctk.CTkSlider(
            refrac_frame, from_=0, to=30, number_of_steps=30,
            variable=refrac_var, command=_on_refrac_slide,
            button_color=self._C_ACCENT, button_hover_color=self._C_ACCENT_H,
            progress_color=self._C_ACCENT, fg_color=self._C_SURFACE2,
        ).pack(fill=tk.X, padx=12, pady=(0, 4))
        ctk.CTkLabel(
            refrac_frame,
            text=tr("How long after cumming before the session auto-resumes. "
                 "Set to 0 to disable the timer and resume manually. "
                 "Clicking the button during refractory skips it early."),
            font=ctk.CTkFont(size=9), text_color=self._C_TEXT_DIM,
            wraplength=420, justify="left", anchor="w",
        ).pack(fill=tk.X, padx=12, pady=(0, 10))

        # ── Edging behaviour (volume recovery) ────────────────────────────────
        ctk.CTkLabel(sf, text=tr("Edging behaviour"),
                     font=lbl, text_color=self._C_TEXT).pack(padx=16, pady=(8, 4), anchor="w")
        edge_frame = ctk.CTkFrame(sf, fg_color=self._C_SURFACE, corner_radius=8)
        edge_frame.pack(fill=tk.X, padx=16, pady=(0, 8))

        recover_var = tk.BooleanVar(value=self._erect_recovery)
        ctk.CTkSwitch(
            edge_frame, text=tr("Recover volume when held in the erect zone"),
            variable=recover_var,
            font=ctk.CTkFont(size=10), text_color=self._C_TEXT_DIM,
            switch_width=36, switch_height=16,
            button_color=self._C_ACCENT, fg_color=self._C_SURFACE2,
            progress_color=self._C_ACCENT,
        ).pack(anchor="w", padx=12, pady=(8, 2))
        ctk.CTkLabel(
            edge_frame,
            text=tr("Per-difficulty tuning below. Recovery = how fast volume climbs back "
                 "while you hold in the erect zone (%/sec). Hold = seconds off the edge "
                 "before Active Edging ramps. Ramp = how fast it climbs to max (%/sec). "
                 "Defaults get more relentless on harder levels."),
            font=ctk.CTkFont(size=9), text_color=self._C_TEXT_DIM,
            wraplength=420, justify="left", anchor="w",
        ).pack(fill=tk.X, padx=12, pady=(0, 6))

        # Column header
        _hdr = ctk.CTkFrame(edge_frame, fg_color="transparent")
        _hdr.pack(fill=tk.X, padx=12, pady=(0, 0))
        ctk.CTkLabel(_hdr, text="", width=64, anchor="w").pack(side=tk.LEFT)
        for _c in ("Recover %/s", "Hold s", "Ramp %/s"):
            ctk.CTkLabel(_hdr, text=tr(_c), font=ctk.CTkFont(size=9),
                         text_color=self._C_TEXT_DIM, width=72,
                         anchor="center").pack(side=tk.LEFT, padx=4)

        rec_vars, hold_vars, ramp_vars = {}, {}, {}
        for level in ("Easy", "Middle", "Hard", "Expert"):
            row = ctk.CTkFrame(edge_frame, fg_color="transparent")
            row.pack(fill=tk.X, padx=12, pady=2)
            ctk.CTkLabel(row, text=tr(level), font=lbl, text_color=self._C_TEXT,
                         width=64, anchor="w").pack(side=tk.LEFT)
            rec_vars[level]  = tk.DoubleVar(value=self._recovery_pct.get(level, 1.6))
            hold_vars[level] = tk.IntVar(value=self._ae_hold_secs.get(level, 10))
            ramp_vars[level] = tk.DoubleVar(value=self._ae_ramp_pct.get(level, 8.0))
            for _v in (rec_vars[level], hold_vars[level], ramp_vars[level]):
                ctk.CTkEntry(row, textvariable=_v, width=72, justify="center",
                             fg_color=self._C_SURFACE2, border_color=self._C_BORDER,
                             text_color=self._C_TEXT).pack(side=tk.LEFT, padx=4)

        # ── Auto-Cum Detection ────────────────────────────────────────────────
        ctk.CTkLabel(sf, text=tr("Auto-Cum Detection"),
                     font=lbl, text_color=self._C_TEXT).pack(padx=16, pady=(8, 4), anchor="w")
        acd_frame = ctk.CTkFrame(sf, fg_color=self._C_SURFACE, corner_radius=8)
        acd_frame.pack(fill=tk.X, padx=16, pady=(0, 8))

        # Delay slider
        acd_delay_hdr = ctk.CTkFrame(acd_frame, fg_color="transparent")
        acd_delay_hdr.pack(fill=tk.X, padx=12, pady=(8, 2))
        ctk.CTkLabel(acd_delay_hdr, text=tr("Countdown delay"),
                     font=ctk.CTkFont(size=10, weight="bold"),
                     text_color=self._C_TEXT_DIM, anchor="w").pack(side=tk.LEFT)
        acd_delay_val = ctk.CTkLabel(
            acd_delay_hdr,
            text="Off" if self._auto_cum_delay == 0 else f"{self._auto_cum_delay}s",
            font=lbl, text_color=self._C_ACCENT, anchor="e")
        acd_delay_val.pack(side=tk.RIGHT)
        acd_delay_var = tk.IntVar(value=self._auto_cum_delay)

        def _on_delay_slide(v):
            val = int(float(v))
            acd_delay_val.configure(text="Off" if val == 0 else f"{val}s")

        ctk.CTkSlider(
            acd_frame, from_=0, to=10, number_of_steps=10,
            variable=acd_delay_var, command=_on_delay_slide,
            button_color=self._C_ACCENT, button_hover_color=self._C_ACCENT_H,
            progress_color=self._C_ACCENT, fg_color=self._C_SURFACE2,
        ).pack(fill=tk.X, padx=12, pady=(0, 6))

        # Sensitivity slider
        acd_sens_hdr = ctk.CTkFrame(acd_frame, fg_color="transparent")
        acd_sens_hdr.pack(fill=tk.X, padx=12, pady=(0, 2))
        ctk.CTkLabel(acd_sens_hdr, text=tr("Detection sensitivity"),
                     font=ctk.CTkFont(size=10, weight="bold"),
                     text_color=self._C_TEXT_DIM, anchor="w").pack(side=tk.LEFT)
        acd_sens_val = ctk.CTkLabel(
            acd_sens_hdr, text=str(self._auto_cum_sensitivity),
            font=lbl, text_color=self._C_ACCENT, anchor="e")
        acd_sens_val.pack(side=tk.RIGHT)
        acd_sens_var = tk.IntVar(value=self._auto_cum_sensitivity)

        def _on_sens_slide(v):
            acd_sens_val.configure(text=str(int(float(v))))

        ctk.CTkSlider(
            acd_frame, from_=1, to=10, number_of_steps=9,
            variable=acd_sens_var, command=_on_sens_slide,
            button_color=self._C_ACCENT, button_hover_color=self._C_ACCENT_H,
            progress_color=self._C_ACCENT, fg_color=self._C_SURFACE2,
        ).pack(fill=tk.X, padx=12, pady=(0, 4))
        ctk.CTkLabel(
            acd_frame,
            text=tr("Higher sensitivity = triggers on less motion change. "
                 "If it fires too easily, lower this. "
                 "The countdown delay lets you cancel false positives."),
            font=ctk.CTkFont(size=9), text_color=self._C_TEXT_DIM,
            wraplength=420, justify="left", anchor="w",
        ).pack(fill=tk.X, padx=12, pady=(0, 10))

        # ── Hands-Free "Let Me Cum?" ──────────────────────────────────────────
        ctk.CTkLabel(sf, text=tr("Hands-Free \"Let Me Cum?\""),
                     font=lbl, text_color=self._C_TEXT).pack(padx=16, pady=(8, 4), anchor="w")
        hf_frame = ctk.CTkFrame(sf, fg_color=self._C_SURFACE, corner_radius=8)
        hf_frame.pack(fill=tk.X, padx=16, pady=(0, 8))

        # Toggle
        hf_toggle_row = ctk.CTkFrame(hf_frame, fg_color="transparent")
        hf_toggle_row.pack(fill=tk.X, padx=12, pady=(10, 4))
        hf_var = tk.BooleanVar(value=self._hf_enabled)
        hf_switch = ctk.CTkSwitch(
            hf_toggle_row, text=tr("Enable hands-free mode"),
            variable=hf_var,
            font=ctk.CTkFont(size=11), text_color=self._C_TEXT,
            switch_width=36, switch_height=18,
            button_color=self._C_ACCENT, fg_color=self._C_SURFACE2,
        )
        hf_switch.pack(side=tk.LEFT)

        ctk.CTkFrame(hf_frame, height=1, fg_color=self._C_BORDER
                     ).pack(fill=tk.X, padx=12, pady=(6, 4))

        # Column headers
        hdr_row = ctk.CTkFrame(hf_frame, fg_color="transparent")
        hdr_row.pack(fill=tk.X, padx=12, pady=(0, 2))
        ctk.CTkLabel(hdr_row, text=tr("Level"), font=ctk.CTkFont(size=9),
                     text_color=self._C_TEXT_DIM, width=70, anchor="w").pack(side=tk.LEFT)
        ctk.CTkLabel(hdr_row, text=tr("Min edges"), font=ctk.CTkFont(size=9),
                     text_color=self._C_TEXT_DIM, width=80, anchor="w").pack(side=tk.LEFT)
        ctk.CTkLabel(hdr_row, text=tr("1-in-N chance"), font=ctk.CTkFont(size=9),
                     text_color=self._C_TEXT_DIM, anchor="w").pack(side=tk.LEFT)

        hf_min_vars    = {}
        hf_chance_vars = {}
        for level in ("Easy", "Middle", "Hard", "Expert"):
            row = ctk.CTkFrame(hf_frame, fg_color="transparent")
            row.pack(fill=tk.X, padx=12, pady=3)
            ctk.CTkLabel(row, text=tr(level), font=lbl, text_color=self._C_TEXT,
                         width=70, anchor="w").pack(side=tk.LEFT)
            min_var = tk.IntVar(value=self._hf_min_edges.get(level, 0))
            hf_min_vars[level] = min_var
            ctk.CTkEntry(row, textvariable=min_var, width=56,
                         fg_color=self._C_SURFACE2, border_color=self._C_BORDER,
                         text_color=self._C_TEXT).pack(side=tk.LEFT, padx=(0, 12))
            chance_var = tk.IntVar(value=self._hf_cum_chance.get(level, 10))
            hf_chance_vars[level] = chance_var
            ctk.CTkLabel(row, text=tr("1 in"), font=ctk.CTkFont(size=10),
                         text_color=self._C_TEXT_DIM).pack(side=tk.LEFT, padx=(0, 4))
            ctk.CTkEntry(row, textvariable=chance_var, width=56,
                         fg_color=self._C_SURFACE2, border_color=self._C_BORDER,
                         text_color=self._C_TEXT).pack(side=tk.LEFT)

        ctk.CTkLabel(
            hf_frame,
            text=tr("When enabled, VSE automatically rolls for cum permission each time an edge "
                 "is detected. 'Min edges' sets how many edges must occur first. "
                 "'1 in N' is the per-edge roll chance (lower N = more likely). "
                 "Uses the same grant window as a manual roll."),
            font=ctk.CTkFont(size=9), text_color=self._C_TEXT_DIM,
            wraplength=420, justify="left", anchor="w",
        ).pack(fill=tk.X, padx=12, pady=(4, 10))

        # ── Bondage Mode Defaults ─────────────────────────────────────────────
        ctk.CTkLabel(sf, text=tr("🎙 Bondage Mode Defaults"),
                     font=lbl, text_color=self._C_TEXT).pack(padx=16, pady=(8, 4), anchor="w")
        bm_frame = ctk.CTkFrame(sf, fg_color=self._C_SURFACE, corner_radius=8)
        bm_frame.pack(fill=tk.X, padx=16, pady=(0, 8))

        bm_sw_row = ctk.CTkFrame(bm_frame, fg_color="transparent")
        bm_sw_row.pack(fill=tk.X, padx=12, pady=(10, 4))
        ctk.CTkLabel(bm_sw_row, text=tr("Default safeword:"), font=ctk.CTkFont(size=10, weight="bold"),
                     text_color=self._C_TEXT_DIM, width=130, anchor="w").pack(side=tk.LEFT)
        bm_sw_var = tk.StringVar(value=self._bondage_safeword)
        ctk.CTkEntry(bm_sw_row, textvariable=bm_sw_var, width=160,
                     fg_color=self._C_SURFACE2, border_color=self._C_BORDER,
                     text_color=self._C_TEXT,
                     placeholder_text="red").pack(side=tk.LEFT)

        bm_mic_row = ctk.CTkFrame(bm_frame, fg_color="transparent")
        bm_mic_row.pack(fill=tk.X, padx=12, pady=(4, 10))
        ctk.CTkLabel(bm_mic_row, text=tr("Default mic device:"), font=ctk.CTkFont(size=10, weight="bold"),
                     text_color=self._C_TEXT_DIM, width=130, anchor="w").pack(side=tk.LEFT)
        _bm_mics   = ["System Default"] + VoiceEngine.list_input_devices()
        _bm_cur    = self._bondage_mic_device if self._bondage_mic_device in _bm_mics else "System Default"
        bm_mic_var = tk.StringVar(value=_bm_cur)
        ctk.CTkComboBox(bm_mic_row, values=_bm_mics, variable=bm_mic_var, width=240,
                        fg_color=self._C_SURFACE2, border_color=self._C_BORDER,
                        button_color=self._C_ACCENT, button_hover_color=self._C_ACCENT_H,
                        dropdown_fg_color=self._C_SURFACE2, text_color=self._C_TEXT,
                        ).pack(side=tk.LEFT)

        # ── Ruin Odds ─────────────────────────────────────────────────────────
        ctk.CTkLabel(sf, text=tr("😈 Ruin Odds (% chance when denied, Evil Mode only)"),
                     font=lbl, text_color=self._C_TEXT).pack(padx=16, pady=(8, 4), anchor="w")
        ruin_odds_frame = ctk.CTkFrame(sf, fg_color=self._C_SURFACE, corner_radius=8)
        ruin_odds_frame.pack(fill=tk.X, padx=16, pady=(0, 8))

        ruin_odds_vars = {}
        for level in ("Easy", "Middle", "Hard", "Expert"):
            row = ctk.CTkFrame(ruin_odds_frame, fg_color="transparent")
            row.pack(fill=tk.X, padx=12, pady=4)
            ctk.CTkLabel(row, text=tr(level), font=lbl, text_color=self._C_TEXT,
                         width=70, anchor="w").pack(side=tk.LEFT)
            ctk.CTkLabel(row, text="%", font=ctk.CTkFont(size=10),
                         text_color=self._C_TEXT_DIM).pack(side=tk.LEFT, padx=(0, 4))
            var = tk.IntVar(value=self._ruin_odds.get(level, 0))
            ruin_odds_vars[level] = var
            ctk.CTkEntry(row, textvariable=var, width=60,
                         fg_color=self._C_SURFACE2, border_color=self._C_BORDER,
                         text_color=self._C_TEXT).pack(side=tk.LEFT)

        # ── Ruin Phrases ──────────────────────────────────────────────────────
        ctk.CTkLabel(sf, text=tr("Ruin Phrases (one per line)"),
                     font=lbl, text_color=self._C_TEXT).pack(padx=16, pady=(8, 4), anchor="w")
        ruin_phrases_box = ctk.CTkTextbox(sf, height=150,
                                          fg_color=self._C_SURFACE, border_color=self._C_BORDER,
                                          text_color=self._C_TEXT, border_width=1,
                                          font=ctk.CTkFont(size=11))
        ruin_phrases_box.pack(fill=tk.BOTH, padx=16, pady=(0, 8))
        ruin_phrases_box.insert("1.0", "\n".join(self._ruin_phrases))

        # ── Cum Odds ──────────────────────────────────────────────────────────
        ctk.CTkLabel(sf, text=tr("\"Let me cum?\" Odds  (1 in N chance)"),
                     font=lbl, text_color=self._C_TEXT).pack(padx=16, pady=(12, 4), anchor="w")
        odds_frame = ctk.CTkFrame(sf, fg_color=self._C_SURFACE, corner_radius=8)
        odds_frame.pack(fill=tk.X, padx=16, pady=(0, 8))

        odds_vars = {}
        for level in ("Easy", "Middle", "Hard", "Expert"):
            row = ctk.CTkFrame(odds_frame, fg_color="transparent")
            row.pack(fill=tk.X, padx=12, pady=4)
            ctk.CTkLabel(row, text=tr(level), font=lbl, text_color=self._C_TEXT,
                         width=70, anchor="w").pack(side=tk.LEFT)
            ctk.CTkLabel(row, text=tr("1 in"), font=ctk.CTkFont(size=10),
                         text_color=self._C_TEXT_DIM).pack(side=tk.LEFT, padx=(0, 4))
            var = tk.IntVar(value=self._cum_odds.get(level, 4))
            odds_vars[level] = var
            ctk.CTkEntry(row, textvariable=var, width=60,
                         fg_color=self._C_SURFACE2, border_color=self._C_BORDER,
                         text_color=self._C_TEXT).pack(side=tk.LEFT)

        # ── Denial Phrases ────────────────────────────────────────────────────
        ctk.CTkLabel(sf, text=tr("Denial Phrases  (one per line)"),
                     font=lbl, text_color=self._C_TEXT).pack(padx=16, pady=(8, 4), anchor="w")
        phrases_box = ctk.CTkTextbox(sf, height=220,
                                     fg_color=self._C_SURFACE, border_color=self._C_BORDER,
                                     text_color=self._C_TEXT, border_width=1,
                                     font=ctk.CTkFont(size=11))
        phrases_box.pack(fill=tk.BOTH, padx=16, pady=(0, 8))
        phrases_box.insert("1.0", "\n".join(self._denial_phrases))

        # ── Buttons ───────────────────────────────────────────────────────────
        btn_row = ctk.CTkFrame(win, fg_color="transparent")
        btn_row.pack(fill=tk.X, padx=16, pady=(0, 12))

        def _cancel():
            """Revert reversible settings to their pre-dialog values and close."""
            self._cum_odds          = _orig["cum_odds"]
            self._ruin_odds         = _orig["ruin_odds"]
            self._denial_phrases    = _orig["denial_phrases"]
            self._ruin_phrases      = _orig["ruin_phrases"]
            self._cum_override_range   = _orig["cum_override_range"]
            self._refractory_mins      = _orig["refractory_mins"]
            self._erect_recovery       = _orig["erect_recovery"]
            self._recovery_pct         = _orig["recovery_pct"]
            self._ae_hold_secs         = _orig["ae_hold_secs"]
            self._ae_ramp_pct          = _orig["ae_ramp_pct"]
            self._auto_cum_delay       = _orig["auto_cum_delay"]
            self._auto_cum_sensitivity = _orig["auto_cum_sensitivity"]
            self._bondage_safeword     = _orig["bondage_safeword"]
            self._hf_enabled           = _orig["hf_enabled"]
            self._hf_min_edges         = _orig["hf_min_edges"]
            self._hf_cum_chance        = _orig["hf_cum_chance"]
            # Note: theme, font size, and output device are NOT reverted —
            # they take effect immediately or require a restart.
            win.destroy()

        def _save():
            for level, var in odds_vars.items():
                val = var.get()
                if val >= 1:
                    self._cum_odds[level] = val
            for level, var in ruin_odds_vars.items():
                val = var.get()
                if 0 <= val <= 100:
                    self._ruin_odds[level] = val
            text = phrases_box.get("1.0", "end").strip()
            self._denial_phrases = [l.strip() for l in text.split("\n") if l.strip()]
            ruin_text = ruin_phrases_box.get("1.0", "end").strip()
            self._ruin_phrases = [l.strip() for l in ruin_text.split("\n") if l.strip()]
            self._cum_override_range   = override_var.get()
            self._refractory_mins      = int(refrac_var.get())
            self._erect_recovery       = bool(recover_var.get())
            for _lvl in ("Easy", "Middle", "Hard", "Expert"):
                try: self._recovery_pct[_lvl] = max(0.0, float(rec_vars[_lvl].get()))
                except Exception: pass
                try: self._ae_hold_secs[_lvl] = max(1, min(120, int(hold_vars[_lvl].get())))
                except Exception: pass
                try: self._ae_ramp_pct[_lvl]  = max(0.0, float(ramp_vars[_lvl].get()))
                except Exception: pass
            self._auto_cum_delay       = int(acd_delay_var.get())
            self._auto_cum_sensitivity = int(acd_sens_var.get())
            # bondage defaults
            _bsw = bm_sw_var.get().strip().lower()
            if _bsw:
                self._bondage_safeword = _bsw
            _bmic = bm_mic_var.get()
            self._bondage_mic_device = "" if _bmic == "System Default" else _bmic
            self._hf_enabled = hf_var.get()
            for level, var in hf_min_vars.items():
                try:
                    self._hf_min_edges[level] = max(0, int(var.get()))
                except Exception:
                    pass
            for level, var in hf_chance_vars.items():
                try:
                    self._hf_cum_chance[level] = max(1, int(var.get()))
                except Exception:
                    pass
            self._ui_font_size         = int(font_var.get())
            self._theme_name = theme_var.get()
            self._save_config()
            win.destroy()

        def _reset():
            self._cum_odds = dict(self._CUM_ODDS_DEFAULT)
            self._denial_phrases = list(self._DENIAL_PHRASES_DEFAULT)
            self._ruin_odds = dict(self._RUIN_ODDS_DEFAULT)
            self._ruin_phrases = list(self._RUIN_PHRASES_DEFAULT)
            for level, var in odds_vars.items():
                var.set(self._CUM_ODDS_DEFAULT.get(level, 4))
            for level, var in ruin_odds_vars.items():
                var.set(self._RUIN_ODDS_DEFAULT.get(level, 0))
            for level, var in hf_min_vars.items():
                var.set(self._HF_MIN_EDGES_DEFAULT.get(level, 0))
            for level, var in hf_chance_vars.items():
                var.set(self._HF_CUM_CHANCE_DEFAULT.get(level, 10))
            hf_var.set(False)
            phrases_box.delete("1.0", "end")
            phrases_box.insert("1.0", "\n".join(self._denial_phrases))
            ruin_phrases_box.delete("1.0", "end")
            ruin_phrases_box.insert("1.0", "\n".join(self._ruin_phrases))

        ctk.CTkButton(btn_row, text=tr("Reset Defaults"), command=_reset, width=120,
                      fg_color=self._C_SURFACE2, hover_color="#4a4a4a",
                      text_color=self._C_TEXT, border_width=1, border_color=self._C_BORDER,
                      font=ctk.CTkFont(size=10)).pack(side=tk.LEFT)
        ctk.CTkButton(btn_row, text=tr("Cancel"), command=_cancel, width=90,
                      fg_color=self._C_SURFACE2, hover_color="#4a4a4a",
                      text_color=self._C_TEXT_DIM, border_width=1, border_color=self._C_BORDER,
                      font=ctk.CTkFont(size=11)).pack(side=tk.RIGHT, padx=(4, 0))
        ctk.CTkButton(btn_row, text=tr("OK"), command=_save, width=100,
                      fg_color=self._C_ACCENT, hover_color=self._C_ACCENT_H,
                      text_color="white", font=ctk.CTkFont(size=12, weight="bold")
                      ).pack(side=tk.RIGHT)
        win.protocol("WM_DELETE_WINDOW", _cancel)

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
            if self._hr_source == "pulsoid":
                tok = self.hr_token_var.get().strip()
                if tok:
                    self.hr_client.resting_bpm = self.hr_resting_var.get()
                    self.hr_client.peak_bpm    = self.hr_peak_var.get()
                    self.hr_client.token       = tok
                    self.hr_client.start()
            else:
                if self._ble_addr:
                    self.ble_hr_client.resting_bpm = int(self.hr_resting_var.get())
                    self.ble_hr_client.peak_bpm    = int(self.hr_peak_var.get())
                    self.ble_hr_client.start(self._ble_addr, self._ble_name)
        else:
            self.hr_client.stop()
            self.ble_hr_client.stop()
        if self.audio_on.get():
            self._windows_opts.pack(fill=tk.X, padx=P, pady=(0, 4), after=after)
            after = self._windows_opts
            if not self.win_devices:
                self._refresh_devices()
        if self.mp3_on.get() and _MINIAUDIO_OK:
            self._mp3_opts.pack(fill=tk.X, padx=P, pady=(0, 4), after=after)
        if not getattr(self, '_loading_config', False):
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
            # Mask — the webhook ID is a device-control capability; vse.log is the
            # file users paste into bug reports.
            _masked = f"…{new_id[-4:]}" if len(new_id) >= 4 else "(set)"
            log.info(f"xToys: webhook ID updated ({_masked})")
            self.xtoys.start_listener()
        self._update_xtoys_id_warning(new_id)
        self._save_config()

    def _xtoys_id_looks_wrong(self, s):
        """Heuristic: does this look like a pasted URL / host:port instead of a
        Webhook ID? A real ID is a short alphanumeric token, e.g. 8hR5acKTCx2s.
        Returns a warning string, or None if it looks fine (or is empty)."""
        s = (s or "").strip()
        if not s:
            return None
        low = s.lower()
        if any(b in low for b in ("http", "://", "localhost", "127.0.0.1", "xtoys.app")):
            return ("That's a URL / address, not a Webhook ID. Get the short "
                    "code from xtoys.app/me → Private Webhook (e.g. 8hR5acKTCx2s).")
        if "/" in s or ":" in s or " " in s:
            return ("That doesn't look like a Webhook ID — paste just the short "
                    "code from xtoys.app/me → Private Webhook (e.g. 8hR5acKTCx2s), no URL or slashes.")
        return None

    def _update_xtoys_id_warning(self, new_id):
        lbl = getattr(self, "_xtoys_warn", None)
        if lbl is None:
            return   # panel not built yet (e.g. during initial config load)
        warn = self._xtoys_id_looks_wrong(new_id)
        ent = getattr(self, "_xtoys_entry", None)
        if warn:
            lbl.configure(text="⚠ " + warn)
            if not getattr(self, "_xtoys_warn_shown", False):
                lbl.pack(anchor="w", padx=12, pady=(0, 6))
                self._xtoys_warn_shown = True
            if ent is not None:
                try: ent.configure(border_color="#ff5555")
                except Exception: pass
        else:
            if getattr(self, "_xtoys_warn_shown", False):
                lbl.pack_forget()
                self._xtoys_warn_shown = False
            if ent is not None:
                try: ent.configure(border_color=self._C_BORDER)
                except Exception: pass

    def _on_hr_source_change(self, value: str):
        self._hr_source = "ble" if value == "BLE Direct" else "pulsoid"
        # Stop whichever was running
        self.hr_client.stop()
        self.ble_hr_client.stop()
        if self._hr_source == "ble":
            # Show BLE row, hide pulsoid token rows
            self._ble_row.pack(fill=tk.X)
            if self._ble_addr:
                self.ble_hr_client.resting_bpm = int(self.hr_resting_var.get())
                self.ble_hr_client.peak_bpm    = int(self.hr_peak_var.get())
                self.ble_hr_client.start(self._ble_addr, self._ble_name)
        else:
            self._ble_row.pack_forget()
            tok = self.hr_token_var.get().strip()
            if tok:
                self.hr_client.start()
        self._save_config()

    def _on_hr_token_change(self, *_):
        new_token = self.hr_token_var.get().strip()
        if new_token != self.hr_client.token and self.hr_on.get():
            self.hr_client.restart(token=new_token)
        elif new_token != self.hr_client.token:
            self.hr_client.token = new_token
        self._save_config()

    def _ble_scan_dialog(self):
        """Open a scan dialog, show found BLE HR devices, let user pick one."""
        dlg = ctk.CTkToplevel(self.root)
        dlg.title(tr("BLE HR — Scan"))
        dlg.geometry("340x280")
        dlg.grab_set()
        dlg.resizable(False, False)

        ctk.CTkLabel(dlg, text=tr("Scanning for BLE HR monitors…"),
                     font=ctk.CTkFont(size=12)).pack(pady=(18, 4))
        status_lbl = ctk.CTkLabel(dlg, text=tr("(up to 6 seconds)"),
                                   font=ctk.CTkFont(size=10),
                                   text_color=self._C_TEXT_DIM)
        status_lbl.pack()

        listbox_frame = ctk.CTkScrollableFrame(dlg, height=120)
        listbox_frame.pack(fill=tk.X, padx=16, pady=8)

        btn_row = ctk.CTkFrame(dlg, fg_color="transparent")
        btn_row.pack(fill=tk.X, padx=16, pady=(0, 12))
        connect_btn = ctk.CTkButton(btn_row, text=tr("Connect"), width=100,
                                     fg_color=self._C_ACCENT, hover_color=self._C_ACCENT_H,
                                     text_color="white", state="disabled",
                                     command=lambda: None)
        connect_btn.pack(side=tk.RIGHT, padx=(4, 0))
        ctk.CTkButton(btn_row, text=tr("Cancel"), width=80,
                       fg_color=self._C_SURFACE2, hover_color="#4a4a4a",
                       text_color=self._C_TEXT,
                       command=dlg.destroy).pack(side=tk.RIGHT)

        selected = {"addr": None, "name": None}
        row_btns = []

        def _pick(addr, name, btn):
            selected["addr"] = addr
            selected["name"] = name
            for b in row_btns:
                b.configure(fg_color=self._C_SURFACE2)
            btn.configure(fg_color=self._C_ACCENT)
            connect_btn.configure(state="normal")

        def _connect():
            addr = selected["addr"]
            name = selected["name"]
            if not addr:
                return
            self._ble_addr = addr
            self._ble_name = name
            self._ble_name_lbl.configure(text=name, text_color=self._C_TEXT)
            self.ble_hr_client.stop()
            self.ble_hr_client.resting_bpm = int(self.hr_resting_var.get())
            self.ble_hr_client.peak_bpm    = int(self.hr_peak_var.get())
            self.ble_hr_client.start(addr, name)
            self._save_config()
            dlg.destroy()

        connect_btn.configure(command=_connect)

        def _do_scan():
            devices = BLEHRClient.scan_sync(timeout=6.0)
            def _update():
                if not dlg.winfo_exists():
                    return
                status_lbl.configure(text=f"Found {len(devices)} device(s)" if devices else "No devices found")
                for addr, name in devices:
                    btn = ctk.CTkButton(
                        listbox_frame, text=f"{name}  ({addr})",
                        fg_color=self._C_SURFACE2, hover_color="#4a4a4a",
                        text_color=self._C_TEXT, font=ctk.CTkFont(size=11),
                        anchor="w", height=28,
                    )
                    btn.configure(command=lambda a=addr, n=name, b=btn: _pick(a, n, b))
                    btn.pack(fill=tk.X, pady=1)
                    row_btns.append(btn)
            self._call_on_main(_update)

        threading.Thread(target=_do_scan, daemon=True).start()

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
            self._save_config()

    def _refresh_devices(self, select_name: str = None):
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
            # Prefer previously-selected name, fall back to first device
            target = select_name if select_name in names else names[0]
            self._device_combo.set(target)
            self._on_device_select(target)
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
            self._hold_btn.configure(text=tr("HELD [|]"),
                                     fg_color=self._C_RED, hover_color=self._C_RED_HOV,
                                     border_color=self._C_RED)
            if hasattr(self, '_play_hold_btn'):
                self._play_hold_btn.configure(text=tr("HELD [|]"),
                                              fg_color=self._C_RED, hover_color=self._C_RED_HOV)
        else:
            self._hold_btn.configure(text=tr("Hold Volume"),
                                     fg_color=self._C_SURFACE2, hover_color="#4a4a4a",
                                     border_color=self._C_BORDER)
            if hasattr(self, '_play_hold_btn'):
                self._play_hold_btn.configure(text=tr("Hold Volume"),
                                              fg_color=self._C_SURFACE2, hover_color="#4a4a4a")

    def _fit_window(self):
        """Fit root HEIGHT to content; WIDTH honours the user's remembered size.

        The content only needs ~560px across, but a widget below over-requests
        width, so letting reqwidth drive it booted the window too wide. Instead
        we pin width to the persisted value (or a sensible default that fits the
        status line) and only auto-fit height — which also drops the empty gap
        at the bottom. The window stays resizable, and any resize is remembered."""
        self.root.update_idletasks()
        h = self.root.winfo_reqheight()
        w = self._win_w or _DEFAULT_WIN_W
        self.root.geometry(f"{w}x{h + 24}")

    def _toggle_play_mode(self):
        self._play_mode = not self._play_mode
        if self._play_mode:
            self._sf.pack_forget()
            self._play_panel.pack(fill=tk.X)
            self._play_btn.configure(text="⚙")
            self.root.title(("👿 " if self._evil_mode else "") + "VisualStimEdger ▶" + (" [EVIL]" if self._evil_mode else ""))
        else:
            self._play_panel.pack_forget()
            self._sf.pack(fill=tk.X)
            self._play_btn.configure(text=tr("▶ Play"))
            self.root.title(("👿 " if self._evil_mode else "") + "VisualStimEdger" + (" [EVIL]" if self._evil_mode else ""))
        self.root.after(50, self._fit_window)

    def _toggle_evil_mode(self):
        self._evil_mode = not self._evil_mode
        self._apply_evil_mode(self._evil_mode)

    def _apply_evil_mode(self, on: bool):
        """Toggle Evil Mode — adds ruin outcome chance."""
        if on:
            # 1. Title
            self.root.title(tr("👿 VisualStimEdger [EVIL MODE]"))
            # 2. Window background — dark red tint (shows through transparent frames)
            self.root.configure(fg_color="#120005")
            # 3. Video border — red
            self._vid_shell.configure(border_width=2, border_color="#cc0000")
            # 4. Evil button — filled red
            self._evil_btn.configure(fg_color="#dd1100", hover_color="#aa0000",
                                     border_color="#ff2200", text_color="white")
            # 5. Let me cum? — red border
            if not self._cum_allowed:
                self._letmecum_btn.configure(fg_color="#8a1a1a", hover_color="#6a0a0a",
                                             border_width=2, border_color="#aa0000")
            else:
                self._letmecum_btn.configure(border_width=2, border_color="#aa0000")
            # Snark label pulse
            self._evil_pulse_snark(0)
            log.info("Evil Mode: on")
        else:
            # 1. Title
            self.root.title("VisualStimEdger")
            # 2. Restore window background
            self.root.configure(fg_color=self._C_BG)
            # 3. Restore video border
            self._vid_shell.configure(border_width=1, border_color=self._C_BORDER)
            # 4. Evil button — outline only
            self._evil_btn.configure(fg_color="transparent", hover_color="#3a0000",
                                     border_color="#cc0000", text_color="#cc0000")
            # 5. Restore let me cum? button
            self._letmecum_btn.configure(fg_color=self._C_GREEN, hover_color=self._C_GREEN_H,
                                         border_width=0)
            # Stop snark pulse
            if self._evil_pulse_job:
                self.root.after_cancel(self._evil_pulse_job)
                self._evil_pulse_job = None
            self._snark_label.configure(text_color="#ff4444")
            log.info("Evil Mode: off")
        self._save_config()

    def _evil_pulse_snark(self, phase: int = 0):
        """Oscillate snark label between two reds while evil mode is active."""
        if not self._evil_mode:
            return
        try:
            self._snark_label.configure(
                text_color="#ff2200" if phase % 2 == 0 else "#880000"
            )
        except Exception:
            pass
        self._evil_pulse_job = self.root.after(600, lambda: self._evil_pulse_snark(phase + 1))

    # ------------------------------------------------------------------ MP3 transport callbacks

    def _mp3_load_file(self):
        path = filedialog.askopenfilename(
            title=tr("Select audio file"),
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
        folder = filedialog.askdirectory(title=tr("Select music folder"))
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

    def _mp3_on_device_change(self, val):
        if self.music_player:
            self.music_player.set_output_device("" if val == "Default" else val)

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
    _PERF_LOG_INTERVAL_S  = 10.0 # how often to log p50/p95/p99 frame time
    _GC_INTERVAL_MS       = 10000 # main-thread cyclic-GC sweep (automatic GC is off)
    # Tracking-resolution presets (target width in px; 0 = native/off). Downscaling
    # the frame before CSRT is the main FPS lever on slow CPUs — the tracker is the
    # per-frame cost, and it scales with pixel count.
    _PROC_PRESETS = {"Full": 0, "Balanced": 960, "Fast": 640, "Fastest": 480}

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
            if self._last_frame_t:
                dt = (now - self._last_frame_t) * 1000.0   # ms since last frame
                if dt < 2000.0:        # skip pause/stall gaps, keep real spikes
                    self._frame_dt.append(dt)
            self._last_frame_t = now

            # Tracking + volume: every other frame in non-Expert modes
            self._proc_frame_count = self._proc_frame_count + 1
            heavy = (self.aggr_var.get() == "Expert" or self._proc_frame_count % 2 == 1)

            if heavy:
                with _CV2_LOCK:
                    if self._color_lock:
                        self._run_color_lock(frame)
                    else:
                        self._maybe_yolo_reanchor(frame)
                        self._run_tracker(frame)
                if self._auto_mode and self.tracking_ok:
                    self._auto_feed(self.head_y)
                # Panic cut runs on the fresh (unsmoothed) head_y every heavy
                # frame — ~15 fps — so it beats the 0.5s volume tick to the punch.
                self._check_ponr()

            state = self._determine_state(self.head_y)

            if heavy:
                self._update_stats(state)
                self._tick_volume()

            # Status label — throttled to _STATUS_INTERVAL_MS
            if now - self._last_status_time >= self._STATUS_INTERVAL_MS / 1000:
                self._update_status_label(state)
                self._last_status_time = now

            # Frame-time percentiles → log (avg FPS alone hides the spikes)
            if now - self._last_perf_log >= self._PERF_LOG_INTERVAL_S:
                self._log_perf()
                self._last_perf_log = now

            # OBS overlay broadcast — 4 Hz
            if now - self._last_overlay_broadcast >= 0.25:
                self._last_overlay_broadcast = now
                self._broadcast_overlay(state)

            # Video display — throttled to _DISPLAY_INTERVAL_MS
            if now - self._last_display_time >= self._DISPLAY_INTERVAL_MS / 1000:
                with _CV2_LOCK:
                    self._draw_height_lines(frame)
                    self._draw_tracking_overlay(frame)
                    self._display_frame(frame)
                self._last_display_time = now

        except Exception:
            log.exception("_update_frame: unhandled exception — loop continues")
        finally:
            interval = 100 if self.tracking_paused else 33  # ~30fps poll
            self.root.after(interval, self._update_frame)

    def _log_perf(self):
        """Log frame-time percentiles. Average FPS can read fine while p95/p99
        spikes make the toy stutter — this surfaces the spikes (Jarry's ask).
        Runs once per _PERF_LOG_INTERVAL_S; the sort is over <=300 samples."""
        dt = list(self._frame_dt)
        if len(dt) < 5:
            return
        dt.sort()
        n = len(dt)
        def pct(p):
            return dt[min(n - 1, int(round((p / 100.0) * (n - 1))))]
        mean = sum(dt) / n
        fps  = 1000.0 / mean if mean > 0 else 0.0
        backend = "Deep" if getattr(self, "_tracker_backend", "") == "vit" else "Classic"
        res = self.proc_res_var.get() if hasattr(self, "proc_res_var") else "?"
        log.info(f"perf: {fps:.1f} fps avg | frame ms p50={pct(50):.0f} "
                 f"p95={pct(95):.0f} p99={pct(99):.0f} max={dt[-1]:.0f} "
                 f"(n={n}) | tracker={backend} res={res}")

    def _on_lock_track_toggle(self):
        """Freeze/unfreeze YOLO reanchoring (the 'Lock on target' switch)."""
        self._reanchor_lock = bool(self._lock_track_var.get())
        self._save_config()
        if self._reanchor_lock:
            self.yolo_candidate = None   # drop any pending reanchor candidate
            self._lock_tpl_time = 0.0    # grab a fresh appearance patch next healthy frame
            self._snark_label.configure(
                text="🔒 Locked on target — re-finds it if lost, no auto-reanchor",
                text_color="#F5A623")
            log.info("Tracking lock ON — YOLO reanchoring disabled; appearance re-find enabled")
        else:
            self._lock_template = None   # drop the saved patch
            self._snark_label.configure(
                text=tr("🔓 Auto-reanchor back on"), text_color="#F5A623")
            log.info("Tracking lock OFF — YOLO reanchoring enabled")

    def _on_color_lock_toggle(self):
        """Toggle appearance lock. On enable, read the target from the selection box
        on the next frame and auto-pick colour or dark mode."""
        self._color_lock = bool(self._color_lock_var.get())
        self._save_config()
        if self._color_lock:
            self._color_needs_derive = True
            self._snark_label.configure(
                text="🎯 Lock by look on — reading your target…",
                text_color="#F5A623")
            log.info("Appearance lock ON — reading target from selection box")
        else:
            self._color_hue = None
            if hasattr(self, "_color_swatch_lbl"):
                self._color_swatch_lbl.configure(fg_color=self._C_SURFACE2)
            self._snark_label.configure(text="🎯 Lock by look off", text_color="#F5A623")
            log.info("Appearance lock OFF")

    def _derive_color_lock(self, frame):
        """Read the target from the selection box and pick a lock mode: HUE for a
        vividly-coloured marker, or DARK for a low-saturation target like a black
        rubber electrode on lighter skin. Sets the swatch and the mode's params."""
        x, y, w, h = self.last_bbox
        fh, fw = frame.shape[:2]
        x0, y0 = max(0, int(x)), max(0, int(y))
        x1, y1 = min(fw, int(x + w)), min(fh, int(y + h))
        if x1 - x0 < 4 or y1 - y0 < 4:
            return False
        # Sample the box AND a surrounding ring (box expanded ~60%) so we can lock the
        # colour that's DISTINCTIVE — present in the box but not in the surroundings —
        # instead of the box's average (which is mostly the skin the marker sits on).
        ex = int((x1 - x0) * 0.6) + 4
        ey = int((y1 - y0) * 0.6) + 4
        rx0, ry0 = max(0, x0 - ex), max(0, y0 - ey)
        rx1, ry1 = min(fw, x1 + ex), min(fh, y1 + ey)
        hsv = cv2.cvtColor(frame[ry0:ry1, rx0:rx1], cv2.COLOR_BGR2HSV)
        H, S, V = hsv[..., 0], hsv[..., 1], hsv[..., 2]
        inner = np.zeros(H.shape, bool)
        inner[y0 - ry0:y1 - ry0, x0 - rx0:x1 - rx0] = True
        sat = (S > 60) & (V > 50)
        in_sat = inner & sat
        box_frac = int(in_sat.sum()) / max(1, int(inner.sum()))
        if int(in_sat.sum()) >= 10 and box_frac >= 0.10:
            # HUE mode. Lock the hue over-represented inside the box vs the ring.
            ring = (~inner) & sat
            hi = np.bincount(H[in_sat], minlength=180).astype(float); hi /= hi.sum()
            if int(ring.sum()) >= 8:
                ho = np.bincount(H[ring], minlength=180).astype(float); ho /= ho.sum()
            else:
                ho = np.zeros(180)
            def _sm(a):                                       # circular ±smooth
                p = np.concatenate([a[-8:], a, a[:8]])
                return np.convolve(p, np.ones(9) / 9.0, "same")[8:-8]
            hue  = int(np.argmax(_sm(hi) - _sm(ho)))
            dist = np.abs(((H.astype(int) - hue + 90) % 180) - 90)
            band = in_sat & (dist <= self._COLOR_HUE_TOL)
            if int(band.sum()) < 5:
                band = in_sat
            ring_s = int(np.median(S[ring])) if int(ring.sum()) >= 8 else 40
            # Isolate the marker within that hue band: the pixels MORE saturated than
            # the surroundings, so same-hue skin doesn't drag the estimate down.
            marker = band & (S > ring_s + 20)
            if int(marker.sum()) >= 5:
                mk_s = int(np.median(S[marker])); mk_v = int(np.median(V[marker]))
            else:
                mk_s = int(np.percentile(S[band], 80)); mk_v = int(np.median(V[band]))
            self._lock_mode   = "hue"
            self._color_hue   = hue
            self._color_tol   = self._COLOR_HUE_TOL
            # Floor sits ABOVE the surroundings (skin) but below the marker, so a red
            # marker on red-ish skin is separated by saturation, not just hue.
            self._color_s_min = int(max(70, min((ring_s + mk_s) // 2, mk_s - 10, 200)))
            self._color_v_min = 45
            sw = (hue, max(mk_s, 120), max(mk_v, 120))
            log.info(f"Lock: HUE mode hue={hue} s_min={self._color_s_min} "
                     f"(ring_s={ring_s} mk_s={mk_s})")
        else:
            # DARK mode — no saturated colour in the box (black rubber / grey).
            v_med = int(np.median(V[inner])); s_med = int(np.median(S[inner]))
            self._lock_mode = "dark"
            self._dark_v = int(max(60, min(v_med + 35, 110)))
            self._dark_s = int(max(80, min(s_med + 60, 140)))
            sw = (0, 0, max(v_med, 30))
            log.info(f"Lock: DARK mode v<{self._dark_v} s<{self._dark_s}")
        bgr = cv2.cvtColor(np.uint8([[list(sw)]]), cv2.COLOR_HSV2BGR)[0, 0]
        if hasattr(self, "_color_swatch_lbl"):
            self._color_swatch_lbl.configure(
                fg_color="#%02x%02x%02x" % (int(bgr[2]), int(bgr[1]), int(bgr[0])))
        return True

    def _color_mask(self, hsv):
        """Binary mask of the locked target — a hue band (hue mode) or dark +
        low-saturation pixels (dark mode). Hue mode handles wrap-around (e.g. red)."""
        if self._lock_mode == "dark":
            S, V = hsv[..., 1], hsv[..., 2]
            return ((V < self._dark_v) & (S < self._dark_s)).astype(np.uint8) * 255
        h, tol = self._color_hue, self._color_tol
        s, v = self._color_s_min, self._color_v_min
        lo, hi = h - tol, h + tol
        if lo < 0:
            return (cv2.inRange(hsv, (0, s, v), (hi, 255, 255)) |
                    cv2.inRange(hsv, (180 + lo, s, v), (179, 255, 255)))
        if hi > 179:
            return (cv2.inRange(hsv, (lo, s, v), (179, 255, 255)) |
                    cv2.inRange(hsv, (0, s, v), (hi - 180, 255, 255)))
        return cv2.inRange(hsv, (lo, s, v), (hi, 255, 255))

    def _run_color_lock(self, frame):
        """Track the target by its locked colour. Sets head_y / last_bbox from the
        best colour blob near the last position; holds position if the colour is hidden."""
        if self._color_needs_derive or self._color_hue is None:
            if not self._derive_color_lock(frame):
                self.tracking_ok = False
                self._track_msg = "COLOUR LOCK - draw a box on the marker"
                return
            self._color_needs_derive = False
        fh, fw = frame.shape[:2]
        px, py, pw, ph = self.last_bbox
        lcx, lcy = px + pw // 2, py + ph // 2
        # Search only a window around the last position — never scan the whole scene,
        # so it can't jump to distant skin/fabric of a similar colour.
        R = int(max(np.hypot(pw, ph), 40.0) * self._COLOR_SEARCH)
        sx0, sy0 = max(0, lcx - R), max(0, lcy - R)
        sx1, sy1 = min(fw, lcx + R), min(fh, lcy + R)
        roi = frame[sy0:sy1, sx0:sx1]
        if roi.shape[0] < 4 or roi.shape[1] < 4:
            self.tracking_ok = False
            self._track_msg  = "COLOUR HIDDEN - holding"
            return
        scale, small = 1.0, roi
        if self._proc_width and roi.shape[1] > self._proc_width:
            scale = self._proc_width / float(roi.shape[1])
            small = cv2.resize(roi, (self._proc_width, max(1, int(roi.shape[0] * scale))),
                               interpolation=cv2.INTER_AREA)
        mask = self._color_mask(cv2.cvtColor(small, cv2.COLOR_BGR2HSV))
        ker  = np.ones((3, 3), np.uint8)
        mask = cv2.morphologyEx(mask, cv2.MORPH_OPEN,  ker)
        mask = cv2.morphologyEx(mask, cv2.MORPH_CLOSE, ker)
        cnts = cv2.findContours(mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)[0]
        inv = 1.0 / scale
        min_area = max(12.0, 0.03 * pw * ph * scale * scale)   # ~3% of the last box
        max_area = self._COLOR_MAX_AREA * pw * ph * scale * scale
        max_jump2 = (max(np.hypot(pw, ph), 40.0) * self._COLOR_MAX_JUMP) ** 2
        best = None; best_d2 = None
        for c in cnts:
            a = cv2.contourArea(c)
            if a < min_area or a > max_area:       # speckle, or a scene-sized skin blob
                continue
            bx, by, bw, bh = cv2.boundingRect(c)
            fx = sx0 + int(bx * inv); fy = sy0 + int(by * inv)
            fw2, fh2 = int(bw * inv), int(bh * inv)
            ccx, ccy = fx + fw2 // 2, fy + fh2 // 2
            if any(zx <= ccx <= zx + zw and zy <= ccy <= zy + zh
                   for zx, zy, zw, zh in self._exclusion_zones):
                continue
            d2 = (ccx - lcx) ** 2 + (ccy - lcy) ** 2
            if best is None or d2 < best_d2:                    # nearest-to-last
                best = (fx, fy, fw2, fh2, ccy); best_d2 = d2
        # Accept the nearest blob only if it's a plausible move — otherwise the target
        # is hidden and we hold, rather than jump to a same-colour thing elsewhere.
        if best is None or best_d2 > max_jump2:
            self.tracking_ok = False
            self._track_msg  = "COLOUR HIDDEN - holding"
            return
        fx, fy, fw2, fh2, ccy = best
        self.last_bbox   = (fx, fy, fw2, fh2)
        self.tracking_ok = True
        self._track_msg  = ""
        self.head_y      = ccy
        self._head_y_history.append(ccy)
        if self._auto_cum_enabled and not self._cum_stopped:
            self._tick_cum_detect(ccy, frame)

    def _on_active_edging_toggle(self):
        """Toggle the Active Edging auto-ramp cycle."""
        self._active_edging = bool(self._active_edging_var.get())
        self._ae_below_since = None   # reset the hold timer either way
        self._save_config()
        if self._active_edging:
            _lvl  = self.aggr_var.get()
            _hold = self._ae_hold_secs.get(_lvl, 10)
            self._snark_label.configure(
                text=tr("🔁 Active Edging on — ramps up after {}s off the edge ({})").format(_hold, _lvl),
                text_color="#F5A623")
            log.info(f"Active Edging ON — {_lvl}: hold {_hold}s, "
                     f"ramp {self._ae_ramp_pct.get(_lvl, 8.0)}%/s")
        else:
            self._snark_label.configure(text=tr("🔁 Active Edging off"), text_color="#F5A623")
            log.info("Active Edging OFF")

    def _maybe_yolo_reanchor(self, frame):
        # Lock mode: user has a stable target (e.g. a metal plug/electrode) and
        # wants it left alone — skip all YOLO reanchoring so it can't be stolen.
        if self._reanchor_lock:
            return
        interval = self._YOLO_INTERVAL_LOST if not self.tracking_ok else self._YOLO_INTERVAL_LOCKED
        self.yolo_frame_counter += 1
        if not self.detector.available or self.yolo_frame_counter < interval:
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
        # Discard detections whose centre falls inside any user-drawn exclusion zone
        for _ez in self._exclusion_zones:
            _ex, _ey, _ew, _eh = _ez
            if _ex <= det_cx <= _ex + _ew and _ey <= det_cy <= _ey + _eh:
                return
        # Clamp to a minimum so a shrunken tracker bbox can't permanently reject
        # valid YOLO detections that are "too far" from a tiny artifact.
        diag = max(np.sqrt(pw ** 2 + ph ** 2), 100.0)

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
            self._tracker_init(frame, yolo_bbox)
            self.last_bbox      = yolo_bbox
            self.yolo_candidate = None
            x, y, w, h = yolo_bbox
            cv2.rectangle(frame, (x, y), (x + w, y + h), (255, 255, 0), 1)

    def _downscale(self, frame):
        """Return (small_frame, scale) for CSRT/detector work. scale = small/orig
        (<=1.0); full-res coords = small_coords / scale. Downscales only when the
        frame is wider than the configured processing width — never upscales, and
        is a no-op (scale 1.0) when Tracking resolution is 'Full'."""
        pw = self._proc_width
        if not pw:
            return frame, 1.0
        h, w = frame.shape[:2]
        if w <= pw:
            return frame, 1.0
        scale = pw / float(w)
        small = cv2.resize(frame, (pw, max(1, int(round(h * scale)))),
                           interpolation=cv2.INTER_AREA)
        return small, scale

    def _make_tracker(self):
        """Create the tracking backend. TrackerVit (deep, occlusion-robust, ~constant
        cost) by default; falls back to CSRT if the model is missing or Vit won't
        load. Sets _tracker_has_score so _run_tracker knows to read a confidence."""
        if self._tracker_backend == "vit":
            try:
                model = resource_path(os.path.join("models", "vittrack.onnx"))
                if os.path.exists(model):
                    params = cv2.TrackerVit_Params()
                    params.net = model
                    t = cv2.TrackerVit_create(params)
                    self._tracker_has_score = True
                    log.info("Tracker backend: Vit (deep)")
                    return t
                log.warning(f"Vit model not found at {model} — using CSRT")
            except Exception as e:
                log.warning(f"TrackerVit unavailable ({type(e).__name__}: {e}) — using CSRT")
            self._tracker_backend = "csrt"        # reflect the fallback in UI + config
            if hasattr(self, "_tracker_var"):
                self._tracker_var.set("Classic")
        self._tracker_has_score = False
        log.info("Tracker backend: CSRT (classic)")
        return cv2.TrackerCSRT_create()

    def _tracker_init(self, full_frame, full_bbox):
        """(Re)initialise the CSRT tracker at the current processing resolution.
        full_bbox is in full-res frame pixels; it is scaled into the tracker's
        working space and clamped in-bounds so update()s stay consistent. Records
        the scale used so _run_tracker can detect a live change and re-anchor."""
        small, scale = self._downscale(full_frame)
        x, y, w, h = full_bbox
        if scale != 1.0:
            x, y, w, h = x * scale, y * scale, w * scale, h * scale
        sh, sw = small.shape[:2]
        x = max(0, min(int(round(x)), sw - 1))
        y = max(0, min(int(round(y)), sh - 1))
        w = max(1, min(int(round(w)), sw - x))
        h = max(1, min(int(round(h)), sh - y))
        self.tracker.init(small, (x, y, w, h))
        self._track_scale = scale
        self._track_init_wh = (max(1, int(full_bbox[2])), max(1, int(full_bbox[3])))

    def _on_proc_res_change(self, _value=None):
        """Tracking-resolution preset changed — apply live (the tracker re-anchors
        itself on the next frame via _run_tracker) and persist."""
        self._proc_width = self._PROC_PRESETS.get(self.proc_res_var.get(), 0)
        self._save_config()
        log.info(f"Tracking resolution -> {self.proc_res_var.get()} "
                 f"(proc_width={self._proc_width or 'native'})")
        self._frame_dt.clear(); self._last_perf_log = 0.0  # fresh perf window

    def _on_tracker_backend_change(self, _value=None):
        """Deep/Classic tracker toggle — rebuild the tracker on the next frame (it
        re-anchors on the current target) and persist."""
        self._tracker_backend = "vit" if self._tracker_var.get() == "Deep" else "csrt"
        self._tracker_dirty = True
        self._save_config()
        log.info(f"Tracking engine -> {self._tracker_var.get()} ({self._tracker_backend})")
        self._frame_dt.clear(); self._last_perf_log = 0.0  # fresh perf window

    def _refresh_lock_template(self, frame):
        """While locked and tracking cleanly, keep a fresh appearance patch of the
        target so _relock_from_template can re-find it if CSRT drops the lock."""
        now = time.time()
        if now - self._lock_tpl_time < self._LOCK_TPL_REFRESH_S:
            return
        x, y, w, h = self.last_bbox
        fh, fw = frame.shape[:2]
        x0, y0 = max(0, int(x)), max(0, int(y))
        x1, y1 = min(fw, int(x + w)), min(fh, int(y + h))
        if x1 - x0 >= 8 and y1 - y0 >= 8:
            self._lock_template = frame[y0:y1, x0:x1].copy()
            self._lock_tpl_time = now

    def _relock_from_template(self, frame):
        """Search a window around the last position for the locked target's saved
        appearance patch; on a confident match, re-init the tracker there. Returns
        True on re-lock. It only accepts the saved patch, so it re-finds the SAME
        target (a plug/clip/marker) and won't wander onto a hand or bare skin."""
        tpl = self._lock_template
        if tpl is None:
            return False
        th, tw = tpl.shape[:2]
        fh, fw = frame.shape[:2]
        if th < 8 or tw < 8 or th >= fh or tw >= fw:
            return False
        px, py, pw, ph = self.last_bbox
        cx, cy = px + pw // 2, py + ph // 2
        sw = min(fw, max(tw + 2, int(tw * self._RELOCK_SEARCH)))
        sh = min(fh, max(th + 2, int(th * self._RELOCK_SEARCH)))
        sx0 = max(0, min(cx - sw // 2, fw - sw))
        sy0 = max(0, min(cy - sh // 2, fh - sh))
        region = frame[sy0:sy0 + sh, sx0:sx0 + sw]
        if region.shape[0] < th or region.shape[1] < tw:
            return False
        res = cv2.matchTemplate(region, tpl, cv2.TM_CCOEFF_NORMED)
        _, max_val, _, max_loc = cv2.minMaxLoc(res)
        if not (max_val >= self._RELOCK_MIN_SCORE):   # also rejects NaN (flat patch)
            return False
        mx, my = sx0 + max_loc[0], sy0 + max_loc[1]
        bbox = (int(mx), int(my), int(tw), int(th))
        self._tracker_init(frame, bbox)
        self.last_bbox   = bbox
        self.tracking_ok = True
        self._track_msg  = ""
        self.head_y      = my + th // 2
        self._head_y_history.append(self.head_y)
        log.debug(f"Lock re-find: matched {max_val:.2f} at ({mx},{my})")
        return True

    def _run_tracker(self, frame):
        """Update tracker state — does NOT draw; call _draw_tracking_overlay every frame."""
        small, scale = self._downscale(frame)
        if self._tracker_dirty:
            # Backend changed (Deep/Classic toggle or config load) — rebuild + re-anchor.
            self.tracker = self._make_tracker()
            self._tracker_dirty = False
            self._tracker_init(frame, self.last_bbox)
            small, scale = self._downscale(frame)
        elif scale != self._track_scale:
            # Processing-resolution changed (setting toggled live, or the capture
            # frame size shifted). The tracker was init on a different-sized image —
            # re-anchor on the current frame before updating, else it loses lock.
            self._tracker_init(frame, self.last_bbox)
        success, new_bbox = self.tracker.update(small)
        good = False
        if success:
            inv = 1.0 / scale
            x, y, w, h_box = [int(v * inv) for v in new_bbox]
            # Clamp against ballooning — some trackers (Vit) grow the box on smooth or
            # ambiguous targets; keep the centre but cap the size at ~2.5x what it was
            # locked at, so it can't swallow half the screen.
            iw, ih = self._track_init_wh
            if w > 2.5 * iw or h_box > 2.5 * ih:
                cx0, cy0 = x + w // 2, y + h_box // 2
                w, h_box = min(w, int(2.5 * iw)), min(h_box, int(2.5 * ih))
                x, y = cx0 - w // 2, cy0 - h_box // 2
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
                good = True
                self.last_bbox    = (x, y, w, h_box)
                self.tracking_ok  = True
                self._track_msg   = ""
                self.head_y       = new_cy
                self._head_y_history.append(new_cy)
                # Feed auto-cum detector
                if self._auto_cum_enabled and not self._cum_stopped:
                    self._tick_cum_detect(new_cy, frame)
                # Keep the locked target's appearance patch current for re-find.
                if self._reanchor_lock:
                    self._refresh_lock_template(frame)
            else:
                reason = "size" if not size_ok else "jump"
                self._track_msg  = f"TRACKING SUSPECT ({reason}) - Frozen"
        else:
            self._track_msg  = tr("TRACKING LOST - Frozen at last position")

        if not good:
            self.tracking_ok = False
            # Lock mode: re-find the SAME target by appearance near where it was,
            # rather than freezing. YOLO reanchor is off while locked (and can't see
            # a plug/clip anyway), so this is the only recovery path.
            if self._reanchor_lock and self._relock_from_template(frame):
                return
            self._tracker_init(frame, self.last_bbox)

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

        # Exclusion zones — semi-transparent dark fill + red border + X
        # BGR: (0,0,255) → RGB (255,0,0) = red
        _EZ_RED = (0, 0, 255)
        if self._exclusion_zones:
            _ez_ol = frame.copy()
            for _ez in self._exclusion_zones:
                _ex, _ey, _ew, _eh = _ez
                cv2.rectangle(_ez_ol, (_ex, _ey), (_ex + _ew, _ey + _eh), (10, 10, 10), -1)
            cv2.addWeighted(_ez_ol, 0.7, frame, 0.3, 0, frame)
            for _ez in self._exclusion_zones:
                _ex, _ey, _ew, _eh = _ez
                cv2.rectangle(frame, (_ex, _ey), (_ex + _ew, _ey + _eh), _EZ_RED, 1)
                cv2.line(frame, (_ex, _ey), (_ex + _ew, _ey + _eh), _EZ_RED, 1)
                cv2.line(frame, (_ex + _ew, _ey), (_ex, _ey + _eh), _EZ_RED, 1)

        # ── Grid navigator overlay ────────────────────────────────────────────
        if getattr(self, '_grid_active', False):
            fh, fw = frame.shape[:2]
            region = self._grid_region or (0, 0, fw, fh)
            gx, gy, gw, gh = region
            # dim everything outside the active region
            dim = frame.copy()
            cv2.rectangle(dim, (0, 0), (fw, fh), (0, 0, 0), -1)
            cv2.addWeighted(dim, 0.45, frame, 0.55, 0, frame)
            # re-brighten the active region
            cv2.rectangle(frame, (gx, gy), (gx + gw, gy + gh), (255, 255, 255), 0)
            cw, ch = gw // 3, gh // 3
            # grid lines
            for col in range(1, 3):
                lx = gx + col * cw
                cv2.line(frame, (lx, gy), (lx, gy + gh), (200, 200, 200), 1)
            for row in range(1, 3):
                ly = gy + row * ch
                cv2.line(frame, (gx, ly), (gx + gw, ly), (200, 200, 200), 1)
            # border
            cv2.rectangle(frame, (gx, gy), (gx + gw, gy + gh), (255, 255, 255), 2)
            # numbers
            for i in range(9):
                row, col = divmod(i, 3)
                ncx = gx + col * cw + cw // 2
                ncy = gy + row * ch + ch // 2
                cv2.putText(frame, str(i + 1), (ncx - 9, ncy + 9),
                            cv2.FONT_HERSHEY_SIMPLEX, 0.65, (255, 255, 255), 2)
            # hint bar
            cv2.putText(frame, "say: 1-9 zoom  |  select/here confirm  |  cancel reset",
                        (5, fh - 8), cv2.FONT_HERSHEY_SIMPLEX, 0.38, (200, 180, 255), 1)

        # In-progress exclusion zone while user is dragging
        if getattr(self, '_ez_drawing', False):
            _s = getattr(self, '_ez_disp_start', None)
            _en = getattr(self, '_ez_disp_end',   None)
            if _s and _en:
                _sc = getattr(self, '_disp_scale', 1.0) or 1.0
                _ox = getattr(self, '_disp_offset_x', 0)
                _oy = getattr(self, '_disp_offset_y', 0)
                _fw = getattr(self, '_disp_frame_w', frame.shape[1])
                _fh = getattr(self, '_disp_frame_h', frame.shape[0])
                def _cl(v, lo, hi): return max(lo, min(hi, v))
                _px0 = _cl(int((min(_s[0], _en[0]) - _ox) / _sc), 0, _fw)
                _py0 = _cl(int((min(_s[1], _en[1]) - _oy) / _sc), 0, _fh)
                _px1 = _cl(int((max(_s[0], _en[0]) - _ox) / _sc), 0, _fw)
                _py1 = _cl(int((max(_s[1], _en[1]) - _oy) / _sc), 0, _fh)
                _ip_ol = frame.copy()
                cv2.rectangle(_ip_ol, (_px0, _py0), (_px1, _py1), (10, 10, 10), -1)
                cv2.addWeighted(_ip_ol, 0.7, frame, 0.3, 0, frame)
                cv2.rectangle(frame, (_px0, _py0), (_px1, _py1), _EZ_RED, 1)

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
                self.session_logger.log_state_change("Edging")
            elif (now - getattr(self, '_edge_enter_time', now) >= 1.0
                  and not getattr(self, '_edge_counted', False)
                  and now - getattr(self, '_last_edge_time', 0.0) >= 10.0):
                self.edge_count += 1
                self._last_edge_time = now
                self._edge_counted = True
                self.session_logger.log_edge_counted(self.edge_count)
                if self._hf_enabled:
                    self._hf_check_edge()
        else:
            if self._prev_state == "Edging":
                self.session_logger.log_state_change(state)
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
            text=(tr("Session: {}:{}  |  Edges: {}  |  Edging {}%  Erect {}%  Flaccid {}%").format(
                  f"{m:02d}", f"{s:02d}", self.edge_count,
                  f"{pcts['Edging']:.0f}", f"{pcts['Erect']:.0f}", f"{pcts['Flaccid']:.0f}"))
        )
        if self._play_mode:
            elapsed = time.time() - self.session_start
            m2, s2 = divmod(int(elapsed), 60)
            self._play_edges_lbl.configure(text=tr("{} edges").format(self.edge_count))
            self._play_time_lbl.configure(text=f"{m2:02d}:{s2:02d}")

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
        # Point of No Return — yellow (BGR 0,215,255). Cross ABOVE it → instant cut.
        ponr_y = self.heights.get("PONR")
        if ponr_y is not None:
            py = int(ponr_y)
            cv2.line(frame, (0, py), (fw, py), (0, 215, 255), 2)
            label_y = py + 16 if py < 16 else py - 5
            cv2.putText(frame, "POINT OF NO RETURN", (4, label_y),
                        cv2.FONT_HERSHEY_SIMPLEX, 0.45, (0, 215, 255), 1, cv2.LINE_AA)

    def _compute_volume_delta(self):
        if any(self.heights.get(k) is None for k in ("Edging", "Erect", "Flaccid")):
            return 0.0
        # SAFETY: when tracking is lost/frozen the head_y is stale — never climb on
        # dead data. Both passive recovery and Active Edging would otherwise ramp to
        # the ceiling on a frozen position (occlusion, lock-on drift) with the user
        # potentially restrained. Hold volume steady and reset the AE hold timer.
        if not self.tracking_ok:
            self._ae_below_since = None
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

        # ── Active Edging timer: track how long we've held off the edge. ──────
        # position 0 = edge line, erect_norm = erect line. "near_edge" = in the
        # half of the sweet zone closest to the edge (or past it). Reset the hold
        # timer whenever we're near the edge; start it the moment we back off.
        if self._active_edging:
            if position < erect_norm * 0.5:
                self._ae_below_since = None
            elif self._ae_below_since is None:
                self._ae_below_since = time.time()

        if 0.0 <= position <= erect_norm:
            # Active Edging: once we've held off the edge long enough, ramp
            # aggressively back toward max until the next edge resets the cycle.
            # Hold + ramp scale per difficulty (relentless: harder = sooner/faster).
            if (self._active_edging and self._ae_below_since is not None
                    and time.time() - self._ae_below_since
                        >= self._ae_hold_secs.get(aggr_level, 10)):
                ramp_pct = self._ae_ramp_pct.get(aggr_level, 8.0)
                return (ramp_pct / 100.0) * VOLUME_UPDATE_INTERVAL

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
                # Passive recovery: a flat climb so holding steady in the erect
                # zone recovers instead of parking. Scaled by distance from the
                # edge (0 at edge → full at erect line) so it never fights an
                # active edge, and per-difficulty (relentless: harder = faster).
                if self._erect_recovery:
                    rec_per_tick = (self._recovery_pct.get(aggr_level, 1.6) / 100.0) \
                                   * VOLUME_UPDATE_INTERVAL
                    recovery = rec_per_tick * (position / max(erect_norm, 1e-3))
                else:
                    recovery = 0.0
                return VOLUME_STEP * vel_nudge * aggr_mult + recovery
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


    def _check_ponr(self):
        """Point of No Return — the instant the head crosses ABOVE the optional
        yellow line, hard-cut every output to the floor.

        Edge-triggered: one instant cut fires on the crossing, then _ponr_triggered
        latches so _tick_volume pins the floor until the head drops back below.
        Uses the RAW head_y (not the smoothed history the reactive ramp uses) so
        it responds to a last-second spike immediately instead of lagging.

        Optional by design: an unset (None) line makes this a no-op, and it never
        touches the state machine or volume curve — nothing else depends on it."""
        ponr_y = self.heights.get("PONR")
        if ponr_y is None:
            self._ponr_triggered = False
            return
        # Never trip on a lost/frozen head_y (stale data would latch the cut, or
        # fire it while the tracker is drifting). Same safety gate the recovery
        # ramp uses.
        if not self.tracking_ok:
            return
        # Deliberate flows that own the outputs get to keep them.
        if self._cum_stopped or self._cum_allowed or self._ruin_active:
            return
        above = self.head_y < ponr_y   # smaller Y = higher on screen = past the line
        if above and not self._ponr_triggered:
            self._ponr_triggered = True
            floor_val = min(self.min_vol_var.get(), self.max_vol_var.get()) / 100.0
            self._set_all_outputs(floor_val, instant=True)
            log.info(f"Point of No Return crossed (head_y={self.head_y} < {int(ponr_y)}) "
                     f"— output cut to floor {floor_val*100:.0f}%")
            try:
                self._snark_label.configure(
                    text=tr("⚡ Point of No Return — cut"), text_color=self._C_YELLOW)
            except Exception:
                pass
        elif not above:
            self._ponr_triggered = False

    def _tick_volume(self):
        cur_time = time.time()
        if cur_time - self.last_vol_time < VOLUME_UPDATE_INTERVAL:
            return
        self.last_vol_time = cur_time

        # ── Ruin in progress — the pulse sequence owns the output entirely ───
        # Don't let position-driven deltas (esp. the new recovery) climb back out
        # of the ruin's deliberate 0% gaps between pulses.
        if self._ruin_active:
            return

        # ── Cum allowed — volume locked at 100% until "I've CUM" ─────────────
        if self._cum_allowed:
            return

        # ── Session stopped after "I've CUM" — volume pinned at 0 ───────────
        # (_on_cum set the volume to 0 already; just keep it there every tick
        # in case something else nudged it.)
        if self._cum_stopped:
            self._set_all_outputs(0.0)
            return

        # ── Point of No Return active — head is past the yellow line ─────────
        # _check_ponr already fired the instant cut on the crossing; keep the
        # floor pinned every tick so position-driven deltas (and recovery/AE)
        # can't climb back out until the head drops below the line. Overrides
        # hold, since a frozen-high volume is exactly what this line prevents.
        if self._ponr_triggered and self.heights.get("PONR") is not None:
            floor_val = min(self.min_vol_var.get(), self.max_vol_var.get()) / 100.0
            self._set_all_outputs(floor_val)
            return

        if self.hold_active:
            return  # Volume frozen by user

        history    = self._head_y_history
        smoothed_y = sum(history) / len(history) if history else self.head_y


        floor_val = min(self.min_vol_var.get(), self.max_vol_var.get()) / 100.0
        ceil_val  = self.max_vol_var.get() / 100.0
        delta     = self._compute_volume_delta()

        # ── HR modifier: tighten denial / slow rewards when heart rate is high ─
        if self.hr_on.get() and self._active_hr.connected:
            hr_mod = self._active_hr.modifier()   # 1.0 (resting) → 2.0 (peak)
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
            else:
                self.xtoys.keepalive(floor=floor_val, ceiling=ceil_val)
            self.xtoys.maybe_reconnect()
        if self.audio_on.get() and self.win_audio and self.win_audio.connected and delta != 0.0:
            self.win_audio.adjust_volume(delta, floor=floor_val, ceiling=ceil_val)
        if self.mp3_on.get() and self.music_player and delta != 0.0:
            self.music_player.adjust_volume(delta, floor=floor_val, ceiling=ceil_val)

    def _update_status_label(self, state):
        quality_str = tr("Track: OK") if self.tracking_ok else tr("Track: LOST")
        parts = []
        conn_color = "#ffaa00"
        help_msg = ""   # contextual fix-it line under the status; blank when healthy

        if self.restim_on.get():
            ok = bool(self.restim.ws)
            if ok: conn_color = "#00ff00"
            elif conn_color != "#00ff00": conn_color = "#ff0000"
            parts.append(tr("WS: {}").format('OK' if ok else tr('Disconnected')))
            if not ok and not help_msg:
                help_msg = tr("Restim disconnected — enable its WebSocket server and check the port matches.")
        if self.xtoys_on.get():
            if not self.xtoys.enabled:
                if conn_color != "#00ff00": conn_color = "#ffaa00"
                parts.append(tr("xToys: No ID"))
                if not help_msg:
                    help_msg = tr("xToys: paste your Webhook ID from xtoys.app/me → Private Webhook (or tap Get ID).")
            elif self.xtoys.listening:
                # Script acked over the socket → genuinely running and receiving us.
                conn_color = "#00ff00"
                toys = self.xtoys.toys
                if toys:
                    _t = toys[0] if len(toys) == 1 else f"{toys[0]} +{len(toys) - 1}"
                    parts.append(tr("xToys: Live ({})").format(_t))
                else:
                    parts.append(tr("xToys: Live"))
            elif self.xtoys.connected:
                # Cloud accepts our POSTs but the script hasn't acked — most likely the
                # block was stopped (lookalikes: an old copy without the ack, or a wrong
                # ID). Hence the honest "block STOPPED?" with the "?" + a fuller hint.
                if conn_color != "#00ff00": conn_color = "#ffaa00"
                parts.append(tr("xToys: block STOPPED?"))
                if not help_msg:
                    help_msg = tr("xToys reached the cloud but the block didn't answer — press ▶ to run the "
                                "VisualStimEdger block in your xToys tab (reload it if it's an old copy, or "
                                "re-check your Webhook ID).")
            else:
                if conn_color != "#00ff00": conn_color = "#ff0000"
                parts.append(tr("xToys: Offline"))
                if not help_msg:
                    help_msg = tr("Can't reach xToys — check your internet connection and Webhook ID.")
        if self.audio_on.get():
            if self.win_audio and self.win_audio.connected:
                conn_color = "#00ff00"
                parts.append(tr("Audio: OK"))
            else:
                parts.append(tr("Audio: No Device"))
        if self.mp3_on.get() and self.music_player:
            self._mp3_update_track_label()
            if self.music_player._state == "playing": conn_color = "#00ff00"
            parts.append(tr("MP3: {}").format(self.music_player.track_name or '—'))
        if self.hr_on.get():
            sbpm = self._active_hr.smooth_bpm()
            if self._active_hr.connected and sbpm is not None:
                parts.append(tr("\u2665 {} bpm").format(f"{sbpm:.0f}"))
                if hasattr(self, '_hr_bpm_label'):
                    self._hr_bpm_label.configure(
                        text=tr("\u2665 {} bpm").format(f"{sbpm:.0f}"),
                        text_color="#e91e63")
            else:
                parts.append(tr("\u2665 No HR"))
                if hasattr(self, '_hr_bpm_label'):
                    self._hr_bpm_label.configure(
                        text=tr("\u2665 -- bpm"),
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
        if self._color_lock:
            yolo_str = "Track: " + ("dark" if self._lock_mode == "dark" else "colour")
        elif not self.detector.available:
            yolo_str = "YOLO: n/a"            # model failed to load
        elif self._reanchor_lock:
            yolo_str = "YOLO: off"            # "Lock on target" is on — reanchoring disabled
        elif self.detector.last_conf > 0:
            yolo_str = f"YOLO: {self.detector.last_conf:.0%}"
        else:
            yolo_str = "YOLO: --"             # enabled but hasn't detected the head yet

        status_text = (tr("State: {}  |  Vol: {}  |  {}  |  {}  |  {} fps  |  {}").format(
                       state, vol_str, quality_str, src_str, f"{fps:.0f}", yolo_str))
        if self._evil_mode:
            status_text += "  |  😈 EVIL"
            conn_color = "#cc0000"
        self.info_label.configure(
            text=status_text,
            text_color=conn_color,
        )
        if hasattr(self, '_help_label'):
            self._help_label.configure(text=help_msg)
        self._maybe_tribute()      # #1: milestone tribute nudge (self-gated, ~once/session)

        # UX-4: calibration hint — show when all three heights are unset and session is live
        if hasattr(self, '_snark_label'):
            _all_unset = all(self.heights.get(k) is None for k in ("Erect", "Flaccid", "Edging"))
            _session_live = self._running and not self._cum_stopped and not getattr(self, 'tracking_paused', False)
            _snark_now = self._snark_label.cget("text")
            _calib_hint = tr("💡 Press AUTO to calibrate height lines")
            if _all_unset and _session_live:
                if _snark_now == "" or _snark_now == _calib_hint:
                    self._snark_label.configure(text=_calib_hint, text_color="#5ba3c9")
            elif _snark_now == _calib_hint:
                self._snark_label.configure(text="")

        # Update play panel if visible
        if self._play_mode:
            _sc = {"Edging": self._C_ACCENT, "Erect": self._C_GREEN, "Flaccid": self._C_BLUE}
            self._play_state_lbl.configure(
                text=state, text_color=_sc.get(state, self._C_TEXT))
            self._play_vol_lbl.configure(text=tr("Vol: {}").format(vol_str))
            if self.hr_on.get():
                sbpm2 = self._active_hr.smooth_bpm()
                self._play_hr_lbl.configure(
                    text=tr("♥ {} bpm").format(f"{sbpm2:.0f}") if sbpm2 else "♥ --")
            else:
                self._play_hr_lbl.configure(text="")
            self._play_snark_lbl.configure(
                text=self._snark_label.cget("text") if hasattr(self, '_snark_label') else "")

    def _maybe_tribute(self):
        """#1 — once per session, after a genuinely real session (>25 min AND >=5
        edges), nudge for a tribute. Gated: only after 2 prior qualifying sessions,
        and at most once a week. A show (or the user's dismiss) buys a week of quiet."""
        if getattr(self, '_tribute_handled_session', False) or not self._running:
            return
        if (time.time() - self.session_start) < 25 * 60 or self.edge_count < 5:
            return
        self._tribute_handled_session = True   # handled once per session (shown or not)
        self._tribute_session_count = getattr(self, '_tribute_session_count', 0) + 1
        now = time.time()
        due = (now - getattr(self, '_tribute_last_shown', 0.0)) > 7 * 24 * 3600
        if self._tribute_session_count > 2 and due:
            self._tribute_last_shown = now
            self._show_tribute_toast()
        self._save_config()

    def _show_tribute_toast(self):
        """Small non-modal, non-focus-stealing toast. Never blocks the session."""
        try:
            t = tk.Toplevel(self.root)
            t.overrideredirect(True)
            t.attributes("-topmost", True)
            self.root.update_idletasks()
            x = self.root.winfo_x() + self.root.winfo_width() // 2 - 200
            y = self.root.winfo_y() + 70
            t.geometry(f"400x120+{max(0, x)}+{max(0, y)}")
            fr = ctk.CTkFrame(t, fg_color="#141414", border_width=2,
                              border_color="#F5A623", corner_radius=10)
            fr.pack(fill=tk.BOTH, expand=True)
            ctk.CTkLabel(fr, text=tr("You've been edging {}+ times, hands-free.").format(self.edge_count),
                         font=ctk.CTkFont(size=13, weight="bold"),
                         text_color=self._C_TEXT).pack(pady=(12, 2), padx=14)
            ctk.CTkLabel(fr, text=tr("Free tool. One guy made it. Tribute your dev? 😈"),
                         font=ctk.CTkFont(size=11), text_color=self._C_TEXT_DIM).pack(pady=(0, 8), padx=14)
            br = ctk.CTkFrame(fr, fg_color="transparent"); br.pack(pady=(0, 10))
            ctk.CTkButton(br, text=tr("☕ Tribute"), width=120, fg_color="#F5A623",
                          hover_color="#d68f14", text_color="#000000",
                          font=ctk.CTkFont(weight="bold"),
                          command=lambda: (webbrowser.open("https://ko-fi.com/stimstation"),
                                           t.destroy())).pack(side=tk.LEFT, padx=8)
            ctk.CTkButton(br, text=tr("Not now"), width=90, fg_color=self._C_SURFACE,
                          hover_color="#4a4a4a", border_width=1, border_color=self._C_BORDER,
                          command=t.destroy).pack(side=tk.LEFT, padx=8)
            t.after(25000, lambda: t.winfo_exists() and t.destroy())
        except Exception as e:
            log.debug(f"tribute toast failed: {type(e).__name__}")

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
        result = self._last_letmecum_result
        if result:
            elapsed_lmc = time.time() - self._last_letmecum_time
            if result == "granted" and self._cum_allowed:
                grant_left = self._cum_grant_expires - time.time()
                letmecum = {"result": "granted", "time_left": round(max(grant_left, 0))}
            elif result == "expired" and elapsed_lmc < 5:
                letmecum = {"result": "expired"}
            elif result == "ruined" and elapsed_lmc < 8:
                letmecum = {"result": "ruined"}
            elif result == "denied":
                cooldown_left = self._letmecum_cooldown_until - time.time()
                if cooldown_left > 0:
                    letmecum = {"result": "denied", "retry_in": round(cooldown_left)}
                elif elapsed_lmc < 5:
                    letmecum = {"result": "denied", "retry_in": 0}

        # Heart rate data for overlay
        hr_data = None
        if self.hr_on.get():
            sbpm = self._active_hr.smooth_bpm()
            hr_data = {
                "bpm": round(sbpm) if sbpm is not None else None,
                "connected": self._active_hr.connected,
                "modifier": round(self._active_hr.modifier(), 2),
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

    # ================================================================ bondage mode

    def _open_bondage_splash(self):
        """Launch the bondage mode setup flow (two-page splash)."""
        if self._bondage_active:
            return  # already live — ignore
        if self._bondage_configured:
            # Setup was already completed this session — skip the splash and go straight in
            self._start_bondage_session()
            return
        model_path = VoiceEngine.model_path_default()
        if not VoiceEngine.model_available(model_path):
            messagebox.showwarning(
                "Vosk model missing",
                f"Bondage Mode requires the Vosk small English model.\n\n"
                f"Download from:\n  https://alphacephei.com/vosk/models\n"
                f"(vosk-model-small-en-us-0.22.zip)\n\n"
                f"Extract so this folder exists:\n  {model_path}",
                parent=self.root,
            )
            return

        win = ctk.CTkToplevel(self.root)
        win.title(tr("🎙 Bondage Mode — Setup"))
        win.configure(fg_color=self._C_BG)
        win.transient(self.root)
        win.grab_set()
        win.resizable(False, False)
        win.geometry("480x540")
        _lbl = ctk.CTkFont(size=11, weight="bold")
        _dim = ctk.CTkFont(size=10)
        P = 16

        # ── title ─────────────────────────────────────────────────────────────
        ctk.CTkLabel(win, text=tr("🎙 Bondage Mode Setup"),
                     font=ctk.CTkFont(size=18, weight="bold"),
                     text_color=self._C_TEXT).pack(pady=(P, 4))
        ctk.CTkLabel(win,
                     text=tr("Configure your microphone and safeword before the session starts."),
                     font=_dim, text_color=self._C_TEXT_DIM, wraplength=440).pack(pady=(0, 8))

        card = ctk.CTkFrame(win, fg_color=self._C_SURFACE, corner_radius=8)
        card.pack(fill=tk.X, padx=P, pady=4)

        # ── mic device ────────────────────────────────────────────────────────
        mic_row = ctk.CTkFrame(card, fg_color="transparent")
        mic_row.pack(fill=tk.X, padx=12, pady=(10, 4))
        ctk.CTkLabel(mic_row, text=tr("Microphone:"), font=_lbl,
                     text_color=self._C_TEXT, width=110, anchor="w").pack(side=tk.LEFT)
        mic_devices = VoiceEngine.list_input_devices()
        mic_options  = ["System Default"] + mic_devices
        _default_sel = self._bondage_mic_device if self._bondage_mic_device in mic_devices else "System Default"
        mic_var = tk.StringVar(value=_default_sel)
        ctk.CTkComboBox(mic_row, values=mic_options, variable=mic_var,
                        width=280,
                        fg_color=self._C_SURFACE2, border_color=self._C_BORDER,
                        button_color=self._C_ACCENT, button_hover_color=self._C_ACCENT_H,
                        dropdown_fg_color=self._C_SURFACE2, text_color=self._C_TEXT,
                        ).pack(side=tk.LEFT)

        # ── level meter ───────────────────────────────────────────────────────
        meter_row = ctk.CTkFrame(card, fg_color="transparent")
        meter_row.pack(fill=tk.X, padx=12, pady=(4, 8))
        ctk.CTkLabel(meter_row, text=tr("Mic level:"), font=_dim,
                     text_color=self._C_TEXT_DIM, width=110, anchor="w").pack(side=tk.LEFT)
        meter_cv = tk.Canvas(meter_row, width=280, height=14,
                             bg=self._C_SURFACE2, highlightthickness=0)
        meter_cv.pack(side=tk.LEFT)

        def _draw_meter():
            meter_cv.delete("all")
            engine = getattr(self, '_splash_voice_engine', None)
            lvl = engine.level if engine else 0.0
            fill_w = int(lvl * 280)
            col = "#3EC941" if lvl < 0.6 else "#F5A623" if lvl < 0.85 else "#FF4444"
            if fill_w > 0:
                meter_cv.create_rectangle(0, 0, fill_w, 14, fill=col, outline="")

        ctk.CTkFrame(card, height=1, fg_color=self._C_BORDER).pack(fill=tk.X, padx=12, pady=2)

        # ── safeword ──────────────────────────────────────────────────────────
        sw_row = ctk.CTkFrame(card, fg_color="transparent")
        sw_row.pack(fill=tk.X, padx=12, pady=(8, 4))
        ctk.CTkLabel(sw_row, text=tr("Safeword:"), font=_lbl,
                     text_color=self._C_TEXT, width=110, anchor="w").pack(side=tk.LEFT)
        sw_var = tk.StringVar(value=self._bondage_safeword)
        sw_entry = ctk.CTkEntry(sw_row, textvariable=sw_var, width=160,
                                fg_color=self._C_SURFACE2, border_color=self._C_BORDER,
                                text_color=self._C_TEXT)
        sw_entry.pack(side=tk.LEFT, padx=(0, 8))
        ctk.CTkLabel(sw_row, text=tr("(single word)"), font=_dim,
                     text_color=self._C_TEXT_DIM).pack(side=tk.LEFT)

        # verification status — saved safeword pre-fills the field but still requires saying it
        _pre_verified = False
        verify_lbl = ctk.CTkLabel(
            card,
            text="✓ Using saved safeword" if _pre_verified else "Say your safeword aloud to verify ↑",
            font=_dim,
            text_color="#3EC941" if _pre_verified else self._C_TEXT_DIM,
        )
        verify_lbl.pack(pady=(0, 10))
        _verified = [_pre_verified]

        ctk.CTkFrame(card, height=1, fg_color=self._C_BORDER).pack(fill=tk.X, padx=12, pady=2)

        # ── skip row ──────────────────────────────────────────────────────────
        skip_row = ctk.CTkFrame(card, fg_color="transparent")
        skip_row.pack(fill=tk.X, padx=12, pady=(8, 10))
        skip_var = tk.BooleanVar(value=False)
        skip_cb  = ctk.CTkCheckBox(skip_row, text=tr("Skip voice verification"),
                                   variable=skip_var, onvalue=True, offvalue=False,
                                   font=_dim, text_color=self._C_TEXT_DIM,
                                   fg_color=self._C_ACCENT, hover_color=self._C_ACCENT_H,
                                   border_color=self._C_BORDER)
        skip_cb.pack(side=tk.LEFT)

        # ── navigation ────────────────────────────────────────────────────────
        nav_row = ctk.CTkFrame(win, fg_color="transparent")
        nav_row.pack(fill=tk.X, padx=P, pady=(8, P))

        cont_btn = ctk.CTkButton(nav_row, text=tr("Continue →"),
                                 font=ctk.CTkFont(size=12, weight="bold"),
                                 height=36, corner_radius=4,
                                 fg_color=self._C_SURFACE2, hover_color="#4a4a4a",
                                 text_color=self._C_TEXT_DIM, state="disabled")
        cont_btn.pack(side=tk.RIGHT, padx=(4, 0))

        practice_btn = ctk.CTkButton(nav_row, text=tr("Practice Grid →"),
                                     font=ctk.CTkFont(size=11),
                                     height=36, corner_radius=4,
                                     fg_color="#2a1a3a", hover_color="#4a2a5a",
                                     text_color=self._C_TEXT_DIM, state="disabled")
        practice_btn.pack(side=tk.RIGHT)

        ctk.CTkButton(nav_row, text=tr("Cancel"), command=win.destroy,
                      font=ctk.CTkFont(size=11), height=36, corner_radius=4,
                      fg_color=self._C_SURFACE2, hover_color="#4a4a4a",
                      text_color=self._C_TEXT).pack(side=tk.LEFT)

        # ── engine lifecycle ──────────────────────────────────────────────────
        def _unlock_continue():
            cont_btn.configure(state="normal", fg_color="#4a0a6a", hover_color="#6a1a8a",
                               text_color="white")
            practice_btn.configure(state="normal", fg_color="#2a1a3a", hover_color="#4a2a5a",
                                   text_color=self._C_TEXT)

        if _pre_verified:
            _unlock_continue()

        def _on_splash_keyword(kw):
            """Called from VoiceEngine audio thread."""
            if kw == "safeword":
                self._call_on_main(_on_verified)

        def _on_verified():
            if _verified[0]:
                return
            _verified[0] = True
            verify_lbl.configure(text=tr("✓ Safeword verified!"), text_color="#3EC941")
            _unlock_continue()

        def _on_skip_toggle():
            if skip_var.get():
                _unlock_continue()
            else:
                if not _verified[0]:
                    cont_btn.configure(state="disabled",
                                       fg_color=self._C_SURFACE2, text_color=self._C_TEXT_DIM)

        skip_var.trace_add("write", lambda *_: _on_skip_toggle())

        def _start_splash_engine(*_):
            old = getattr(self, '_splash_voice_engine', None)
            if old:
                old.stop()
            sel = mic_var.get()
            dev = "" if sel == "System Default" else sel
            eng = VoiceEngine(VoiceEngine.model_path_default(), device_name=dev)
            eng.set_safeword(sw_var.get().strip().lower() or "red")
            eng.set_callback(_on_splash_keyword)
            self._splash_voice_engine = eng
            eng.start()

        # Restart engine when mic changes (immediate) or safeword changes (debounced
        # so we don't reload Vosk on every keystroke).
        _sw_restart_job = [None]
        def _debounce_sw(*_):
            if _sw_restart_job[0]:
                win.after_cancel(_sw_restart_job[0])
            _sw_restart_job[0] = win.after(700, _start_splash_engine)

        mic_var.trace_add("write", _start_splash_engine)
        sw_var.trace_add("write", _debounce_sw)
        _start_splash_engine()   # start immediately with current device

        # meter poll
        def _meter_tick():
            if not win.winfo_exists():
                return
            _draw_meter()
            win.after(80, _meter_tick)
        _meter_tick()

        # ── continue action ───────────────────────────────────────────────────
        def _do_continue():
            sel = mic_var.get()
            self._bondage_mic_device = "" if sel == "System Default" else sel
            self._bondage_safeword   = sw_var.get().strip().lower() or "red"
            self._save_config()
            eng = getattr(self, '_splash_voice_engine', None)
            if eng:
                eng.stop()
                self._splash_voice_engine = None
            win.destroy()
            self._start_bondage_session()

        def _do_practice():
            sel = mic_var.get()
            self._bondage_mic_device = "" if sel == "System Default" else sel
            self._bondage_safeword   = sw_var.get().strip().lower() or "red"
            eng = getattr(self, '_splash_voice_engine', None)
            if eng:
                eng.stop()
                self._splash_voice_engine = None
            win.destroy()
            self._open_grid_practice()

        cont_btn.configure(command=_do_continue)
        practice_btn.configure(command=_do_practice)
        win.protocol("WM_DELETE_WINDOW", lambda: (
            setattr(self, '_splash_voice_engine',
                    getattr(self, '_splash_voice_engine', None) and
                    getattr(self, '_splash_voice_engine').stop() or None),
            win.destroy()
        ))

    def _open_grid_practice(self):
        """Splash 2 — 🍆 bouncing mini-game. Navigate the grid to pin the green stem."""
        CV_W, CV_H = 360, 360
        EMOJI_FONT = ("Segoe UI Emoji", 128)
        # Base stem offset for 0° orientation (upper-left of glyph)
        _BASE_CX = -36
        _BASE_CY = -52
        STEM_RX  = 20
        STEM_RY  = 12
        # Only upright and flipped — sideways is too hard to navigate
        _ROT_ANGLES  = [0, 180]
        _ROT_OFFSETS = [
            (_BASE_CX,  _BASE_CY),    # 0°:   stem upper-left
            (-_BASE_CX, -_BASE_CY),   # 180°: stem lower-right
        ]

        P = 12

        win = ctk.CTkToplevel(self.root)
        win.title(tr("🍆 Grid Practice"))
        win.configure(fg_color=self._C_BG)
        win.transient(self.root)
        win.grab_set()
        win.geometry("510x700")
        win.resizable(False, False)

        ctk.CTkLabel(win, text=tr("🍆  Grid Navigator Practice"),
                     font=ctk.CTkFont(size=16, weight="bold"),
                     text_color=self._C_TEXT).pack(pady=(P, 2))
        ctk.CTkLabel(win,
                     text=tr("1–9  zoom in    •    here  confirm    •    back  step back    •    cancel  reset"),
                     font=ctk.CTkFont(size=13), text_color=self._C_TEXT,
                     wraplength=490).pack(pady=(0, 6))

        # Score + status row
        hdr = ctk.CTkFrame(win, fg_color="transparent")
        hdr.pack(fill=tk.X, padx=P, pady=(0, 4))
        score_lbl = ctk.CTkLabel(hdr, text=tr("Score: 0"),
                                  font=ctk.CTkFont(size=13, weight="bold"),
                                  text_color="#44ff44")
        score_lbl.pack(side=tk.LEFT, padx=8)
        status_lbl = ctk.CTkLabel(hdr, text=tr("Navigate to the 🌿 stem!"),
                                   font=ctk.CTkFont(size=11), text_color=self._C_TEXT_DIM)
        status_lbl.pack(side=tk.LEFT, padx=4)

        # Game canvas
        cv = tk.Canvas(win, width=CV_W, height=CV_H, bg="#12121e",
                       highlightthickness=2, highlightbackground=self._C_BORDER)
        cv.pack(pady=4)

        # ── game state ────────────────────────────────────────────────────────
        gs = {
            "score": 0,
            "ex": CV_W // 2, "ey": CV_H // 2,   # emoji centre position
            "tick": 0,
            "celebrate": 0,    # frames of celebration remaining
            "miss_flash": 0,   # frames of red-X flash remaining
            "miss_x": 0, "miss_y": 0,
            "pin_x": None, "pin_y": None, "pin_w": None, "pin_h": None,
            "grid_x": 0, "grid_y": 0, "grid_w": CV_W, "grid_h": CV_H,
            "grid_stack": [],   # undo stack for back-stepping
            "rot_idx": 0,
            "running": True,
        }

        # Pre-render 4 rotations with the target ring baked in — ring always tracks the stem
        _eggplant_photos = []
        try:
            from PIL import Image, ImageDraw, ImageFont, ImageTk as _ITk
            _ef  = ImageFont.truetype("C:/Windows/Fonts/seguiemj.ttf", 120)
            _bim = Image.new("RGBA", (160, 160), (0, 0, 0, 0))
            _bdraw = ImageDraw.Draw(_bim)
            _bdraw.text((80, 80), "🍆", font=_ef, anchor="mm", embedded_color=True)
            # Bake ring at stem position in the unrotated image
            _rx0 = 80 + _BASE_CX - STEM_RX - 3
            _ry0 = 80 + _BASE_CY - STEM_RY - 3
            _rx1 = 80 + _BASE_CX + STEM_RX + 3
            _ry1 = 80 + _BASE_CY + STEM_RY + 3
            _bdraw.ellipse([_rx0, _ry0, _rx1, _ry1], outline=(68, 255, 68, 255), width=3)
            for ang in _ROT_ANGLES:
                _eggplant_photos.append(_ITk.PhotoImage(_bim.rotate(ang)))
        except Exception:
            pass   # fallback: create_text + separate ring below

        def _reset_eggplant():
            margin = 80
            gs["ex"] = random.randint(margin, CV_W - margin)
            gs["ey"] = random.randint(margin, CV_H - margin)
            gs["grid_x"], gs["grid_y"] = 0, 0
            gs["grid_w"], gs["grid_h"] = CV_W, CV_H
            gs["grid_stack"].clear()
            gs["pin_x"] = gs["pin_y"] = None
            gs["rot_idx"] = random.randint(0, len(_ROT_ANGLES) - 1)

        def _draw():
            cv.delete("all")
            ex, ey = int(gs["ex"]), int(gs["ey"])
            t = gs["tick"]

            # Emoji — wobble during celebration
            wobble = 0
            if gs["celebrate"] > 0:
                wobble = 7 if (t // 3) % 2 == 0 else -7
            cx = ex + wobble
            ri = gs["rot_idx"]
            if _eggplant_photos:
                cv.create_image(cx, ey, image=_eggplant_photos[ri], anchor="center")
            else:
                # Fallback: emoji text + separate ring
                cv.create_text(cx, ey, text="🍆", font=EMOJI_FONT, anchor="center")
                scx_off, scy_off = _ROT_OFFSETS[ri]
                scx, scy = cx + scx_off, ey + scy_off
                cv.create_oval(scx - STEM_RX - 3, scy - STEM_RY - 3,
                               scx + STEM_RX + 3, scy + STEM_RY + 3,
                               outline="#44ff44", width=3, fill="")

            # Dim overlay outside active grid region
            gx, gy, gw, gh = gs["grid_x"], gs["grid_y"], gs["grid_w"], gs["grid_h"]
            if gw < CV_W or gh < CV_H:
                for rx, ry, rw, rh in [
                    (0,       0,        CV_W,          gy),
                    (0,       gy + gh,  CV_W,          CV_H - gy - gh),
                    (0,       gy,       gx,            gh),
                    (gx + gw, gy,       CV_W - gx - gw, gh),
                ]:
                    if rw > 0 and rh > 0:
                        cv.create_rectangle(rx, ry, rx + rw, ry + rh,
                                             fill="#000000", stipple="gray50", outline="")

            # Grid lines inside active region
            cw, ch = gw // 3, gh // 3
            for c in range(1, 3):
                cv.create_line(gx + c * cw, gy, gx + c * cw, gy + gh,
                               fill="#888888", width=1)
            for r in range(1, 3):
                cv.create_line(gx, gy + r * ch, gx + gw, gy + r * ch,
                               fill="#888888", width=1)
            cv.create_rectangle(gx, gy, gx + gw, gy + gh, outline="#cccccc", width=2)
            for i in range(9):
                r, c = divmod(i, 3)
                ncx = gx + c * cw + cw // 2
                ncy = gy + r * ch + ch // 2
                cv.create_text(ncx, ncy, text=str(i + 1),
                               fill="#aaaaaa", font=("Segoe UI", 10, "bold"))

            # Pinned region box
            if gs["pin_x"] is not None:
                px, py, pw, ph = gs["pin_x"], gs["pin_y"], gs["pin_w"], gs["pin_h"]
                cv.create_rectangle(px, py, px + pw, py + ph,
                                     outline="#44aaff", width=3)

            # Miss flash: red X
            if gs["miss_flash"] > 0:
                fx, fy, r2 = int(gs["miss_x"]), int(gs["miss_y"]), 18
                cv.create_line(fx - r2, fy - r2, fx + r2, fy + r2,
                               fill="#ff3333", width=4)
                cv.create_line(fx + r2, fy - r2, fx - r2, fy + r2,
                               fill="#ff3333", width=4)

            # Hit banner
            if gs["celebrate"] > 20:
                cv.create_text(cx, ey - 100, text=tr("✓  HIT!"), fill="#44ff44",
                               font=("Segoe UI", 17, "bold"))

        def _tick():
            if not gs["running"]:
                return
            gs["tick"] += 1

            if gs["celebrate"] > 0:
                gs["celebrate"] -= 1
                if gs["celebrate"] == 0:
                    _reset_eggplant()
            if gs["miss_flash"] > 0:
                gs["miss_flash"] -= 1

            try:
                _draw()
            except Exception:
                pass
            try:
                win.after(33, _tick)
            except Exception:
                pass

        def _zoom_cell(n: int):
            gx, gy, gw, gh = gs["grid_x"], gs["grid_y"], gs["grid_w"], gs["grid_h"]
            gs["grid_stack"].append((gx, gy, gw, gh))
            cw, ch = gw // 3, gh // 3
            r, c = divmod(n - 1, 3)
            gs["grid_x"] = gx + c * cw
            gs["grid_y"] = gy + r * ch
            gs["grid_w"] = cw
            gs["grid_h"] = ch
            gs["pin_x"] = gs["pin_y"] = None
            depth = len(gs["grid_stack"])
            status_lbl.configure(text=tr("Cell {} — 'here' to pin · up/down/left/right to nudge · 'back' to step back · 'cancel' to reset").format(n))

        def _confirm_pin():
            gx, gy, gw, gh = gs["grid_x"], gs["grid_y"], gs["grid_w"], gs["grid_h"]
            if gw >= CV_W and gh >= CV_H:
                status_lbl.configure(text=tr("Zoom in first — say a number!"))
                return
            gs["pin_x"], gs["pin_y"] = gx, gy
            gs["pin_w"], gs["pin_h"] = gw, gh
            pin_cx = gx + gw // 2
            pin_cy = gy + gh // 2
            ex, ey = int(gs["ex"]), int(gs["ey"])
            scx_off, scy_off = _ROT_OFFSETS[gs["rot_idx"]]
            stem_left  = ex + scx_off - STEM_RX
            stem_right = ex + scx_off + STEM_RX
            stem_top   = ey + scy_off - STEM_RY
            stem_bot   = ey + scy_off + STEM_RY
            hit = stem_left <= pin_cx <= stem_right and stem_top <= pin_cy <= stem_bot
            if hit:
                gs["score"] += 1
                score_lbl.configure(text=tr("Score: {}").format(gs['score']))
                gs["celebrate"] = 45
                status_lbl.configure(text=tr("🎉 Nice shot!"))
            else:
                gs["miss_x"] = pin_cx
                gs["miss_y"] = pin_cy
                gs["miss_flash"] = 28
                gs["grid_x"], gs["grid_y"] = 0, 0
                gs["grid_w"], gs["grid_h"] = CV_W, CV_H
                gs["grid_stack"].clear()
                gs["pin_x"] = gs["pin_y"] = None
                status_lbl.configure(text=tr("Miss — aim for the green ring at the top!"))

        def _cancel_grid():
            """Step back one zoom level. 'again' does a full reset."""
            if gs["grid_stack"]:
                gx, gy, gw, gh = gs["grid_stack"].pop()
                gs["grid_x"], gs["grid_y"] = gx, gy
                gs["grid_w"], gs["grid_h"] = gw, gh
                gs["pin_x"] = gs["pin_y"] = None
                if gs["grid_stack"]:
                    status_lbl.configure(text=tr("Stepped back — say a number or 'back' again"))
                else:
                    status_lbl.configure(text=tr("Back to full view — navigate to the 🌿 stem!"))
            else:
                status_lbl.configure(text=tr("Already at full view!"))

        def _reset_grid():
            gs["grid_x"], gs["grid_y"] = 0, 0
            gs["grid_w"], gs["grid_h"] = CV_W, CV_H
            gs["grid_stack"].clear()
            gs["pin_x"] = gs["pin_y"] = None
            status_lbl.configure(text=tr("Reset — navigate to the 🌿 stem!"))

        def _nudge_grid(dx: int, dy: int):
            gw, gh = gs["grid_w"], gs["grid_h"]
            if gw >= CV_W and gh >= CV_H:
                status_lbl.configure(text=tr("Zoom in first — say a number!"))
                return
            gs["grid_x"] = max(0, min(CV_W - gw, gs["grid_x"] + dx))
            gs["grid_y"] = max(0, min(CV_H - gh, gs["grid_y"] + dy))
            gs["pin_x"] = gs["pin_y"] = None

        # ── voice engine for practice ─────────────────────────────────────────
        def _on_practice_voice(kw: str):
            n = VoiceEngine._NUM_MAP.get(kw)
            if n is not None:
                win.after(0, lambda k=n: _zoom_cell(k))
                return
            if kw in ("select", "here", "confirm"):
                win.after(0, _confirm_pin)
            elif kw == "back":
                win.after(0, _cancel_grid)
            elif kw in ("cancel", "again"):
                win.after(0, _reset_grid)
            elif kw == "up":
                win.after(0, lambda: _nudge_grid(0, -gs["grid_h"]))
            elif kw == "down":
                win.after(0, lambda: _nudge_grid(0,  gs["grid_h"]))
            elif kw == "left":
                win.after(0, lambda: _nudge_grid(-gs["grid_w"], 0))
            elif kw == "right":
                win.after(0, lambda: _nudge_grid( gs["grid_w"], 0))

        _practice_engine = VoiceEngine(
            VoiceEngine.model_path_default(),
            device_name=self._bondage_mic_device,
        )
        _practice_engine.set_safeword(self._bondage_safeword)
        _practice_engine.set_callback(_on_practice_voice)
        _practice_engine.start()

        def _stop_practice_engine():
            gs["running"] = False
            try:
                _practice_engine.stop()
            except Exception:
                pass

        win.protocol("WM_DELETE_WINDOW", lambda: (_stop_practice_engine(), win.destroy()))

        # Start bondage session when done
        def _done():
            _stop_practice_engine()
            win.destroy()
            self._start_bondage_session()

        # ── mic level meter ───────────────────────────────────────────────────
        mic_row = ctk.CTkFrame(win, fg_color="transparent")
        mic_row.pack(fill=tk.X, padx=P, pady=(4, 0))
        ctk.CTkLabel(mic_row, text=tr("🎙 MIC"),
                     font=ctk.CTkFont(size=10), text_color=self._C_TEXT_DIM).pack(side=tk.LEFT)
        mic_bar_cv = tk.Canvas(mic_row, width=260, height=12, bg=self._C_SURFACE2,
                               highlightthickness=0)
        mic_bar_cv.pack(side=tk.LEFT, padx=(6, 0))
        mic_status = ctk.CTkLabel(mic_row, text=tr("loading…"),
                                   font=ctk.CTkFont(size=9), text_color=self._C_TEXT_DIM, width=60)
        mic_status.pack(side=tk.LEFT, padx=(6, 0))

        def _mic_poll():
            if not gs["running"]:
                return
            level = _practice_engine.level
            alive = _practice_engine.running
            mic_bar_cv.delete("all")
            if alive:
                w = int(level * 260)
                color = "#44ff44" if level < 0.6 else "#ffaa00"
                if w > 0:
                    mic_bar_cv.create_rectangle(0, 0, w, 12, fill=color, outline="")
                mic_status.configure(text="listening", text_color="#44ff44")
            else:
                mic_status.configure(text=tr("loading…"), text_color=self._C_TEXT_DIM)
            try:
                win.after(80, _mic_poll)
            except Exception:
                pass

        nav = ctk.CTkFrame(win, fg_color="transparent")
        nav.pack(fill=tk.X, padx=P, pady=(6, P))
        ctk.CTkButton(nav, text=tr("← Back"),
                      command=lambda: (_stop_practice_engine(), win.destroy()),
                      font=ctk.CTkFont(size=11), height=34,
                      fg_color=self._C_SURFACE2, hover_color="#4a4a4a",
                      text_color=self._C_TEXT).pack(side=tk.LEFT)
        ctk.CTkButton(nav, text=tr("Start Bondage Mode →"), command=_done,
                      font=ctk.CTkFont(size=12, weight="bold"), height=34,
                      fg_color="#4a0a6a", hover_color="#6a1a8a",
                      text_color="white").pack(side=tk.RIGHT)

        # Kick off
        _reset_eggplant()
        _tick()
        _mic_poll()

    def _start_bondage_session(self):
        """Activate bondage mode: start VoiceEngine, update button, show indicator."""
        model_path = VoiceEngine.model_path_default()
        if self._voice_engine:
            self._voice_engine.stop()
        self._voice_engine = VoiceEngine(
            model_path,
            device_name=self._bondage_mic_device,
        )
        self._voice_engine.set_safeword(self._bondage_safeword)
        self._voice_engine.set_callback(self._voice_raw_cb)
        self._voice_engine.start()
        self._bondage_active     = True
        self._bondage_configured = True
        self._bondage_btn.configure(text=tr("🎙 BONDAGE"), fg_color="#6a1a8a",
                                    hover_color="#4a0a6a", border_color="#9a3aaa",
                                    text_color="white")
        self._snark_label.configure(
            text=tr('🎙 Bondage active — safeword: "{}"').format(self._bondage_safeword),
            text_color="#c080ff")
        log.info(f"Bondage mode started (safeword={self._bondage_safeword!r}, "
                 f"mic={self._bondage_mic_device or 'default'})")
        self._show_voice_cheatsheet()

    def _show_voice_cheatsheet(self):
        """Pop a closeable reference window listing all voice commands (2-col grid)."""
        ref = ctk.CTkToplevel(self.root)
        ref.title("🎙 Voice Commands")
        ref.configure(fg_color="#0d0012")
        ref.attributes("-topmost", True)
        ref.resizable(False, False)
        ref.geometry("780x600")

        _head = ctk.CTkFont(size=13, weight="bold")
        _body = ctk.CTkFont(size=12)
        _dim  = ctk.CTkFont(size=11)
        P = 12

        ctk.CTkLabel(ref, text=tr("🎙 Voice Commands"),
                     font=ctk.CTkFont(size=18, weight="bold"),
                     text_color="#c080ff").pack(pady=(P, 6))

        # ── Safeword — full-width banner ──────────────────────────────────────
        sw_card = ctk.CTkFrame(ref, fg_color="#2a0008", corner_radius=8,
                               border_width=2, border_color="#cc0000")
        sw_card.pack(fill=tk.X, padx=P, pady=(0, 8))
        sw_inner = ctk.CTkFrame(sw_card, fg_color="transparent")
        sw_inner.pack(pady=8)
        ctk.CTkLabel(sw_inner, text=tr("🛑 SAFEWORD"),
                     font=ctk.CTkFont(size=14, weight="bold"),
                     text_color="#ff6060").pack(side=tk.LEFT, padx=(0, 12))
        ctk.CTkLabel(sw_inner, text=f'"{self._bondage_safeword}"',
                     font=ctk.CTkFont(size=26, weight="bold"),
                     text_color="#ff2020").pack(side=tk.LEFT, padx=(0, 12))
        ctk.CTkLabel(sw_inner,
                     text=tr("hard stop — mic stays on.\nsay \"resume session\" to continue"),
                     font=_dim, text_color="#ff8080", justify="left").pack(side=tk.LEFT)

        # ── 2-column grid of sections ─────────────────────────────────────────
        grid = ctk.CTkFrame(ref, fg_color="transparent")
        grid.pack(fill=tk.BOTH, expand=True, padx=P, pady=(0, 4))
        grid.grid_columnconfigure(0, weight=1, uniform="col")
        grid.grid_columnconfigure(1, weight=1, uniform="col")

        def _section(r, c, title, rows, title_color="#c080ff"):
            f = ctk.CTkFrame(grid, fg_color="#1e002c", corner_radius=6)
            f.grid(row=r, column=c, sticky="nsew", padx=4, pady=4)
            ctk.CTkLabel(f, text=title, font=_head,
                         text_color=title_color).pack(anchor="w", padx=10, pady=(8, 4))
            for cmd, desc in rows:
                row = ctk.CTkFrame(f, fg_color="transparent")
                row.pack(fill=tk.X, padx=10, pady=1)
                ctk.CTkLabel(row, text=cmd, font=_body, text_color="#e0d0ff",
                             width=170, anchor="w").pack(side=tk.LEFT)
                ctk.CTkLabel(row, text=desc, font=_dim, text_color="#9a8aaa",
                             anchor="w", justify="left").pack(side=tk.LEFT)
            ctk.CTkFrame(f, height=1, fg_color="#2d1a40").pack(fill=tk.X, padx=10, pady=(6, 8))

        _section(0, 0, "🎬  Session", [
            ('"let me cum" / "please"', "beg — rolls the dice"),
            ('"came" / "cumming"',      "log that you came"),
            ('"pause"',                 "freeze tracking"),
            ('"resume"',                "unfreeze tracking"),
        ])

        _section(0, 1, "📏  Height lines", [
            ('"erect up / down"',       "±3% frame height"),
            ('"flaccid up / down"',     "±3% frame height"),
            ('"edging up / down"',      "±3% frame height"),
            ('"set lines"',             "auto-place all three"),
        ])

        _section(1, 0, "🎯  Head tracking", [
            ('"find head"',             "open grid navigator"),
            ('"one" – "nine"',          "zoom into that cell"),
            ('"here" / "select"',       "lock on & reanchor"),
            ('"back" / "cancel"',       "reset grid / exit"),
        ])

        _section(1, 1, "🚫  Exclusion zones", [
            ('"exclude"',               "open grid to mark a zone"),
            ('"one" – "nine"',          "zoom into that cell"),
            ('"exclude" / "here"',      "add that region"),
            ('"back" / "cancel"',       "exit without adding"),
            ('"clear exclude"',         "remove all zones"),
        ], title_color="#ff8800")

        _section(2, 0, "🖥️  Feed source", [
            ('"switch source"',         "list open windows 1–9"),
            ('"one" – "nine"',          "switch to that window"),
            ('"cancel" / "again"',      "abort picker"),
        ])

        _section(2, 1, "🛑  Standby", [
            ('safeword (above)',        "hard stop into standby"),
            ('"resume session"',        "leave standby, continue"),
        ], title_color="#ff6060")

        ctk.CTkButton(ref, text=tr("Got it — close"),
                      command=ref.destroy,
                      font=ctk.CTkFont(size=13, weight="bold"), height=34,
                      fg_color="#2d1a40", hover_color="#4a0a6a",
                      text_color="#c080ff", corner_radius=6,
                      ).pack(pady=(2, P), padx=P, fill=tk.X)

    def _stop_bondage_mode(self):
        """Deactivate bondage mode cleanly (full stop, including engine)."""
        self._bondage_active  = False
        self._bondage_standby = False
        self._grid_active     = False
        self._grid_mode       = 'head'
        self._grid_region     = None
        self._grid_depth      = 0
        self._source_picker_close()
        if self._voice_engine:
            self._voice_engine.stop()
            self._voice_engine = None
        try:
            self._bondage_btn.configure(text=tr("🎙 BONDAGE"), fg_color="transparent",
                                        hover_color="#1a0028",
                                        border_color="#8a2a9a", text_color="#c080ff")
        except Exception:
            pass
        log.info("Bondage mode stopped")

    # ── voice keyword dispatcher ──────────────────────────────────────────────

    def _voice_raw_cb(self, kw: str):
        """Called from VoiceEngine audio thread — posts to main thread."""
        self._call_on_main(self._on_voice_keyword, kw)

    def _on_voice_keyword(self, kw: str):
        """Dispatches a recognised keyword (runs on main thread)."""
        # Standby after safeword — only "resume session" (or safeword repeat) accepted
        if self._bondage_standby:
            if kw == "resume session":
                self._voice_resume_session()
            return

        if not self._bondage_active:
            return
        log.info(f"Voice command: {kw!r}")

        # ── safeword — hard stop, always wins ────────────────────────────────
        if kw == "safeword":
            self._voice_safeword()
            return

        # ── grid navigator takes over when active ─────────────────────────────
        if self._grid_active:
            n = VoiceEngine._NUM_MAP.get(kw)
            if n is not None:
                self._grid_zoom(n)
                return
            if kw in ("select", "here", "confirm") or \
               (kw == "exclude" and self._grid_mode == 'exclude'):
                self._grid_confirm()
                return
            if kw in ("cancel", "again", "back"):
                self._grid_cancel()
                return
            return  # ignore other keywords while grid is up

        # ── source picker takes over when active ──────────────────────────────
        if self._source_picker_active:
            n = VoiceEngine._NUM_MAP.get(kw)
            if n is not None:
                self._source_picker_select(n)
                return
            if kw in ("cancel", "again"):
                self._source_picker_close()
                self._snark_label.configure(text=tr("🎙 Source switch cancelled"),
                                            text_color="#F5A623")
                return
            return  # absorb everything else while picker is up

        # ── regular session commands ──────────────────────────────────────────
        if kw in ("came", "cumming"):
            self._on_cum(source="voice")
            return
        if kw in ("let me cum", "please"):
            # Beg for permission — fires the same roll as the button
            # (grant / denial / Evil-Mode ruin all apply). Cooldown is respected.
            self._on_letmecum()
            return
        if kw == "pause":
            if not self.tracking_paused:
                self.tracking_paused = True
                self._snark_label.configure(text=tr("🎙 Paused (say 'resume')"),
                                            text_color="#F5A623")
            return
        if kw == "resume":
            self.tracking_paused = False
            self._snark_label.configure(text="", text_color="#ff4444")
            return
        if kw == "find head":
            self._grid_start()
            return
        if kw == "exclude":
            self._grid_start_exclude()
            return
        if kw == "clear exclude":
            self._exclusion_zones.clear()
            self._save_config()
            self._snark_label.configure(text=tr("🎙 Exclusion zones cleared"),
                                        text_color="#F5A623")
            log.info("Voice: all exclusion zones cleared")
            return
        if kw == "set lines":
            self._voice_set_lines()
            return
        if kw == "switch source":
            self._voice_switch_source()
            return
        # Height adjustments
        for which in ("erect", "flaccid", "edging"):
            if kw == f"{which} up":
                self._voice_adjust_height(which.capitalize(), -1)
                return
            if kw == f"{which} down":
                self._voice_adjust_height(which.capitalize(), +1)
                return

    def _voice_safeword(self):
        """Safeword — pause everything but keep mic alive for 'resume session'."""
        # EMERGENCY STOP FIRST: kill all output immediately and unconditionally,
        # before touching any state — a safeword must never depend on state logic
        # to cut the stim (e.g. mid-grant, mid-ruin, or during the undo window).
        try:
            self._set_all_outputs(0.0, instant=True)
        except Exception:
            pass
        self._cancel_ruin()   # stop any ruin pulses immediately
        # Partial stop: don't kill the engine
        self._bondage_active  = False
        self._bondage_standby = True
        self._grid_active     = False
        self._grid_mode       = 'head'
        self._grid_region     = None
        self._grid_depth      = 0
        self._source_picker_close()
        self.tracking_paused  = True
        # End the edge/cum session if running
        try:
            if not self._cum_stopped:
                self._on_cum(source="voice")
        except Exception:
            pass
        # UI — repurpose the bondage button as a bright, unmissable RESUME control.
        # Standby pauses tracking (the video freezes), so voice is no longer the
        # only way out — if the mic doesn't catch "resume session", this button does
        # the same thing. Restored to the normal BONDAGE button by _voice_resume_session.
        try:
            self._bondage_btn.configure(text=tr("▶ RESUME SESSION"),
                                        command=self._voice_resume_session,
                                        fg_color="#3EC941", hover_color="#32a435",
                                        border_color="#3EC941", text_color="white")
        except Exception:
            pass
        try:
            self._snark_label.configure(
                text=tr('🛑 SAFEWORD — say "resume session" or click ▶ RESUME'),
                text_color="#FF4444")
        except Exception:
            pass
        log.warning("SAFEWORD triggered — standing by for 'resume session'")

    def _voice_resume_session(self):
        """Resume bondage mode after safeword — engine was never stopped."""
        self._bondage_standby = False
        self._bondage_active  = True
        self.tracking_paused  = False
        try:
            # Restore the button's normal command too — safeword had repointed it
            # at _voice_resume_session.
            self._bondage_btn.configure(text=tr("🎙 BONDAGE"), command=self._open_bondage_splash,
                                        fg_color="#6a1a8a", hover_color="#4a0a6a",
                                        border_color="#9a3aaa", text_color="white")
        except Exception:
            pass
        try:
            self._snark_label.configure(text=tr("🎙 Session resumed"),
                                        text_color="#3EC941")
        except Exception:
            pass
        log.info("Bondage mode resumed after safeword")

    def _voice_adjust_height(self, which: str, direction: int):
        """Move a height line up (direction=-1) or down (+1) by ~3% of frame height."""
        fh = getattr(self, '_disp_frame_h', None) or self.rel_box.get('height', 300)
        step = max(5, int(fh * 0.03))
        current = self.heights.get(which)
        if current is None:
            # Line not set — place it at current head position
            current = self.head_y
        new_val = max(0, min(fh, current + direction * step))
        self.heights[which] = new_val
        self._disable_auto()
        self._save_config()
        names = {"Erect": "erect", "Flaccid": "flaccid", "Edging": "edging"}
        arrow = "↑" if direction < 0 else "↓"
        self._snark_label.configure(
            text=tr("🎙 {} {}  ({}px)").format(names.get(which, which), arrow, int(new_val)),
            text_color="#c080ff")
        log.info(f"Voice: {which} line → {new_val}px")

    def _voice_set_lines(self):
        """Auto-position all three lines spread across the current frame height."""
        fh = getattr(self, '_disp_frame_h', None) or self.rel_box.get('height', 300)
        self.heights["Edging"]  = int(fh * 0.20)
        self.heights["Erect"]   = int(fh * 0.50)
        self.heights["Flaccid"] = int(fh * 0.80)
        self._disable_auto()
        self._save_config()
        self._snark_label.configure(text=tr("🎙 Lines set — adjust with voice"),
                                    text_color="#c080ff")
        log.info("Voice: set lines auto-positioned")

    # ── source picker (voice "switch source") ─────────────────────────────────

    def _voice_switch_source(self):
        """Enumerate visible windows, show a numbered picker; user says 1-9."""
        # Collect candidate windows — visible, non-minimised, titled, not us
        candidates: list[tuple[int, str]] = []
        try:
            own_hwnd = int(self.root.wm_frame(), 0)
        except Exception:
            own_hwnd = 0

        def _enum_cb(hwnd, _):
            if not win32gui.IsWindowVisible(hwnd):
                return True
            if win32gui.IsIconic(hwnd):          # minimised
                return True
            title = win32gui.GetWindowText(hwnd).strip()
            if not title:
                return True
            # Skip ourselves
            try:
                if hwnd == own_hwnd:
                    return True
            except Exception:
                pass
            candidates.append((hwnd, title))
            return True

        try:
            win32gui.EnumWindows(_enum_cb, None)
        except Exception as e:
            log.warning(f"switch source: EnumWindows failed: {e}")

        if not candidates:
            self._snark_label.configure(text=tr("🎙 No windows found"), text_color="#F5A623")
            return

        # Limit to 9 and store
        self._source_picker_windows = candidates[:9]
        self._source_picker_active  = True

        # Close any stale picker window (without resetting state)
        try:
            if self._source_picker_win:
                self._source_picker_win.destroy()
        except Exception:
            pass
        self._source_picker_win = None

        win = ctk.CTkToplevel(self.root)
        win.title("🎙 Switch Source")
        win.configure(fg_color="#0d0012")
        win.attributes("-topmost", True)
        win.resizable(False, False)
        win.protocol("WM_DELETE_WINDOW", self._source_picker_close)
        self._source_picker_win = win

        _hfont = ctk.CTkFont(size=13, weight="bold")
        _bfont = ctk.CTkFont(size=11)

        ctk.CTkLabel(win, text=tr("🎙 Switch Source"),
                     font=_hfont, text_color="#c080ff").pack(pady=(10, 4), padx=14)
        ctk.CTkLabel(win, text=tr('Say a number, or "cancel"'),
                     font=ctk.CTkFont(size=9, slant="italic"),
                     text_color="#7a6a90").pack(pady=(0, 8))

        scroll_frame = ctk.CTkScrollableFrame(win, fg_color="#0d0012",
                                              border_width=0, height=300)
        scroll_frame.pack(fill=tk.X, padx=12, pady=(0, 4))

        for i, (hwnd, title) in enumerate(self._source_picker_windows, 1):
            short = title if len(title) <= 44 else title[:41] + "…"
            row = ctk.CTkFrame(scroll_frame, fg_color="#1e002c", corner_radius=6)
            row.pack(fill=tk.X, padx=0, pady=2)
            ctk.CTkLabel(row, text=f" {i} ", font=_hfont,
                         text_color="#ff90ff", width=28, anchor="e").pack(side=tk.LEFT, padx=(6, 0))
            ctk.CTkLabel(row, text=short, font=_bfont,
                         text_color="#e0d0ff", anchor="w").pack(side=tk.LEFT, padx=6, pady=6)

        ctk.CTkButton(win, text=tr("Cancel"),
                      command=self._source_picker_close,
                      font=ctk.CTkFont(size=10), height=28,
                      fg_color="#2d1a40", hover_color="#4a0a6a",
                      text_color="#c080ff", corner_radius=6,
                      ).pack(pady=(6, 12), padx=12, fill=tk.X)

        self._snark_label.configure(
            text="🎙 Say 1-%d to pick source, 'cancel' to abort" % len(self._source_picker_windows),
            text_color="#c080ff")
        log.info(f"Source picker: {len(self._source_picker_windows)} windows listed")

    def _source_picker_select(self, n: int):
        """User said number n — switch feed to that window."""
        if n < 1 or n > len(self._source_picker_windows):
            return
        hwnd, title = self._source_picker_windows[n - 1]
        self._source_picker_close()

        try:
            wx, wy, wr, wb = win32gui.GetWindowRect(hwnd)
            ww = wr - wx
            wh = wb - wy
            if ww <= 0 or wh <= 0:
                raise ValueError("zero-size window")
            new_rel_box = {
                'x1': 0, 'y1': 0,
                'x2': ww, 'y2': wh,
                'width': ww, 'height': wh,
            }
            self.hwnd    = hwnd
            self.rel_box = new_rel_box
            self._reset_heights()
            self._snark_label.configure(
                text=tr("🎙 Source → {}… — say 'find head' to reanchor").format(title[:30]),
                text_color="#3EC941")
            log.info(f"Source switched to hwnd={hwnd} '{title}'")
            # Kick grid so user can reanchor head hands-free
            self._grid_start()
        except Exception as e:
            log.warning(f"switch source: failed to switch to hwnd={hwnd}: {e}")
            self._snark_label.configure(text=tr("🎙 Source switch failed"), text_color="#FF4444")

    def _source_picker_close(self):
        """Dismiss the picker window and reset state."""
        self._source_picker_active  = False
        self._source_picker_windows = []
        try:
            if self._source_picker_win:
                self._source_picker_win.destroy()
        except Exception:
            pass
        self._source_picker_win = None

    # ── grid navigator ────────────────────────────────────────────────────────

    def _grid_start(self):
        """Activate the grid overlay in head-reanchor mode."""
        self._grid_active = True
        self._grid_mode   = 'head'
        self._grid_region = None
        self._grid_depth  = 0
        self._snark_label.configure(text=tr("🎙 Grid — say 1-9 to zoom, 'here' to lock"),
                                    text_color="#c080ff")
        log.info("Grid navigator: head mode started")

    def _grid_start_exclude(self):
        """Activate the grid overlay in exclusion-zone mode."""
        self._grid_active = True
        self._grid_mode   = 'exclude'
        self._grid_region = None
        self._grid_depth  = 0
        self._snark_label.configure(
            text=tr("🎙 Exclude zone — say 1-9 to zoom, 'exclude'/'here' to add"),
            text_color="#ff8800")
        log.info("Grid navigator: exclude mode started")

    def _grid_zoom(self, cell: int):
        """Zoom into cell 1-9 within the current grid region."""
        fh = getattr(self, '_disp_frame_h', None) or self.rel_box.get('height', 300)
        fw = getattr(self, '_disp_frame_w', None) or self.rel_box.get('width', 400)
        if self._grid_region is None:
            rx, ry, rw, rh = 0, 0, fw, fh
        else:
            rx, ry, rw, rh = self._grid_region
        cw, ch = rw // 3, rh // 3
        r, c   = divmod(cell - 1, 3)
        self._grid_region = (rx + c * cw, ry + r * ch, cw, ch)
        self._grid_depth  = min(self._grid_depth + 1, 3)
        self._snark_label.configure(
            text=tr("🎙 Grid cell {} (zoom {}) — say 'here' to lock, or zoom more").format(cell, self._grid_depth),
            text_color="#c080ff")
        log.info(f"Grid zoom: cell={cell} region={self._grid_region} depth={self._grid_depth}")

    def _grid_confirm(self):
        """Confirm current grid region — reanchor head or add exclusion zone."""
        if getattr(self, '_grid_mode', 'head') == 'exclude':
            self._grid_confirm_exclude()
            return
        fh = getattr(self, '_disp_frame_h', None) or self.rel_box.get('height', 300)
        fw = getattr(self, '_disp_frame_w', None) or self.rel_box.get('width', 400)
        if self._grid_region is None:
            rx, ry, rw, rh = 0, 0, fw, fh
        else:
            rx, ry, rw, rh = self._grid_region
        cx = rx + rw // 2
        cy = ry + rh // 2
        # Build a small bbox around the confirmed point
        # Use last known bbox size as reference, or fall back to 60×60
        pw, ph = self.last_bbox[2], self.last_bbox[3]
        new_bbox = (max(0, cx - pw // 2), max(0, cy - ph // 2), pw, ph)
        # Reinit tracker — grab latest frame from queue, or capture fresh if empty
        frame = None
        try:
            frame = self._frame_queue.get_nowait()
            self._frame_queue.put_nowait(frame)  # put it back
        except Exception:
            pass
        if frame is None:
            frame = capture_window_region(self.hwnd, self.rel_box)
        self._grid_active = False
        self._grid_mode   = 'head'
        self._grid_region = None
        self._grid_depth  = 0
        if frame is not None:
            self._tracker_init(frame, new_bbox)
            self.last_bbox   = new_bbox
            self.head_y      = cy
            self.tracking_ok = True
            self._head_y_history.clear()
            self._snark_label.configure(text=tr("🎙 Head reselected ✓"), text_color="#3EC941")
            log.info(f"Grid confirm: head reanchored at ({cx}, {cy})")
        else:
            self._snark_label.configure(text=tr("🎙 Confirm failed — feed gone?"),
                                        text_color="#FF4444")
            log.warning("Grid confirm: could not grab frame to reinit tracker")

    def _grid_confirm_exclude(self):
        """Add current grid region as an exclusion zone."""
        fh = getattr(self, '_disp_frame_h', None) or self.rel_box.get('height', 300)
        fw = getattr(self, '_disp_frame_w', None) or self.rel_box.get('width', 400)
        if self._grid_region is None:
            rx, ry, rw, rh = 0, 0, fw, fh
        else:
            rx, ry, rw, rh = self._grid_region
        self._grid_active = False
        self._grid_mode   = 'head'
        self._grid_region = None
        self._grid_depth  = 0
        self._exclusion_zones.append((rx, ry, rw, rh))
        self._save_config()
        n = len(self._exclusion_zones)
        self._snark_label.configure(
            text=tr("🎙 Exclusion zone added ({} total) — say 'clear exclude' to remove all").format(n),
            text_color="#3EC941")
        log.info(f"Voice: exclusion zone added ({rx},{ry},{rw},{rh}), total={n}")

    def _grid_cancel(self):
        """Reset grid back to full frame, or exit if already at depth 0."""
        if self._grid_depth > 0:
            self._grid_region = None
            self._grid_depth  = 0
            self._snark_label.configure(
                text=tr("🎙 Grid reset — say 1-9 to zoom"), text_color="#c080ff")
        else:
            self._grid_active = False
            self._snark_label.configure(text=tr("🎙 Grid cancelled"), text_color="#F5A623")
        log.info("Grid cancel")

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
        # Cancel the splash blink timer before teardown so its pending
        # after() callback doesn't fire on a destroyed interp ("invalid
        # command name ..._blink" on stderr).
        try:
            if _blink_id[0]:
                root.after_cancel(_blink_id[0])
        except Exception:
            pass
        root.destroy()

    # ── PNG path (bundled resource) ────────────────────────────────────────────
    splash_path = pathlib.Path(resource_path("splash.png"))

    if splash_path.exists():
        try:
            from PIL import Image, ImageDraw, ImageFont, ImageTk

            img = Image.open(splash_path).convert("RGBA")
            W, H = img.size

            # Flatten RGBA onto the dark background colour (#0d0d0d = 13,13,13)
            bg_flat = Image.new("RGB", (W, H), (13, 13, 13))
            bg_flat.paste(img, mask=img.split()[3])

            # Scale up for high-DPI / 4K displays
            try:
                dpi_scale = root.winfo_fpixels('1i') / 96.0
            except Exception:
                dpi_scale = 1.0
            if dpi_scale > 1.05:
                W = int(W * dpi_scale)
                H = int(H * dpi_scale)
                bg_flat = bg_flat.resize((W, H), Image.LANCZOS)

            photo = ImageTk.PhotoImage(bg_flat)

            lbl = tk.Label(root, image=photo, bg="#0d0d0d",
                           borderwidth=0, highlightthickness=0)
            lbl.image = photo  # keep reference
            lbl.pack()

            # Click-zone over the baked-in button in the PNG.
            # Fractions from make_splash.py: BX1=88,BX2=612,BY1=525,BY2=581 in 700x640
            _btn_y1 = int(0.820 * H)   # 525/640
            _btn_y2 = int(0.908 * H)   # 581/640
            _btn_x1 = int(0.126 * W)   # 88/700
            _btn_x2 = int(0.874 * W)   # 612/700

            def _in_btn(e):
                return _btn_x1 <= e.x <= _btn_x2 and _btn_y1 <= e.y <= _btn_y2

            # Ko-fi footer click zone (full width, bottom ~28px of image)
            _kofi_y1 = int(0.950 * H)
            _kofi_y2 = H

            def _in_kofi(e):
                return _kofi_y1 <= e.y <= _kofi_y2

            def _on_click(e):
                if _in_btn(e):
                    _start()
                elif _in_kofi(e):
                    webbrowser.open("https://ko-fi.com/stimstation")

            lbl.bind("<Button-1>", _on_click)
            lbl.configure(cursor="arrow")

            # Blink: alternate between normal image and a bright-flash
            # version to draw attention to the Start button.
            flash = Image.new("RGBA", (W, H), (0, 0, 0, 0))
            flash_draw = ImageDraw.Draw(flash)
            flash_draw.rounded_rectangle(
                [_btn_x1, _btn_y1, _btn_x2, _btn_y2],
                radius=int(12 * W / 700), fill=(255, 255, 255, 60))
            bg_bright = bg_flat.copy().convert("RGBA")
            bg_bright = Image.alpha_composite(bg_bright, flash).convert("RGB")
            photo_bright = ImageTk.PhotoImage(bg_bright)

            _blink_id = [None]
            blink_state = [False]
            def _blink():
                # Bail if the splash was torn down (Start button / X) before
                # this pending after() fired — avoids "invalid command name
                # ..._blink" noise on stderr.
                try:
                    if not lbl.winfo_exists():
                        return
                except Exception:
                    return
                blink_state[0] = not blink_state[0]
                lbl.configure(image=photo_bright if blink_state[0] else photo)
                _blink_id[0] = root.after(700, _blink)
            _blink_id[0] = root.after(700, _blink)

            # Hand cursor only when entering/leaving the button area
            _cursor = ["arrow"]
            def _on_motion(e):
                want = "hand2" if (_in_btn(e) or _in_kofi(e)) else "arrow"
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

    tk.Label(root, text=tr("VisualStimEdger"), font=f_title,
             bg=BG, fg=TEXT).pack(pady=(P, 2))
    tk.Label(root, text=tr("v{}  ·  edge smarter").format(VERSION), font=f_sub,
             bg=BG, fg=DIM).pack(pady=(0, P//2))

    card = tk.Frame(root, bg=CARD, padx=16, pady=12)
    card.pack(fill="x", padx=P, pady=(0, 10))

    tk.Label(card, text=tr("Before you hit Start"), font=f_head,
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

    btn = tk.Button(root, text=tr("I'm ready — select my camera feed  \u2192"),
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
    # Disable the automatic cyclic garbage collector process-wide. The video feed
    # builds a new Tk PhotoImage every frame and the capture thread allocates hard,
    # so an automatic GC pass periodically fired on a *background* thread and
    # finalized a Tk object created by the main thread — a cross-thread Tcl call
    # that aborts with "Tcl_AsyncDelete: async handler deleted by the wrong thread"
    # (the intermittent mid-session exit-3 crash). The App re-collects on the main
    # thread on a timer (see _gc_collect), so memory stays bounded and every Tk
    # finalizer runs on the thread that owns the interpreter.
    gc.disable()
    # Safety assertion: cyclic GC should now only ever run on the main thread. If it
    # ever fires off-thread, that's the abort condition — log it loudly so we know.
    def _gc_guard(phase, _info):
        if phase == "start" and threading.current_thread() is not threading.main_thread():
            log.warning("cyclic GC ran on background thread %s — exit-3 abort risk",
                        threading.current_thread().name)
    gc.callbacks.append(_gc_guard)

    parser = argparse.ArgumentParser(description="VisualStimEdger")
    parser.add_argument("--debug", action="store_true",
                        help="Enable verbose debug logging to console")
    args = parser.parse_args()

    # Always log to a file, because the windowed exe has no console.
    # Keep last 3 sessions: vse.log (current), vse.log.1, vse.log.2
    log_path = CONFIG_PATH.parent / "vse.log"
    try:
        log_path.parent.mkdir(parents=True, exist_ok=True)
    except Exception:
        pass
    handlers = [logging.StreamHandler()]
    try:
        from logging.handlers import RotatingFileHandler
        fh = RotatingFileHandler(
            log_path, mode="a", encoding="utf-8",
            maxBytes=2 * 1024 * 1024,  # rotate at 2 MB (safety net)
            backupCount=2,
        )
        fh.doRollover()  # always start a fresh file each launch; old ones shift to .1/.2
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

    # Load the UI language up-front so the pre-app screens (region/head selection,
    # their window titles) localise too — the App re-loads it in __init__; this
    # just moves it ahead of show_splash()/select_region()/select_head().
    try:
        _lang0 = (json.loads(CONFIG_PATH.read_text(encoding="utf-8")).get("language", "en")
                  if CONFIG_PATH.exists() else "en")
    except Exception:
        _lang0 = "en"
    _load_catalog(_lang0 if isinstance(_lang0, str) else "en")

    # Native-crash tracer. A segfault in a C extension (OpenCV, Tk, pycaw)
    # kills the process below Python's level — no traceback in vse.log. faulthandler
    # installs OS-level fault handlers (SIGSEGV/SIGFPE/SIGABRT/SIGILL) that dump the
    # current Python stack of every thread to this file at the moment of the fault.
    # The file is kept OPEN for the whole process lifetime (handler writes to the fd).
    try:
        import faulthandler
        crash_path = CONFIG_PATH.parent / "vse_crash.log"
        # Keep a module-level ref so the file isn't GC'd / closed.
        main._crash_fp = open(crash_path, "a", encoding="utf-8", buffering=1)
        main._crash_fp.write(f"\n===== session start v{VERSION} =====\n")
        main._crash_fp.flush()
        # OFF by default: the in-process traps (faulthandler + the SEH minidump
        # filter below) intercept the fault on the crashing thread and can stop the
        # process from ever reaching WER LocalDumps — the one external tool that
        # reliably captures these. With WER armed, leave these off so the crash
        # propagates cleanly. Set VSE_INPROC_TRAP=1 to re-enable them.
        if os.environ.get("VSE_INPROC_TRAP"):
            faulthandler.enable(file=main._crash_fp, all_threads=True)
            log.info(f"faulthandler armed → {crash_path}")
    except Exception as e:
        log.warning(f"faulthandler setup failed: {e}")

    # Windows native minidump on unhandled SEH (e.g. access violation in cv2 /
    # vosk / portaudio). faulthandler is signal()-based and does NOT catch SEH,
    # which is why our crashes leave no trace. SetUnhandledExceptionFilter +
    # MiniDumpWriteDump runs on the faulting thread and writes a .dmp containing
    # every thread's stack — enough for `!analyze -v` to name the faulting module.
    # No admin / registry needed. Defensive: any failure just logs and continues.
    # OFF by default (see note above) — this filter can block WER LocalDumps from
    # capturing the crash. Enable with VSE_INPROC_TRAP=1 only if WER is unavailable.
    if sys.platform == "win32" and os.environ.get("VSE_INPROC_TRAP"):
        try:
            import ctypes
            from ctypes import wintypes

            dump_dir = CONFIG_PATH.parent / "crashdumps"
            dump_dir.mkdir(parents=True, exist_ok=True)

            kernel32 = ctypes.windll.kernel32
            dbghelp  = ctypes.windll.dbghelp

            class _MEI(ctypes.Structure):
                _fields_ = [
                    ("ThreadId",          wintypes.DWORD),
                    ("ExceptionPointers", ctypes.c_void_p),
                    ("ClientPointers",    wintypes.BOOL),
                ]

            _FILTER = ctypes.WINFUNCTYPE(ctypes.c_long, ctypes.c_void_p)

            kernel32.GetCurrentProcess.restype   = wintypes.HANDLE
            kernel32.GetCurrentProcessId.restype = wintypes.DWORD
            kernel32.GetCurrentThreadId.restype  = wintypes.DWORD
            kernel32.CreateFileW.restype  = wintypes.HANDLE
            kernel32.CreateFileW.argtypes = [
                wintypes.LPCWSTR, wintypes.DWORD, wintypes.DWORD,
                ctypes.c_void_p, wintypes.DWORD, wintypes.DWORD, wintypes.HANDLE]
            dbghelp.MiniDumpWriteDump.restype  = wintypes.BOOL
            dbghelp.MiniDumpWriteDump.argtypes = [
                wintypes.HANDLE, wintypes.DWORD, wintypes.HANDLE, wintypes.DWORD,
                ctypes.c_void_p, ctypes.c_void_p, ctypes.c_void_p]

            GENERIC_WRITE         = 0x40000000
            CREATE_ALWAYS         = 2
            FILE_ATTRIBUTE_NORMAL = 0x80
            INVALID_HANDLE_VALUE  = ctypes.c_void_p(-1).value
            DUMP_TYPE             = 0x00000000 | 0x00001000  # Normal | WithThreadInfo
            EXCEPTION_CONTINUE_SEARCH = 0

            _pid       = kernel32.GetCurrentProcessId()
            _dump_path = str(dump_dir / f"vse-crash-{_pid}.dmp")

            @_FILTER
            def _on_unhandled(exc_ptrs):
                try:
                    h = kernel32.CreateFileW(_dump_path, GENERIC_WRITE, 0, None,
                                             CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, None)
                    if h and h != INVALID_HANDLE_VALUE:
                        mei = _MEI()
                        mei.ThreadId          = kernel32.GetCurrentThreadId()
                        mei.ExceptionPointers = exc_ptrs
                        mei.ClientPointers    = False
                        dbghelp.MiniDumpWriteDump(
                            kernel32.GetCurrentProcess(), _pid, h,
                            DUMP_TYPE, ctypes.byref(mei), None, None)
                        kernel32.CloseHandle(h)
                except Exception:
                    pass
                return EXCEPTION_CONTINUE_SEARCH  # let WER/normal teardown proceed

            kernel32.SetUnhandledExceptionFilter(_on_unhandled)
            # Keep refs alive for the whole process lifetime.
            main._mdump_cb   = _on_unhandled
            main._mdump_path = _dump_path
            log.info(f"minidump handler armed → {_dump_path}")
        except Exception as e:
            log.warning(f"minidump handler setup failed: {e}")

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
