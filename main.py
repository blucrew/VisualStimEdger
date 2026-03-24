import cv2
import numpy as np
import time
import threading
import webbrowser
import requests
import websocket
import tkinter as tk
from tkinter import ttk
from PIL import Image, ImageTk
from mss import mss
import win32gui
import win32ui
import win32con
import win32api
import ctypes
import os
import sys
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
from comtypes import CLSCTX_ALL

import atexit

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
            print("[DickDetector] Model loaded OK")
        except Exception as e:
            print(f"[DickDetector] Failed to load model: {e}")

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
        return tuple(boxes[best])

# --- CONFIGURATION ---
VERSION = "1.0.0"
GITHUB_REPO = "blucrew/VisualStimEdger"
RESTIM_HOST = '127.0.0.1'
RESTIM_PORT = 12346
TCODE_AXIS = 'L0'
VOLUME_STEP = 0.05
VOLUME_UPDATE_INTERVAL = 0.5


# Aggressiveness levels: (label, delta multiplier)
AGGR_LEVELS = [
    ("Easy",   0.4),
    ("Middle", 1.0),
    ("Hard",   2.0),
    ("Expert", 4.0),
]

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
            mon = sct.monitors[0]
            self.offset_x = mon["left"]
            self.offset_y = mon["top"]
            w, h = mon["width"], mon["height"]
            
        self.root.geometry(f"{w}x{h}+{self.offset_x}+{self.offset_y}")
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

        if w <= 0 or h <= 0: return None

        hwndDC = win32gui.GetWindowDC(hwnd)
        mfcDC  = win32ui.CreateDCFromHandle(hwndDC)
        saveDC = mfcDC.CreateCompatibleDC()

        saveBitMap = win32ui.CreateBitmap()
        saveBitMap.CreateCompatibleBitmap(mfcDC, w, h)

        saveDC.SelectObject(saveBitMap)

        # 2 = PW_RENDERFULLCONTENT
        result = ctypes.windll.user32.PrintWindow(hwnd, saveDC.GetSafeHdc(), 2)

        frame = None
        if result == 1:
            bmpinfo = saveBitMap.GetInfo()
            bmpstr = saveBitMap.GetBitmapBits(True)
            img = np.frombuffer(bmpstr, dtype=np.uint8).reshape((bmpinfo['bmHeight'], bmpinfo['bmWidth'], 4))
            frame = cv2.cvtColor(img, cv2.COLOR_BGRA2BGR)
            
        win32gui.DeleteObject(saveBitMap.GetHandle())
        saveDC.DeleteDC()
        mfcDC.DeleteDC()
        win32gui.ReleaseDC(hwnd, hwndDC)
        
        if frame is None or not frame.any():
            abs_x1 = left + rel_box['x1']
            abs_y1 = top + rel_box['y1']
            abs_width = rel_box['width']
            abs_height = rel_box['height']
            monitor = {"top": abs_y1, "left": abs_x1, "width": abs_width, "height": abs_height}
            grab = _get_sct().grab(monitor)
            return cv2.cvtColor(np.array(grab), cv2.COLOR_BGRA2BGR)
            
        else:
            x1 = max(0, min(w, rel_box['x1']))
            y1 = max(0, min(h, rel_box['y1']))
            x2 = max(0, min(w, rel_box['x2']))
            y2 = max(0, min(h, rel_box['y2']))
            if x2-x1 <= 0 or y2-y1 <= 0: return None
            return frame[y1:y2, x1:x2].copy()
    except Exception as e:
        print(f"Capture Exception: {e}")
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
    def __init__(self, host, port, axis):
        self.host   = host
        self.port   = port
        self.axis   = axis
        self.volume = 0.5
        self.ws     = None
        self._lock  = threading.Lock()
        # Connect lazily on first use — don't block startup if Restim isn't running

    def connect(self):
        with self._lock:
            try:
                ws_url = f"ws://{self.host}:{self.port}"
                ws = websocket.create_connection(ws_url, timeout=2.0)
                self.ws = ws
                print(f"[Restim] Connected to WebSocket at {ws_url}")
                self.set_volume(self.volume, instant=True)
            except Exception as e:
                print(f"[Restim] WebSocket connection failed: {e}. Ensure WebSocket Server is enabled on port {self.port}")
                self.ws = None

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


def list_audio_devices():
    """Return active render (output) devices as list of pycaw AudioDevice objects."""
    try:
        devices = AudioUtilities.GetAllDevices()
        return [d for d in devices if d.flow == 0 and d.state == 1]
    except Exception as e:
        print(f"[WinAudio] Could not enumerate devices: {e}")
        return []


class WindowsAudioClient:
    def __init__(self, device):
        self._volume_interface = None
        try:
            interface = device._dev.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
            self._volume_interface = interface.QueryInterface(IAudioEndpointVolume)
            print(f"[WinAudio] Connected to: {device.FriendlyName}")
        except Exception as e:
            print(f"[WinAudio] Failed to activate device: {e}")

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
            print(f"[WinAudio] set_volume failed: {e}")

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
        if latest and tuple(int(x) for x in latest.split(".")) > tuple(int(x) for x in VERSION.split(".")):
            on_update_available(latest, url)
    except Exception:
        pass  # silently ignore — no internet, rate limit, etc.


class App:
    # YOLO reanchoring
    _YOLO_INTERVAL = 15    # run detector every N frames
    _YOLO_CONFIRM  = 2     # consecutive detections in same area before reanchoring
    _YOLO_MAX_JUMP = 2.0   # max allowed jump as multiple of current bbox diagonal
    # Tracker plausibility
    _SIZE_RATIO_MAX  = 2.5
    _MAX_JUMP_FACTOR = 2.5

    def __init__(self, hwnd, rel_box, initial_frame, bbox):
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
        self.yolo_frame_counter = 0
        self.yolo_candidate     = None  # (bbox, hits) pending confirmation

        # CV
        self.tracker  = cv2.TrackerCSRT_create()
        self.tracker.init(initial_frame, bbox)
        self.detector = DickDetector()

        # Output clients
        self.restim      = RestimClient(RESTIM_HOST, RESTIM_PORT, TCODE_AXIS)
        self.win_audio   = None
        self.win_devices = []

        # Root window
        self.root = tk.Tk()
        self.root.title("Cock Volume Controller")
        self.root.configure(bg="#222")

        # tkinter vars — must be created after root exists
        self.min_vol_var = tk.DoubleVar(value=0.0)
        self.max_vol_var = tk.DoubleVar(value=100.0)
        self.aggr_var    = tk.IntVar(value=1)
        self.mode_var    = tk.StringVar(value="restim")
        self.port_var    = tk.StringVar(value="12346")

        self._build_ui()
        self._start_update_check()

    def run(self):
        self.root.after(5, self._update_frame)
        self.root.mainloop()

    # ------------------------------------------------------------------ UI

    def _build_ui(self):
        root = self.root

        # Update banner (hidden until an update is found)
        self._update_banner = tk.Frame(root, bg="#b8860b")
        self._update_label  = tk.Label(self._update_banner, text="", bg="#b8860b", fg="white",
                                       font=("Arial", 10, "bold"))
        self._update_label.pack(side=tk.LEFT, padx=10, pady=4)
        self._update_btn = tk.Button(self._update_banner, text="Download", bg="#8B6914", fg="white",
                                     font=("Arial", 10, "bold"), relief=tk.FLAT, cursor="hand2")
        self._update_btn.pack(side=tk.RIGHT, padx=10, pady=4)

        # Video feed + height buttons (stored so the update banner can be inserted before it)
        top_frame = tk.Frame(root, bg="#222")
        top_frame.pack(padx=10, pady=10)
        self._first_widget = top_frame

        self.video_label = tk.Label(top_frame, bg="#222")
        self.video_label.pack(side=tk.LEFT)

        btn_font         = ("Arial", 12, "bold")
        height_btn_frame = tk.Frame(top_frame, bg="#222")
        height_btn_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(10, 0))
        tk.Button(height_btn_frame, text="Set Edging Height",  command=self._set_edging,  bg="#ff9999", font=btn_font).pack(fill=tk.X, pady=(0, 2))
        tk.Frame(height_btn_frame, bg="#222").pack(fill=tk.BOTH, expand=True)
        tk.Button(height_btn_frame, text="Set Erect Height",   command=self._set_erect,   bg="#99ff99", font=btn_font).pack(fill=tk.X, pady=2)
        tk.Frame(height_btn_frame, bg="#222").pack(fill=tk.BOTH, expand=True)
        tk.Button(height_btn_frame, text="Set Flaccid Height", command=self._set_flaccid, bg="#9999ff", font=btn_font).pack(fill=tk.X, pady=(2, 0))

        lbl_font = ("Arial", 10, "bold")

        # Volume floor / ceiling
        vol_frame = tk.Frame(root, bg="#222")
        vol_frame.pack(fill=tk.X, padx=10, pady=5)
        tk.Label(vol_frame, text="Vol Floor (%):",   bg="#222", fg="white", font=lbl_font).pack(side=tk.LEFT)
        tk.Scale(vol_frame, from_=0, to=100, orient=tk.HORIZONTAL, variable=self.min_vol_var,
                 bg="#222", fg="white", highlightthickness=0, length=120).pack(side=tk.LEFT, padx=(2, 10))
        tk.Label(vol_frame, text="Vol Ceiling (%):", bg="#222", fg="white", font=lbl_font).pack(side=tk.LEFT)
        tk.Scale(vol_frame, from_=0, to=100, orient=tk.HORIZONTAL, variable=self.max_vol_var,
                 bg="#222", fg="white", highlightthickness=0, length=120).pack(side=tk.LEFT, padx=(2, 0))

        # Aggressiveness dial
        aggr_frame = tk.Frame(root, bg="#222")
        aggr_frame.pack(fill=tk.X, padx=10, pady=(0, 5))
        tk.Label(aggr_frame, text="Aggressiveness:", bg="#222", fg="white", font=lbl_font).pack(side=tk.LEFT)
        self._aggr_name_label = tk.Label(aggr_frame, text=AGGR_LEVELS[1][0], bg="#222", fg="#ffcc00",
                                         font=("Arial", 10, "bold"), width=7)
        self._aggr_name_label.pack(side=tk.RIGHT, padx=(0, 5))
        tk.Scale(aggr_frame, from_=0, to=3, resolution=1, orient=tk.HORIZONTAL,
                 variable=self.aggr_var, command=self._on_aggr_change,
                 bg="#222", fg="white", highlightthickness=0, showvalue=0, length=180,
                 tickinterval=1).pack(side=tk.LEFT, padx=(8, 4))

        # Mode toggle
        mode_frame = tk.Frame(root, bg="#222")
        mode_frame.pack(fill=tk.X, padx=10, pady=(5, 0))
        tk.Label(mode_frame, text="Output mode:", bg="#222", fg="white", font=lbl_font).pack(side=tk.LEFT, padx=(0, 8))
        tk.Radiobutton(mode_frame, text="Restim",        variable=self.mode_var, value="restim",
                       bg="#222", fg="white", selectcolor="#444", font=lbl_font,
                       command=self._on_mode_change).pack(side=tk.LEFT)
        tk.Radiobutton(mode_frame, text="Windows Audio", variable=self.mode_var, value="windows",
                       bg="#222", fg="white", selectcolor="#444", font=lbl_font,
                       command=self._on_mode_change).pack(side=tk.LEFT, padx=(8, 0))

        # Restim options panel
        self._restim_opts = tk.Frame(root, bg="#222")
        tk.Label(self._restim_opts, text="Port:", bg="#222", fg="white", font=("Arial", 10)).pack(side=tk.LEFT, padx=(10, 2))
        tk.Entry(self._restim_opts, textvariable=self.port_var, width=6).pack(side=tk.LEFT)
        self.port_var.trace_add("write", self._on_port_change)

        # Windows Audio options panel
        self._windows_opts = tk.Frame(root, bg="#222")
        tk.Label(self._windows_opts, text="Device:", bg="#222", fg="white", font=("Arial", 10)).pack(side=tk.LEFT, padx=(10, 2))
        self._device_var   = tk.StringVar()
        self._device_combo = ttk.Combobox(self._windows_opts, textvariable=self._device_var,
                                          state="readonly", width=35)
        self._device_combo.pack(side=tk.LEFT, padx=(0, 5))
        self._device_combo.bind("<<ComboboxSelected>>", self._on_device_select)
        tk.Button(self._windows_opts, text="Refresh", command=self._refresh_devices,
                  bg="#444", fg="white", font=("Arial", 9)).pack(side=tk.LEFT)

        # Show default mode panel
        self._restim_opts.pack(fill=tk.X, padx=10, pady=(0, 5))

        # Re-select buttons
        reselect_frame = tk.Frame(root, bg="#222")
        reselect_frame.pack(fill=tk.X, padx=10, pady=5)
        tk.Button(reselect_frame, text="Re-Select Video Feed Area", command=self._reselect_feed,
                  bg="#555", fg="white", font=("Arial", 10)).pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5)
        tk.Button(reselect_frame, text="Re-Select Cock Head", command=self._reselect_head,
                  bg="#444", fg="white", font=("Arial", 10)).pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5)

        self.info_label = tk.Label(root, text="State: Erect | Vol: 50% | WS: Disconnected",
                                   font=("Arial", 14), bg="#222", fg="white")
        self.info_label.pack(pady=10)

    def _start_update_check(self):
        def callback(latest, url):
            self.root.after(0, self._show_update_banner, latest, url)
        threading.Thread(target=check_for_update, args=(callback,), daemon=True).start()

    # ------------------------------------------------------------------ callbacks

    def _show_update_banner(self, latest, url):
        self._update_label.config(text=f"Update available: v{latest}")
        self._update_btn.config(command=lambda: webbrowser.open(url))
        self._update_banner.pack(fill=tk.X, before=self._first_widget)

    def _set_edging(self):
        self.heights["Edging"] = self.head_y
        print(f"Edging height set at Y: {self.head_y}")

    def _set_erect(self):
        self.heights["Erect"] = self.head_y
        print(f"Erect height set at Y: {self.head_y}")

    def _set_flaccid(self):
        self.heights["Flaccid"] = self.head_y
        print(f"Flaccid height set at Y: {self.head_y}")

    def _reselect_feed(self):
        self.tracking_paused = True
        new_hwnd, new_rel_box = select_region(self.root)
        if new_hwnd and new_rel_box['width'] > 10 and new_rel_box['height'] > 10:
            self.hwnd    = new_hwnd
            self.rel_box = new_rel_box
            self._reselect_head()
        else:
            self.tracking_paused = False

    def _reselect_head(self):
        self.tracking_paused = True
        pause_frame = capture_window_region(self.hwnd, self.rel_box)
        if pause_frame is not None:
            new_bbox = select_head(pause_frame, parent=self.root)
            if new_bbox[2] > 0 and new_bbox[3] > 0:
                self.tracker.init(pause_frame, new_bbox)
                self.last_bbox   = new_bbox
                self.head_y      = new_bbox[1] + new_bbox[3] // 2
                self.tracking_ok = True
        self.tracking_paused = False

    def _on_aggr_change(self, *_):
        self._aggr_name_label.config(text=AGGR_LEVELS[self.aggr_var.get()][0])

    def _on_mode_change(self):
        if self.mode_var.get() == "restim":
            self._windows_opts.pack_forget()
            self._restim_opts.pack(fill=tk.X, padx=10, pady=(0, 5))
        else:
            self._restim_opts.pack_forget()
            self._windows_opts.pack(fill=tk.X, padx=10, pady=(0, 5))
            if not self.win_devices:
                self._refresh_devices()

    def _on_port_change(self, *_):
        val = self.port_var.get().strip()
        if val.isdigit():
            new_port = int(val)
            if new_port != self.restim.port:
                self.restim.port = new_port
                if self.restim.ws:
                    try:
                        self.restim.ws.close()
                    except Exception:
                        pass
                    self.restim.ws = None

    def _refresh_devices(self):
        self.win_devices = list_audio_devices()
        self._device_combo["values"] = [d.FriendlyName for d in self.win_devices]
        if self.win_devices:
            self._device_combo.current(0)
            self._on_device_select(None)

    def _on_device_select(self, _event):
        idx = self._device_combo.current()
        if 0 <= idx < len(self.win_devices):
            self.win_audio = WindowsAudioClient(self.win_devices[idx])

    # ------------------------------------------------------------------ frame loop

    def _update_frame(self):
        if self.tracking_paused:
            self.root.after(100, self._update_frame)
            return

        frame = capture_window_region(self.hwnd, self.rel_box)
        if frame is None:
            self.info_label.config(text="State: WINDOW HIDDEN? | Vol: -- | WS: --", fg="yellow")
            self.root.after(200, self._update_frame)
            return

        self._maybe_yolo_reanchor(frame)
        self._run_tracker(frame)

        state = self._determine_state(self.head_y)

        self._draw_height_lines(frame)
        self._tick_volume()
        self._update_status_label(state)
        self._display_frame(frame)

        self.root.after(5, self._update_frame)

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
                self.last_bbox   = (x, y, w, h_box)
                self.tracking_ok = True
                self.head_y      = new_cy
                cv2.rectangle(frame, (x, y), (x + w, y + h_box), (0, 255, 0), 2)
                cv2.circle(frame, (new_cx, new_cy), 4, (0, 0, 255), -1)
            else:
                reason = "size" if not size_ok else "jump"
                cv2.putText(frame, f"TRACKING SUSPECT ({reason}) - Frozen", (10, 30),
                            cv2.FONT_HERSHEY_SIMPLEX, 0.6, (0, 165, 255), 2)
                px, py, pw, ph = self.last_bbox
                cv2.rectangle(frame, (px, py), (px + pw, py + ph), (0, 165, 255), 2)
                self.tracking_ok = False
                self.tracker.init(frame, self.last_bbox)
        else:
            cv2.putText(frame, "TRACKING LOST - Frozen at last position", (10, 30),
                        cv2.FONT_HERSHEY_SIMPLEX, 0.6, (0, 0, 255), 2)
            px, py, pw, ph = self.last_bbox
            cv2.rectangle(frame, (px, py), (px + pw, py + ph), (0, 0, 255), 2)
            self.tracking_ok = False
            self.tracker.init(frame, self.last_bbox)

    def _determine_state(self, y_pos):
        if any(v is None for v in self.heights.values()):
            return "Erect (Needs Calibration)"
        dist_edging  = abs(y_pos - self.heights["Edging"])
        dist_erect   = abs(y_pos - self.heights["Erect"])
        dist_flaccid = abs(y_pos - self.heights["Flaccid"])
        minimum = min(dist_edging, dist_erect, dist_flaccid)
        if minimum == dist_edging:  return "Edging"
        if minimum == dist_flaccid: return "Flaccid"
        return "Erect"

    def _draw_height_lines(self, frame):
        fw = frame.shape[1]
        if self.heights["Edging"]  is not None: cv2.line(frame, (0, self.heights["Edging"]),  (fw, self.heights["Edging"]),  (0, 0, 255), 2)
        if self.heights["Erect"]   is not None: cv2.line(frame, (0, self.heights["Erect"]),   (fw, self.heights["Erect"]),   (0, 255, 0), 2)
        if self.heights["Flaccid"] is not None: cv2.line(frame, (0, self.heights["Flaccid"]), (fw, self.heights["Flaccid"]), (255, 0, 0), 2)

    def _compute_volume_delta(self):
        if any(v is None for v in self.heights.values()):
            return 0.0
        full_range = self.heights["Flaccid"] - self.heights["Edging"]
        if abs(full_range) < 1:
            return 0.0
        position   = (self.head_y            - self.heights["Edging"]) / full_range
        erect_norm = (self.heights["Erect"]   - self.heights["Edging"]) / full_range
        _, aggr_mult = AGGR_LEVELS[self.aggr_var.get()]

        if 0.0 <= position <= erect_norm:
            return 0.0
        if position < 0.0:
            return -VOLUME_STEP * 0.5 * aggr_mult
        dist = (position - erect_norm) / max(1.0 - erect_norm, 0.01)
        return VOLUME_STEP * min(dist, 1.0) * aggr_mult

    def _tick_volume(self):
        cur_time = time.time()
        if cur_time - self.last_vol_time < VOLUME_UPDATE_INTERVAL:
            return
        self.last_vol_time = cur_time

        floor_val = min(self.min_vol_var.get(), self.max_vol_var.get()) / 100.0
        ceil_val  = self.max_vol_var.get() / 100.0
        delta     = self._compute_volume_delta()
        mode      = self.mode_var.get()

        if mode == "restim":
            if delta != 0.0:
                self.restim.adjust_volume(delta, floor=floor_val, ceiling=ceil_val)
            if self.restim.ws is None:
                threading.Thread(target=self.restim.connect, daemon=True).start()
        elif mode == "windows" and self.win_audio and self.win_audio.connected and delta != 0.0:
            self.win_audio.adjust_volume(delta, floor=floor_val, ceiling=ceil_val)

    def _update_status_label(self, state):
        quality_str = "Track: OK" if self.tracking_ok else "Track: LOST"
        mode = self.mode_var.get()
        if mode == "restim":
            conn_color = "#00ff00" if self.restim.ws else "#ff0000"
            vol_str    = f"{self.restim.volume * 100:.0f}%"
            src_str    = f"WS: {'Connected' if self.restim.ws else 'Disconnected'}"
        elif self.win_audio and self.win_audio.connected:
            conn_color = "#00ff00"
            vol_str    = f"{self.win_audio.get_volume() * 100:.0f}%"
            src_str    = "Win Audio: OK"
        else:
            conn_color = "#ffaa00"
            vol_str    = "--"
            src_str    = "Win Audio: No Device"
        self.info_label.config(
            text=f"State: {state}  |  Vol: {vol_str}  |  {quality_str}  |  {src_str}",
            fg=conn_color,
        )

    def _display_frame(self, frame):
        img   = Image.fromarray(cv2.cvtColor(frame, cv2.COLOR_BGR2RGB))
        imgtk = ImageTk.PhotoImage(image=img)
        self.video_label.imgtk = imgtk  # prevent GC
        self.video_label.configure(image=imgtk)


def main():
    print("Cock Volume Controller starting...")

    hwnd, rel_box = select_region()
    if not hwnd or rel_box['width'] <= 10 or rel_box['height'] <= 10:
        print("Invalid region selected. Exiting.")
        return

    initial_frame = capture_window_region(hwnd, rel_box)
    if initial_frame is None:
        print("Failed to capture window. Ensure it is not fully minimized.")
        return

    bbox = select_head(initial_frame)
    if bbox[2] == 0 or bbox[3] == 0:
        print("No head selected. Exiting.")
        return

    App(hwnd, rel_box, initial_frame, bbox).run()


if __name__ == "__main__":
    main()
