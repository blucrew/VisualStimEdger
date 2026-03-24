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

sct_global = mss()


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

# YOLO reanchoring tuning
YOLO_INTERVAL = 15     # run detector every N frames
YOLO_CONFIRM  = 2      # consecutive detections in same area before reanchoring
YOLO_MAX_JUMP = 2.0    # max allowed jump as multiple of current bbox diagonal

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
            grab = sct_global.grab(monitor)
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
        self.host = host
        self.port = port
        self.axis = axis
        self.ws = None
        self.volume = 0.5 
        self.connect()

    def connect(self):
        try:
            ws_url = f"ws://{self.host}:{self.port}"
            self.ws = websocket.create_connection(ws_url, timeout=2.0)
            print(f"[Restim] Connected to WebSocket at {ws_url}")
            self.set_volume(self.volume, instant=True)
        except Exception as e:
            print(f"[Restim] WebSocket connection failed: {e}. Ensure WebSocket Server is enabled on port {self.port}")
            self.ws = None

    def set_volume(self, vol, instant=False, floor=0.0, ceiling=1.0):
        self.volume = max(floor, min(ceiling, vol))
        if self.ws:
            try:
                val_int = int(round(self.volume * 9999))
                interval = 0 if instant else int(VOLUME_UPDATE_INTERVAL * 1000)
                cmd = f"{self.axis}{val_int:04d}I{interval}"
                self.ws.send(cmd)
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


def main():
    print("Cock Volume Controller starting...")
    
    hwnd, rel_box = select_region()
    if not hwnd or rel_box['width'] <= 10 or rel_box['height'] <= 10:
        print("Invalid region selected. Exiting.")
        return

    initial_frame = capture_window_region(hwnd, rel_box)
    if initial_frame is None:
        print("Failed to capture parent application window. Ensure it is not fully minimized.")
        return
        
    bbox = select_head(initial_frame)
    if bbox[2] == 0 or bbox[3] == 0:
        print("No head selected. Exiting.")
        return

    tracker = cv2.TrackerCSRT_create()
    tracker.init(initial_frame, bbox)

    detector = DickDetector()

    restim = RestimClient(RESTIM_HOST, RESTIM_PORT, TCODE_AXIS)

    root = tk.Tk()
    root.title("Cock Volume Controller")
    root.configure(bg="#222")

    # --- Update check ---
    update_banner = tk.Frame(root, bg="#b8860b")
    update_label  = tk.Label(update_banner, text="", bg="#b8860b", fg="white", font=("Arial", 10, "bold"))
    update_label.pack(side=tk.LEFT, padx=10, pady=4)
    update_btn = tk.Button(update_banner, text="Download", bg="#8B6914", fg="white",
                           font=("Arial", 10, "bold"), relief=tk.FLAT, cursor="hand2")
    update_btn.pack(side=tk.RIGHT, padx=10, pady=4)

    def _show_update_banner(latest, url):
        update_label.config(text=f"Update available: v{latest}")
        update_btn.config(command=lambda: webbrowser.open(url))
        update_banner.pack(fill=tk.X, before=root.winfo_children()[0])

    threading.Thread(
        target=check_for_update,
        args=(lambda latest, url: root.after(0, _show_update_banner, latest, url),),
        daemon=True
    ).start()

    app_state = {
        "hwnd": hwnd,
        "rel_box": rel_box,
        "head_y": bbox[1] + bbox[3]//2,
        "heights": {"Edging": None, "Erect": None, "Flaccid": None},
        "current_state": "Erect",
        "last_vol_time": time.time(),
        "tracking_paused": False,
        "min_vol_var": tk.DoubleVar(value=0.0),
        "last_bbox": tuple(int(v) for v in bbox),
        "tracking_quality": 1.0,
        "yolo_frame_counter": 0,
        "yolo_candidate": None,   # (bbox, hits) — pending confirmation before reanchor
        "mode_var": tk.StringVar(value="restim"),
        "win_audio": None,
        "win_devices": [],
        "max_vol_var": tk.DoubleVar(value=100.0),
        "aggr_var": tk.IntVar(value=1),
    }
    
    # --- Video + height buttons side by side ---
    top_frame = tk.Frame(root, bg="#222")
    top_frame.pack(padx=10, pady=10)

    video_label = tk.Label(top_frame, bg="#222")
    video_label.pack(side=tk.LEFT)

    def set_edging():
        app_state["heights"]["Edging"] = app_state["head_y"]
        print(f"Edging height set at Y: {app_state['head_y']}")
    def set_erect():
        app_state["heights"]["Erect"] = app_state["head_y"]
        print(f"Erect height set at Y: {app_state['head_y']}")
    def set_flaccid():
        app_state["heights"]["Flaccid"] = app_state["head_y"]
        print(f"Flaccid height set at Y: {app_state['head_y']}")

    btn_font = ("Arial", 12, "bold")
    height_btn_frame = tk.Frame(top_frame, bg="#222")
    height_btn_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(10, 0))
    tk.Button(height_btn_frame, text="Set Edging Height", command=set_edging, bg="#ff9999", font=btn_font).pack(fill=tk.X, pady=(0, 2))
    tk.Frame(height_btn_frame, bg="#222").pack(fill=tk.BOTH, expand=True)
    tk.Button(height_btn_frame, text="Set Erect Height", command=set_erect, bg="#99ff99", font=btn_font).pack(fill=tk.X, pady=2)
    tk.Frame(height_btn_frame, bg="#222").pack(fill=tk.BOTH, expand=True)
    tk.Button(height_btn_frame, text="Set Flaccid Height", command=set_flaccid, bg="#9999ff", font=btn_font).pack(fill=tk.X, pady=(2, 0))

    def reselect_feed():
        app_state["tracking_paused"] = True
        new_hwnd, new_rel_box = select_region(root)
        if new_hwnd and new_rel_box['width'] > 10 and new_rel_box['height'] > 10:
            app_state["hwnd"] = new_hwnd
            app_state["rel_box"] = new_rel_box
            reselect_head()
        else:
            app_state["tracking_paused"] = False

    def reselect_head():
        app_state["tracking_paused"] = True
        pause_frame = capture_window_region(app_state["hwnd"], app_state["rel_box"])
        if pause_frame is not None:
            new_bbox = select_head(pause_frame, parent=root)
            if new_bbox[2] > 0 and new_bbox[3] > 0:
                tracker.init(pause_frame, new_bbox)
                app_state["last_bbox"] = new_bbox
                app_state["head_y"] = new_bbox[1] + new_bbox[3]//2
                app_state["tracking_quality"] = 1.0
        app_state["tracking_paused"] = False
    
    # --- Volume floor / ceiling ---
    vol_range_frame = tk.Frame(root, bg="#222")
    vol_range_frame.pack(fill=tk.X, padx=10, pady=5)
    lbl_font = ("Arial", 10, "bold")
    tk.Label(vol_range_frame, text="Vol Floor (%):", bg="#222", fg="white", font=lbl_font).pack(side=tk.LEFT)
    tk.Scale(vol_range_frame, from_=0, to=100, orient=tk.HORIZONTAL, variable=app_state["min_vol_var"], bg="#222", fg="white", highlightthickness=0, length=120).pack(side=tk.LEFT, padx=(2, 10))
    tk.Label(vol_range_frame, text="Vol Ceiling (%):", bg="#222", fg="white", font=lbl_font).pack(side=tk.LEFT)
    tk.Scale(vol_range_frame, from_=0, to=100, orient=tk.HORIZONTAL, variable=app_state["max_vol_var"], bg="#222", fg="white", highlightthickness=0, length=120).pack(side=tk.LEFT, padx=(2, 0))

    # --- Aggressiveness dial ---
    aggr_frame = tk.Frame(root, bg="#222")
    aggr_frame.pack(fill=tk.X, padx=10, pady=(0, 5))
    tk.Label(aggr_frame, text="Aggressiveness:", bg="#222", fg="white", font=lbl_font).pack(side=tk.LEFT)
    aggr_name_label = tk.Label(aggr_frame, text=AGGR_LEVELS[1][0], bg="#222", fg="#ffcc00", font=("Arial", 10, "bold"), width=7)
    aggr_name_label.pack(side=tk.RIGHT, padx=(0, 5))
    def _on_aggr_change(*_):
        aggr_name_label.config(text=AGGR_LEVELS[app_state["aggr_var"].get()][0])
    tk.Scale(aggr_frame, from_=0, to=3, resolution=1, orient=tk.HORIZONTAL,
             variable=app_state["aggr_var"], command=_on_aggr_change,
             bg="#222", fg="white", highlightthickness=0, showvalue=0, length=180,
             tickinterval=1).pack(side=tk.LEFT, padx=(8, 4))

    # --- Mode toggle ---
    mode_frame = tk.Frame(root, bg="#222")
    mode_frame.pack(fill=tk.X, padx=10, pady=(5, 0))
    tk.Label(mode_frame, text="Output mode:", bg="#222", fg="white", font=lbl_font).pack(side=tk.LEFT, padx=(0, 8))
    tk.Radiobutton(mode_frame, text="Restim", variable=app_state["mode_var"], value="restim",
                   bg="#222", fg="white", selectcolor="#444", font=lbl_font,
                   command=_on_mode_change).pack(side=tk.LEFT)
    tk.Radiobutton(mode_frame, text="Windows Audio", variable=app_state["mode_var"], value="windows",
                   bg="#222", fg="white", selectcolor="#444", font=lbl_font,
                   command=_on_mode_change).pack(side=tk.LEFT, padx=(8, 0))

    # --- Restim options ---
    restim_opts = tk.Frame(root, bg="#222")
    tk.Label(restim_opts, text="Port:", bg="#222", fg="white", font=("Arial", 10)).pack(side=tk.LEFT, padx=(10, 2))
    app_state["port_var"] = tk.StringVar(value="12346")
    tk.Entry(restim_opts, textvariable=app_state["port_var"], width=6).pack(side=tk.LEFT)

    def on_port_change(*args):
        val = app_state["port_var"].get().strip()
        if val.isdigit():
            new_port = int(val)
            if new_port != restim.port:
                restim.port = new_port
                if restim.ws:
                    try: restim.ws.close()
                    except: pass
                    restim.ws = None
    app_state["port_var"].trace_add("write", on_port_change)

    # --- Windows Audio options ---
    windows_opts = tk.Frame(root, bg="#222")
    tk.Label(windows_opts, text="Device:", bg="#222", fg="white", font=("Arial", 10)).pack(side=tk.LEFT, padx=(10, 2))
    device_var = tk.StringVar()
    device_combo = ttk.Combobox(windows_opts, textvariable=device_var, state="readonly", width=35)
    device_combo.pack(side=tk.LEFT, padx=(0, 5))

    def refresh_devices():
        app_state["win_devices"] = list_audio_devices()
        device_combo["values"] = [d.FriendlyName for d in app_state["win_devices"]]
        if app_state["win_devices"]:
            device_combo.current(0)
            _on_device_select(None)

    def _on_device_select(event):
        idx = device_combo.current()
        if 0 <= idx < len(app_state["win_devices"]):
            app_state["win_audio"] = WindowsAudioClient(app_state["win_devices"][idx])

    device_combo.bind("<<ComboboxSelected>>", _on_device_select)
    tk.Button(windows_opts, text="Refresh", command=refresh_devices, bg="#444", fg="white", font=("Arial", 9)).pack(side=tk.LEFT)

    def _on_mode_change():
        if app_state["mode_var"].get() == "restim":
            windows_opts.pack_forget()
            restim_opts.pack(fill=tk.X, padx=10, pady=(0, 5))
        else:
            restim_opts.pack_forget()
            windows_opts.pack(fill=tk.X, padx=10, pady=(0, 5))
            if not app_state["win_devices"]:
                refresh_devices()

    # Show default (restim) options
    restim_opts.pack(fill=tk.X, padx=10, pady=(0, 5))
    
    reselect_frame = tk.Frame(root, bg="#222")
    reselect_frame.pack(fill=tk.X, padx=10, pady=5)
    tk.Button(reselect_frame, text="Re-Select Video Feed Area", command=reselect_feed, bg="#555", fg="white", font=("Arial", 10)).pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5)
    tk.Button(reselect_frame, text="Re-Select Cock Head", command=reselect_head, bg="#444", fg="white", font=("Arial", 10)).pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5)
    
    info_label = tk.Label(root, text="State: Erect | Vol: 50% | WS: Disconnected", font=("Arial", 14), bg="#222", fg="white")
    info_label.pack(pady=10)

    def determine_state(y_pos):
        heights = app_state["heights"]
        if heights["Edging"] is None or heights["Erect"] is None or heights["Flaccid"] is None:
            return "Erect (Needs Calibration)"

        dist_edging  = abs(y_pos - heights["Edging"])
        dist_erect   = abs(y_pos - heights["Erect"])
        dist_flaccid = abs(y_pos - heights["Flaccid"])

        minimum = min(dist_edging, dist_erect, dist_flaccid)
        if minimum == dist_edging:  return "Edging"
        elif minimum == dist_flaccid: return "Flaccid"
        else: return "Erect"

    def update_frame():
        if app_state["tracking_paused"]:
            root.after(100, update_frame)
            return
            
        frame = capture_window_region(app_state["hwnd"], app_state["rel_box"])
        if frame is None:
            info_label.config(text="State: WINDOW HIDDEN? | Vol: -- | WS: --", fg="yellow")
            root.after(200, update_frame)
            return
            
        # --- YOLO reanchor ---
        app_state["yolo_frame_counter"] += 1
        if detector.available and app_state["yolo_frame_counter"] >= YOLO_INTERVAL:
            app_state["yolo_frame_counter"] = 0
            yolo_bbox = detector.detect_head(frame)

            if yolo_bbox is not None:
                # Distance gate — reject if detection is too far from current position
                px, py, pw, ph = app_state["last_bbox"]
                cur_cx, cur_cy = px + pw // 2, py + ph // 2
                yx, yy, yw, yh = yolo_bbox
                det_cx, det_cy = yx + yw // 2, yy + yh // 2
                diag = np.sqrt(pw ** 2 + ph ** 2)
                dist = np.sqrt((det_cx - cur_cx) ** 2 + (det_cy - cur_cy) ** 2)

                if dist <= diag * YOLO_MAX_JUMP:
                    # Accumulate confirmation hits
                    cand = app_state["yolo_candidate"]
                    if cand is not None:
                        prev_bbox, hits = cand
                        # Check the new detection is close to the previous candidate
                        pcx = prev_bbox[0] + prev_bbox[2] // 2
                        pcy = prev_bbox[1] + prev_bbox[3] // 2
                        if np.sqrt((det_cx - pcx) ** 2 + (det_cy - pcy) ** 2) <= diag:
                            hits += 1
                        else:
                            hits = 1  # new candidate, reset
                    else:
                        hits = 1

                    app_state["yolo_candidate"] = (yolo_bbox, hits)

                    if hits >= YOLO_CONFIRM:
                        tracker.init(frame, yolo_bbox)
                        app_state["last_bbox"] = yolo_bbox
                        app_state["yolo_candidate"] = None
                        x, y, w, h = yolo_bbox
                        cv2.rectangle(frame, (x, y), (x + w, y + h), (255, 255, 0), 1)
                else:
                    # Too far — could be balls or hand; discard and reset candidate
                    app_state["yolo_candidate"] = None
            else:
                app_state["yolo_candidate"] = None

        success, new_bbox = tracker.update(frame)

        SIZE_RATIO_MAX  = 2.5    # max allowed bbox size change factor per frame
        MAX_JUMP_FACTOR = 2.5    # max allowed center jump relative to bbox diagonal

        if success:
            x, y, w, h_box = [int(v) for v in new_bbox]
            px, py, pw, ph = app_state["last_bbox"]

            # --- Plausibility: size change ---
            size_ok = (
                0 < w <= frame.shape[1] and 0 < h_box <= frame.shape[0] and
                (1/SIZE_RATIO_MAX) < (w / max(pw, 1)) < SIZE_RATIO_MAX and
                (1/SIZE_RATIO_MAX) < (h_box / max(ph, 1)) < SIZE_RATIO_MAX
            )

            # --- Plausibility: position jump ---
            prev_cx, prev_cy = px + pw//2, py + ph//2
            new_cx, new_cy = x + w//2, y + h_box//2
            diag = np.sqrt(pw**2 + ph**2)
            jump = np.sqrt((new_cx - prev_cx)**2 + (new_cy - prev_cy)**2)
            jump_ok = jump < diag * MAX_JUMP_FACTOR

            if size_ok and jump_ok:
                app_state["last_bbox"] = (x, y, w, h_box)
                app_state["tracking_quality"] = 1.0
                cv2.rectangle(frame, (x, y), (x + w, y + h_box), (0, 255, 0), 2)
                cv2.circle(frame, (new_cx, new_cy), 4, (0, 0, 255), -1)
                app_state["head_y"] = new_cy
            else:
                # Plausibility failed — freeze at last known position, warn user
                reason = "size" if not size_ok else "jump"
                cv2.putText(frame, f"TRACKING SUSPECT ({reason}) - Frozen", (10, 30),
                            cv2.FONT_HERSHEY_SIMPLEX, 0.6, (0, 165, 255), 2)
                px, py, pw, ph = app_state["last_bbox"]
                cv2.rectangle(frame, (px, py), (px + pw, py + ph), (0, 165, 255), 2)
                app_state["tracking_quality"] = 0.0
                # Reinit tracker to last good bbox so it can recover from that position
                tracker.init(frame, app_state["last_bbox"])
        else:
            cv2.putText(frame, "TRACKING LOST - Frozen at last position", (10, 30),
                        cv2.FONT_HERSHEY_SIMPLEX, 0.6, (0, 0, 255), 2)
            px, py, pw, ph = app_state["last_bbox"]
            cv2.rectangle(frame, (px, py), (px + pw, py + ph), (0, 0, 255), 2)
            app_state["tracking_quality"] = 0.0
            # Reinit at last known good location (not the center)
            tracker.init(frame, app_state["last_bbox"])

        state = determine_state(app_state["head_y"])
        app_state["current_state"] = state

        heights = app_state["heights"]
        if heights["Edging"] is not None:  cv2.line(frame, (0, heights["Edging"]),  (frame.shape[1], heights["Edging"]),  (0, 0, 255), 2)
        if heights["Erect"] is not None:   cv2.line(frame, (0, heights["Erect"]),   (frame.shape[1], heights["Erect"]),   (0, 255, 0), 2)
        if heights["Flaccid"] is not None: cv2.line(frame, (0, heights["Flaccid"]), (frame.shape[1], heights["Flaccid"]), (255, 0, 0), 2)

        cur_time = time.time()
        if cur_time - app_state["last_vol_time"] > VOLUME_UPDATE_INTERVAL:
            floor_val = app_state["min_vol_var"].get() / 100.0
            ceil_val  = app_state["max_vol_var"].get() / 100.0
            floor_val = min(floor_val, ceil_val)
            mode = app_state["mode_var"].get()

            delta = 0.0
            if all(v is not None for v in heights.values()):
                # Normalise head position along the flaccid→edging axis.
                # position = 0.0  → head exactly at edging height
                # position = 1.0  → head exactly at flaccid height
                # erect_norm      → where erect sits in that 0-1 range
                full_range = heights["Flaccid"] - heights["Edging"]
                if abs(full_range) >= 1:
                    position   = (app_state["head_y"] - heights["Edging"]) / full_range
                    erect_norm = (heights["Erect"]     - heights["Edging"]) / full_range

                    _, aggr_mult = AGGR_LEVELS[app_state["aggr_var"].get()]

                    if 0.0 <= position <= erect_norm:
                        # Head is in the sweet zone (edging ↔ erect): hold volume
                        delta = 0.0
                    elif position < 0.0:
                        # Past edging threshold — ease off
                        delta = -VOLUME_STEP * 0.5 * aggr_mult
                    else:
                        # Past erect, drifting toward flaccid — nudge up proportionally
                        dist = (position - erect_norm) / max(1.0 - erect_norm, 0.01)
                        delta = VOLUME_STEP * min(dist, 1.0) * aggr_mult

            if mode == "restim":
                if delta != 0.0:
                    restim.adjust_volume(delta, floor=floor_val, ceiling=ceil_val)
                if restim.ws is None:
                    threading.Thread(target=restim.connect, daemon=True).start()
            elif mode == "windows":
                win_audio = app_state["win_audio"]
                if win_audio and win_audio.connected and delta != 0.0:
                    win_audio.adjust_volume(delta, floor=floor_val, ceiling=ceil_val)

            app_state["last_vol_time"] = cur_time

        quality = app_state["tracking_quality"]
        quality_str = f"Track: {quality*100:.0f}%"
        mode = app_state["mode_var"].get()
        if mode == "restim":
            conn_str = "Connected" if restim.ws else "Disconnected"
            conn_color = "#00ff00" if restim.ws else "#ff0000"
            vol_str = f"{restim.volume*100:.0f}%"
            src_str = f"WS: {conn_str}"
        else:
            win_audio = app_state["win_audio"]
            if win_audio and win_audio.connected:
                conn_color = "#00ff00"
                vol_str = f"{win_audio.get_volume()*100:.0f}%"
                src_str = "Win Audio: OK"
            else:
                conn_color = "#ffaa00"
                vol_str = "--"
                src_str = "Win Audio: No Device"
        info_label.config(
            text=f"State: {state}  |  Vol: {vol_str}  |  {quality_str}  |  {src_str}",
            fg=conn_color
        )

        frame_rgb = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
        img = Image.fromarray(frame_rgb)
        imgtk = ImageTk.PhotoImage(image=img)
        video_label.imgtk = imgtk 
        video_label.configure(image=imgtk)
        
        root.after(5, update_frame)

    root.after(5, update_frame)
    root.mainloop()

if __name__ == "__main__":
    main()
