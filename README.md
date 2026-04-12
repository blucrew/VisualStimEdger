# VisualStimEdger

> **⚡ This is a companion app — it does not generate e-stim signals or audio on its own.**
> You need an existing setup already producing stimulus: [Restim](https://restim.io), xToys, an `.mp3` file playing through a stereo box, or anything else you were already using. VisualStimEdger simply controls the **volume/intensity** of that existing signal based on what it sees on camera.

Tracks your cock on a camera feed and uses its position — flaccid, erect, or edging — to automatically control stimulus intensity. The goal is to keep you in the erect-to-edging zone by easing off stimulus when you get too close to the edge, and ramping it back up when you start to drop.

Output modes: **Restim**, **xToys** (any Lovense/Kiiroo/etc. toy via the xToys browser app), **Windows Audio**, and a built-in **MP3 Player**.

---

## Requirements

- Python 3.9+
- A camera feed showing your cock (OBS, a webcam viewer, or any window with a live feed)
- Restim running with WebSocket server enabled, **or** any Windows audio output device
- Electrodes wired up and ready
- A dick that goes erect when you're close to cumming - this works visually

Install dependencies:

```bash
pip install -r requirements.txt
```

---

## Setup

### 1. Wire up and get on camera

Get your electrodes attached and your cock in frame on whatever camera feed you're using. The feed just needs to be visible in a window on your screen — OBS, a browser stream, a webcam app, anything works.

### 2. Run the program

```bash
python main.py
```

### 3. Choose your output

In the **Output mode** section, select:

- **Restim** — enter the WebSocket port (default `12346`). Make sure Restim's WebSocket server is enabled under its settings.
- **xToys** — drives any toy connected in the [xToys browser app](https://xtoys.app). See xToys setup below.
- **Windows Audio** — click **Refresh** to list your output devices, then pick the one you want from the dropdown.
- **MP3 Player** — built-in player for stim audio files. Load a file and let the tracker drive the volume.

Set your **Vol Floor** and **Vol Ceiling** sliders to define the range the program is allowed to operate within.

#### xToys setup

1. Open [xtoys.app](https://xtoys.app) in a browser and sign in
2. Go to **Scripts** → search for **VisualStimEdger** → **Load Script**
3. In the script's **Controls** view, open the **Connections** panel
4. Under **LOCAL WEBHOOK**, add your toy (Lovense, etc.) — **not** Generic Output
5. Click **Save**, then click the **⚡ satellite icon** on the Local Webhook card → **Connect**
6. Copy the **Webhook ID** shown (short string, e.g. `8hR5acKTCx2s`)
7. Paste it into the **Webhook ID** field in VSE — the status bar will show `xToys: OK`

> **Note:** The Webhook ID changes every time you reconnect in xToys. If your toy stops responding, re-copy it and paste it into VSE again. Keep the xToys browser tab open while using VSE.

### 4. Draw a box around your video feed

A fullscreen overlay will appear. Click and drag to draw a rectangle around the window or region showing your camera feed, then release to lock it in.

### 5. Select the tracking point

A snapshot of your feed will open. Draw a box around what you want to track. Two good options:

- **The head of your cock** — works well if it's clearly visible and has decent contrast against the background.
- **An electrode or ring on the head** — usually better. The electrode tends to be a distinct colour or shape that the tracker can lock onto more reliably, especially as arousal changes the appearance of the skin.

Click **Confirm** when done.

### 6. Calibrate heights

With the program running and tracking, move through your states and click the buttons to record each position:

| Button | When to press |
|---|---|
| **Set Flaccid Height** | When your cock is fully soft |
| **Set Erect Height** | When you're fully hard but comfortable |
| **Set Edging Height** | When you're at the edge |

The program will then keep volume in a hold zone between erect and edging, reduce stimulus slowly if you push past the edging threshold, and increase it proportionally if you start to drop back toward flaccid.

---

## Tips

- **Lighting matters** — consistent, even lighting gives the tracker the best chance of staying locked. Avoid backlighting.
- **Electrode tracking is more reliable** — skin tone and shape changes significantly with arousal; a silicone or metal electrode doesn't. If tracking keeps drifting, try selecting the electrode instead.
- **If tracking is lost or jumps** — the tracker freezes at the last known position (shown in orange) rather than snapping to something random. Use **Re-Select Cock Head** to relock onto the target.
- **Re-Select Video Feed Area** — use this if you move or resize the source window.
- **Vol Floor** — useful to make sure stimulus never cuts out completely. Set to 20–30% if you want a baseline.
- **Vol Ceiling** — caps how high the program will push volume. Set conservatively until you know how the session feels.

---

## Troubleshooting

**Restim shows Disconnected**
Make sure the WebSocket server is enabled in Restim settings and the port matches.

**Tracking keeps jumping to my hand**
The tracker uses size and position plausibility filters to reject sudden jumps, but a hand close to the tracked area can still confuse it. Try to keep hands out of frame, or relock onto an electrode which has more distinct visual features.

**No audio devices showing**
Click **Refresh**. If still empty, check that pycaw installed correctly (`pip install pycaw`).
On Windows 11, audio device enumeration can fail due to permissions — try running the app as Administrator once to confirm this is the issue.
