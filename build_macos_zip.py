"""
Build the macOS source distribution zip.

Regenerates the patched macOS source from VSE.py, then packages it with a
double-clickable launcher (VisualStimEdger.command) that creates an isolated
venv, installs deps, and runs the app — so end users never touch pip or main.py.

Usage:  python build_macos_zip.py
Output: dist/VisualStimEdger-<VERSION>-macOS-ARM64-source.zip
"""
import re, zipfile, pathlib, subprocess, sys, stat

BASE = pathlib.Path(__file__).parent
VERSION = re.search(r'VERSION\s*=\s*"([^"]+)"', (BASE / "VSE.py").read_text(encoding="utf-8")).group(1)

# 1. Regenerate main-macos.py from VSE.py (single source of truth)
print("Regenerating macOS source from VSE.py …")
subprocess.run([sys.executable, "apply_macos_patches.py"], cwd=BASE, check=True)

DIST = BASE / "dist"; DIST.mkdir(exist_ok=True)
zip_name = DIST / f"VisualStimEdger-{VERSION}-macOS-ARM64-source.zip"

README = f"""# VisualStimEdger {VERSION} — macOS

## Easiest way to run
1. Unzip this folder somewhere (e.g. Documents).
2. **Double-click `VisualStimEdger.command`.**
   - First run sets up everything automatically (a minute or two) and launches.
   - Every run after that just launches — instantly.

Nothing is installed system-wide; dependencies live in a `.venv` folder next to
the app. To uninstall, delete the folder.

> If macOS says the launcher "can't be opened" (Gatekeeper), right-click it →
> **Open** → **Open**, just the first time. Or, in Terminal:
> `chmod +x VisualStimEdger.command` then double-click.

## Requirements
- macOS 13+ (Apple Silicon or Intel)
- Python 3.10+ with Tk. If the launcher says it's missing, install via Homebrew:
  ```
  brew install python@3.12 python-tk@3.12
  ```
  (Get Homebrew at https://brew.sh first.)

## Manual route (if you prefer)
```
python3.12 -m venv .venv
./.venv/bin/python -m pip install -r requirements-macos.txt
./.venv/bin/python main.py
```

Support: https://ko-fi.com/stimstation
"""

# 2. Assemble the zip
files = [
    ("main-macos.py",          "main.py"),
    ("requirements-macos.txt", "requirements-macos.txt"),
    ("icon.ico",               "icon.ico"),
    ("splash.png",             "splash.png"),
    ("overlay.html",           "overlay.html"),
    ("ko-fi.png",              "ko-fi.png"),
]

with zipfile.ZipFile(zip_name, "w", zipfile.ZIP_DEFLATED, compresslevel=9) as z:
    z.writestr("README.md", README)

    # Launcher — write with the executable bit set so double-click works on macOS
    launcher = (BASE / "macos-launcher.command").read_text(encoding="utf-8")
    info = zipfile.ZipInfo("VisualStimEdger.command")
    info.external_attr = (0o755 & 0xFFFF) << 16   # -rwxr-xr-x
    info.compress_type = zipfile.ZIP_DEFLATED
    z.writestr(info, launcher)

    for src, dst in files:
        p = BASE / src
        if p.exists():
            z.write(p, dst)
        else:
            print(f"  WARN: {src} missing, skipped")

    models = BASE / "models"
    for f in models.rglob("*"):
        if f.is_file():
            z.write(f, f"models/{f.relative_to(models)}")

print(f"Built {zip_name.name}  ({zip_name.stat().st_size/1e6:.1f} MB)")
