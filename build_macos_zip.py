"""
Build the macOS source distribution tarball (.tar.gz).

tar stores the unix executable bit natively, so the launcher extracts already
runnable on every macOS extractor — no "you do not have permission to open"
that a .zip's flakier exec-bit handling causes.

Regenerates the patched macOS source from VSE.py, then packages it with a
double-clickable launcher (VisualStimEdger.command) that creates an isolated
venv, installs deps, and runs the app — so end users never touch pip or main.py.

Usage:  python build_macos_zip.py
Output: dist/VisualStimEdger-<VERSION>-macOS-ARM64-source.tar.gz
"""
import re, tarfile, io, time, pathlib, subprocess, sys

BASE = pathlib.Path(__file__).parent
VERSION = re.search(r'VERSION\s*=\s*"([^"]+)"', (BASE / "VSE.py").read_text(encoding="utf-8")).group(1)

# 1. Regenerate main-macos.py from VSE.py (single source of truth)
print("Regenerating macOS source from VSE.py …")
subprocess.run([sys.executable, "apply_macos_patches.py"], cwd=BASE, check=True)

DIST = BASE / "dist"; DIST.mkdir(exist_ok=True)
tar_name = DIST / f"VisualStimEdger-{VERSION}-macOS-ARM64-source.tar.gz"

README = f"""# VisualStimEdger {VERSION} — macOS

## Easiest way to run
1. Put this folder somewhere handy (e.g. Documents).
2. **Double-click `VisualStimEdger.command`.**
   - First run sets up everything automatically (a minute or two) and launches.
   - Every run after that just launches — instantly.

Nothing is installed system-wide; dependencies live in a `.venv` folder next to
the app. To uninstall, delete the folder.

## macOS won't open it?  ("no permission" / "unidentified developer")

This is normal for anything downloaded outside the App Store — macOS strips the
launcher's permission to run. **One line fixes it.** Open **Terminal**, paste the
line below, press Return. It unblocks the launcher, marks it runnable, and starts it:

```
cd ~/Downloads/VisualStimEdger-*-macOS* && xattr -cr . && chmod +x VisualStimEdger.command && open VisualStimEdger.command
```

Extracted somewhere other than Downloads? Type `cd ` (with a trailing space), drag the
extracted folder into the Terminal window, then paste the rest starting at `&&`.

Note: **Get Info → read & write does NOT fix this.** A `.command` needs the *execute*
bit, which Finder never shows you — the line above is what actually sets it.

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

# 2. Assemble the tarball
files = [
    ("main-macos.py",          "main.py"),
    ("requirements-macos.txt", "requirements-macos.txt"),
    ("icon.ico",               "icon.ico"),
    ("splash.png",             "splash.png"),
    ("splash_logo_sheet.png",  "splash_logo_sheet.png"),
    ("splash_logo_meta.json",  "splash_logo_meta.json"),
    ("overlay.html",           "overlay.html"),
    ("ko-fi.png",              "ko-fi.png"),
    ("i18n/vse_ui_strings_zh.tsv", "i18n/vse_ui_strings_zh.tsv"),
    ("i18n/vse_ui_strings_es.tsv", "i18n/vse_ui_strings_es.tsv"),
    ("i18n/vse_ui_strings_de.tsv", "i18n/vse_ui_strings_de.tsv"),
    ("i18n/vse_ui_strings_fr.tsv", "i18n/vse_ui_strings_fr.tsv"),
    ("i18n/vse_ui_strings_ru.tsv", "i18n/vse_ui_strings_ru.tsv"),
    ("i18n/vse_ui_strings_pt.tsv", "i18n/vse_ui_strings_pt.tsv"),
]

def _add_bytes(tar, arcname, data: bytes, mode: int):
    """Add in-memory content (README, launcher) with an explicit unix mode."""
    ti = tarfile.TarInfo(arcname)
    ti.size = len(data)
    ti.mode = mode
    ti.mtime = int(time.time())
    ti.uid = ti.gid = 0
    ti.uname = ti.gname = ""
    tar.addfile(ti, io.BytesIO(data))


def _norm(ti):
    """Deterministic perms for on-disk files: 0o644, no host uid/gid/username leak
    (Windows has no real unix mode, so never trust the source file's stat)."""
    ti.mode = 0o644
    ti.uid = ti.gid = 0
    ti.uname = ti.gname = ""
    return ti


with tarfile.open(tar_name, "w:gz", compresslevel=9) as tar:
    _add_bytes(tar, "README.md", README.encode("utf-8"), 0o644)

    # Launcher gets the executable bit (0o755). tar stores the unix mode natively, so
    # every macOS extractor preserves +x and double-click just works — no "you do not
    # have permission to open the command file" on a fresh download the way a .zip
    # (exec bit in an optional external_attr that extractors flake on) can produce.
    launcher = (BASE / "macos-launcher.command").read_text(encoding="utf-8")
    _add_bytes(tar, "VisualStimEdger.command", launcher.encode("utf-8"), 0o755)

    for src, dst in files:
        p = BASE / src
        if p.exists():
            tar.add(p, arcname=dst, filter=_norm)
        else:
            print(f"  WARN: {src} missing, skipped")

    models = BASE / "models"
    for f in models.rglob("*"):
        if f.is_file():
            tar.add(f, arcname=f"models/{f.relative_to(models)}", filter=_norm)

print(f"Built {tar_name.name}  ({tar_name.stat().st_size/1e6:.1f} MB)")
