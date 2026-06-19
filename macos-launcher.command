#!/bin/bash
# ─────────────────────────────────────────────────────────────────────────────
# VisualStimEdger — macOS launcher.  Double-click this file to run.
# First run sets up an isolated virtual environment and installs dependencies;
# every run after that just launches the app. Nothing is installed system-wide.
# ─────────────────────────────────────────────────────────────────────────────
cd "$(dirname "$0")" || exit 1

echo "════════════════════════════════════════════"
echo "  VisualStimEdger — macOS launcher"
echo "════════════════════════════════════════════"
echo

pause_exit() { echo; echo "$1"; echo "Press Return to close this window."; read -r _; exit "${2:-1}"; }

# ── Find a usable Python 3 (prefer 3.12) ────────────────────────────────────
PY=""
for c in python3.12 python3.11 python3.13 python3.10 python3; do
  if command -v "$c" >/dev/null 2>&1; then PY="$c"; break; fi
done
if [ -z "$PY" ]; then
  echo "Python 3 isn't installed."
  echo
  echo "Install Homebrew (https://brew.sh), then run:"
  echo "    brew install python@3.12 python-tk@3.12"
  pause_exit "Then double-click this launcher again."
fi
echo "Using: $($PY --version 2>&1)  ($PY)"

# ── tkinter must be present (the UI needs it) ────────────────────────────────
if ! "$PY" -c "import tkinter" >/dev/null 2>&1; then
  PYVER="$("$PY" -c 'import sys;print(f"{sys.version_info.major}.{sys.version_info.minor}")' 2>/dev/null)"
  echo "Your Python is missing Tk (the GUI toolkit)."
  echo
  echo "Install it with Homebrew:"
  echo "    brew install python-tk@${PYVER:-3.12}"
  pause_exit "Then double-click this launcher again."
fi

# ── First-run setup: isolated venv + dependencies ───────────────────────────
if [ ! -d ".venv" ] || [ ! -x ".venv/bin/python" ]; then
  echo
  echo "First run — setting things up (a minute or two)…"
  "$PY" -m venv .venv || pause_exit "Could not create the virtual environment."
  ./.venv/bin/python -m pip install --upgrade pip >/dev/null 2>&1
  echo "Installing dependencies…"
  if ! ./.venv/bin/python -m pip install -r requirements-macos.txt; then
    rm -rf .venv
    pause_exit "Dependency install failed (see messages above). Check your internet connection and try again."
  fi
  echo "Setup complete."
fi

# ── Launch ──────────────────────────────────────────────────────────────────
echo
echo "Starting VisualStimEdger…"
./.venv/bin/python main.py
status=$?
if [ $status -ne 0 ]; then
  pause_exit "VisualStimEdger exited with an error (code $status)." "$status"
fi
