@echo off
title VisualStimEdger — Build
setlocal

:: ── Check Python ────────────────────────────────────────────────────────────
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo.
    echo  Python not found.
    echo  Install Python 3.10+ from https://www.python.org/downloads/
    echo  Tick "Add Python to PATH" during install, then run this file again.
    echo.
    pause
    exit /b 1
)

:: ── Install dependencies (pre-built wheels only — no compiler needed) ───────
echo [1/3] Installing dependencies (pre-built wheels only)...
pip install --only-binary :all: -r requirements.txt --quiet
if %errorlevel% neq 0 (
    echo.
    echo  Wheel install failed. Your Python version may not have pre-built wheels.
    echo  Make sure you are using Python 3.10, 3.11, 3.12, or 3.13 (64-bit).
    echo  Download from https://www.python.org/downloads/
    echo.
    pause
    exit /b 1
)

:: ── Install Nuitka + Zig backend (no VS BuildTools needed) ──────────────────
echo [2/3] Installing Nuitka build tools...
pip install nuitka zstandard --quiet
if %errorlevel% neq 0 (
    echo.
    echo  Failed to install Nuitka. Try: pip install nuitka zstandard
    echo.
    pause
    exit /b 1
)

:: ── Build ───────────────────────────────────────────────────────────────────
echo [3/3] Building VisualStimEdger.exe (this takes a few minutes)...
echo        Using Zig compiler — no Visual Studio required.
echo.
python -m nuitka main.py ^
    --onefile ^
    --zig ^
    --assume-yes-for-downloads ^
    --windows-console-mode=disable ^
    --windows-icon-from-ico=icon.ico ^
    --output-filename=VisualStimEdger.exe ^
    --output-dir=dist_nuitka ^
    --include-data-files=icon.ico=icon.ico ^
    --include-data-files=splash.png=splash.png ^
    --include-data-files=models/yolo-fastest.cfg=models/yolo-fastest.cfg ^
    --include-data-files=models/best.weights=models/best.weights ^
    --include-module=sounddevice ^
    --include-module=miniaudio
if %errorlevel% neq 0 (
    echo.
    echo  Build failed. Check the output above for errors.
    echo.
    pause
    exit /b 1
)

echo.
echo  Build complete: dist_nuitka\VisualStimEdger.exe
echo.
pause
