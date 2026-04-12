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

:: ── Install PyInstaller ──────────────────────────────────────────────────────
echo [2/3] Installing PyInstaller...
pip install pyinstaller --quiet
if %errorlevel% neq 0 (
    echo.
    echo  Failed to install PyInstaller. Try: pip install pyinstaller
    echo.
    pause
    exit /b 1
)

:: ── Build ───────────────────────────────────────────────────────────────────
echo [3/3] Building VisualStimEdger.exe (this takes a few minutes)...
echo.
python -m PyInstaller VisualStimEdger.spec --noconfirm --distpath dist
if %errorlevel% neq 0 (
    echo.
    echo  Build failed. Check the output above for errors.
    echo.
    pause
    exit /b 1
)

echo.
echo  Build complete: dist\VisualStimEdger.exe
echo.
pause
