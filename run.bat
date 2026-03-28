@echo off
title VisualStimEdger

python --version >/dev/null 2>&1
if %errorlevel% neq 0 (
    echo.
    echo  Python not found.
    echo  Install Python 3.10 or newer from https://www.python.org/downloads/
    echo  Tick "Add Python to PATH" during install, then run this file again.
    echo.
    pause
    exit /b 1
)

echo Installing / updating dependencies...
pip install --only-binary :all: -r requirements.txt --quiet
if %errorlevel% neq 0 (
    echo.
    echo  Dependency install failed. Try running as Administrator.
    echo.
    pause
    exit /b 1
)

echo Starting VisualStimEdger...
python main.py
