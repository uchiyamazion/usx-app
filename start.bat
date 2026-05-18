@echo off
title USX App
echo Starting USX server...
echo.
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python not found.
    echo Please install Python 3 from https://www.python.org/downloads/
    echo Make sure to check "Add Python to PATH" during installation.
    pause
    exit /b 1
)
echo Installing required libraries...
python -m pip install flask openpyxl --quiet --exists-action i
echo.
echo Server starting... Browser will open automatically.
echo To stop: close this window.
echo.
python server.py
pause
