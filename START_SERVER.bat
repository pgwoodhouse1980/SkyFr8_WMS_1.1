@echo off
title AirCargo WMS Server
echo ================================================
echo   AirCargo WMS - Starting server...
echo ================================================
echo.

python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python is not installed or not in PATH.
    echo Please install Python from https://www.python.org/downloads/
    echo Make sure to check "Add Python to PATH" during installation.
    pause
    exit /b
)

echo Installing required packages...
pip install flask --quiet

echo.
echo ================================================
echo   Server is starting...
echo   Staff portal:   http://localhost:5000
echo   Customer track: http://localhost:5000/track
echo.
echo   Share on your network using your IP address.
echo   Run ipconfig in Command Prompt to find it.
echo   Example: http://192.168.1.10:5000
echo ================================================
echo.
echo   DO NOT close this window while the system is in use.
echo.

python server.py
pause
