@echo off
title AirCargo WMS - Ngrok Tunnel
echo ================================================
echo   AirCargo WMS - Starting internet sharing...
echo ================================================
echo.
echo   Make sure START_SERVER.bat is already running!
echo.

if not exist "%~dp0ngrok.exe" (
    echo ERROR: ngrok.exe not found in this folder.
    echo.
    echo Please follow these steps:
    echo 1. Go to https://ngrok.com and create a free account
    echo 2. Download ngrok for Windows from https://ngrok.com/download
    echo 3. Extract ngrok.exe and place it in this folder
    echo 4. Run: ngrok config add-authtoken YOUR_TOKEN
    echo    (find your token at https://dashboard.ngrok.com/get-started/your-authtoken^)
    echo.
    pause
    exit /b
)

echo ================================================
echo   Ngrok is starting...
echo   Look for a line that says "Forwarding"
echo   Your customer tracking page will be at:
echo   https://XXXXXX.ngrok.io/track
echo.
echo   Share this link with your customers!
echo   (The link changes each time you restart Ngrok)
echo.
echo   DO NOT close this window while customers
echo   need to access the tracking page.
echo ================================================
echo.

"%~dp0ngrok.exe" http 5000
pause
