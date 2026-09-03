@echo off
setlocal
title Bank2Tally Hybrid Bridge
set /p BRIDGE_TOKEN=Enter the Bridge token configured in Render: 
if "%BRIDGE_TOKEN%"=="" (
  echo A non-empty token is required.
  pause
  exit /b 1
)
set BRIDGE_HOST=127.0.0.1
set BRIDGE_PORT=9010
set TALLY_URL=http://127.0.0.1:9000
start "Bank2Tally Local Bridge" /D "%~dp0" cmd /k "set BRIDGE_TOKEN=%BRIDGE_TOKEN%&& set BRIDGE_HOST=127.0.0.1&& set BRIDGE_PORT=9010&& set TALLY_URL=http://127.0.0.1:9000&& python tally_bridge.py"
where cloudflared >nul 2>nul
if errorlevel 1 (
  echo cloudflared was not found. Install it, then run: cloudflared tunnel --url http://127.0.0.1:9010
  pause
  exit /b 1
)
echo Starting a temporary secure tunnel. Copy its https URL into Render as TALLY_BRIDGE_URL.
cloudflared tunnel --url http://127.0.0.1:9010
