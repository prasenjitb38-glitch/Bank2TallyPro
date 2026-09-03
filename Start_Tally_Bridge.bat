@echo off
setlocal
title Bank2Tally Local Bridge
set /p BRIDGE_TOKEN=Enter the same Bridge token configured in Render: 
if "%BRIDGE_TOKEN%"=="" (
  echo A non-empty token is required.
  pause
  exit /b 1
)
set BRIDGE_HOST=127.0.0.1
set BRIDGE_PORT=9010
set TALLY_URL=http://127.0.0.1:9000
python "%~dp0tally_bridge.py"
pause
