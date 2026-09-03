@echo off
setlocal
title Bank2Tally Connector
set /p TALLY_CONNECTOR_TOKEN=Enter the Connector token configured in Render: 
if "%TALLY_CONNECTOR_TOKEN%"=="" (
  echo A non-empty token is required.
  pause
  exit /b 1
)
set TALLY_URL=http://127.0.0.1:9000
python "%~dp0tally_outbound_connector.py"
pause
