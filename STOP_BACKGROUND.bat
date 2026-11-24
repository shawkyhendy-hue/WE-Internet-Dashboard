@echo off
setlocal
echo ================================================================
echo Internet Dashboard - Stop Background Service
echo ================================================================
set "FOUND="
for /f "tokens=5" %%P in ('netstat -ano ^| findstr ":5000" ^| findstr "LISTENING"') do (
  set "FOUND=1"
  echo Stopping process PID %%P listening on port 5000...
  taskkill /PID %%P /F >nul 2>&1
  if errorlevel 1 (
    echo [ERROR] Could not stop PID %%P. Try Run as administrator.
  ) else (
    echo [OK] Internet Dashboard background service stopped.
  )
)
if not defined FOUND echo Internet Dashboard is not currently running on port 5000.
echo.
echo Auto Start remains installed. The dashboard will start again at the next Windows sign-in.
echo To start it now, run START_BACKGROUND_NOW.bat
echo To disable future Auto Start, run REMOVE_AUTOSTART.bat
pause
