@echo off
setlocal
cd /d "%~dp0"
echo ================================================================
echo WE Dashboard - Persistent Data v4
echo ================================================================
if not exist .venv\Scripts\python.exe (
  echo [ERROR] Virtual environment not found.
  echo Run install_python313.bat first.
  pause
  exit /b 1
)
for /f "tokens=5" %%P in ('netstat -ano ^| findstr ":5000" ^| findstr "LISTENING"') do (
  echo [WARNING] Port 5000 is already in use by PID %%P.
  echo This is usually an older WE Dashboard still running.
  echo Close its CMD window or press CTRL+C there, then run this file again.
  echo.
  pause
  exit /b 2
)
call .venv\Scripts\activate.bat
echo Persistent data: %USERPROFILE%\WE-Dashboard-Data
python app.py
pause
