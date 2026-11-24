@echo off
setlocal
cd /d "%~dp0"
echo ================================================================
echo Internet Dashboard - Install Windows Auto Start
echo ================================================================
if not exist ".venv\Scripts\pythonw.exe" (
  echo [ERROR] Python virtual environment not found.
  echo Run install_python313.bat first.
  pause
  exit /b 1
)
set "APPDIR=%CD%"
set "LINK=%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup\Internet Dashboard.lnk"
powershell -NoProfile -ExecutionPolicy Bypass -Command "$ws=New-Object -ComObject WScript.Shell; $sc=$ws.CreateShortcut('%LINK%'); $sc.TargetPath='%APPDIR%\.venv\Scripts\pythonw.exe'; $sc.Arguments='\"%APPDIR%\app.py\"'; $sc.WorkingDirectory='%APPDIR%'; $sc.Description='Internet Dashboard background server and hourly WE sync'; $sc.Save()"
if errorlevel 1 (
  echo [ERROR] Could not create the Windows Startup shortcut.
  pause
  exit /b 2
)
echo.
echo [OK] Internet Dashboard will now start automatically when you sign in to Windows.
echo [OK] It will run hidden in the background without run_dashboard.bat.
echo [OK] Scheduled WE sync interval: 60 minutes.
echo.
echo Starting it now in the background...
start "" /b ".venv\Scripts\pythonw.exe" "app.py"
timeout /t 3 /nobreak >nul
echo.
echo Open: http://127.0.0.1:5000/
pause
