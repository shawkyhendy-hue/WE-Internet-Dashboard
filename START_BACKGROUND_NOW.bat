@echo off
setlocal
cd /d "%~dp0"
if not exist ".venv\Scripts\pythonw.exe" (
  echo [ERROR] Run install_python313.bat first.
  pause
  exit /b 1
)
start "" /b ".venv\Scripts\pythonw.exe" "app.py"
timeout /t 2 /nobreak >nul
start "" "http://127.0.0.1:5000/"
