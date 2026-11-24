@echo off
setlocal
cd /d "%~dp0"
if not exist ".venv\Scripts\python.exe" (
  echo [ERROR] Virtual environment not found. Run install_python313.bat first.
  pause
  exit /b 1
)
echo Updating dashboard dependencies...
".venv\Scripts\python.exe" -m pip install -r requirements.txt
if errorlevel 1 (
  echo [ERROR] Dependency update failed.
  pause
  exit /b 1
)
echo Dependencies updated successfully.
pause
