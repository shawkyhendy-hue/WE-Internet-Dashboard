@echo off
setlocal
cd /d "%~dp0"

echo ================================================
echo WE Dashboard - Python 3.13 Installer
echo ================================================

py -3.13 --version >nul 2>&1
if errorlevel 1 (
  echo Python 3.13 was not found. Trying Windows Python Install Manager...
  py install 3.13 >nul 2>&1
)
py -3.13 --version >nul 2>&1
if errorlevel 1 (
  echo [ERROR] Python 3.13 x64 is not installed or not visible to the Python launcher.
  echo Install Python 3.13 from https://www.python.org/downloads/windows/
  echo Then reopen this installer.
  pause
  exit /b 1
)

if exist .venv (
  echo Removing old virtual environment...
  rmdir /s /q .venv
)

echo Creating Python 3.13 virtual environment...
py -3.13 -m venv .venv
if errorlevel 1 goto :fail
call .venv\Scripts\activate.bat
if errorlevel 1 goto :fail

echo Updating pip, setuptools and wheel...
python -m pip install --upgrade pip setuptools wheel
if errorlevel 1 goto :fail

echo Installing Python 3.13 compatible packages...
python -m pip install --only-binary=:all: greenlet==3.5.5
if errorlevel 1 goto :fail
python -m pip install -r requirements.txt
if errorlevel 1 goto :fail

echo Installing Chromium for Playwright...
python -m playwright install chromium
if errorlevel 1 goto :fail

if not exist .env copy /y .env.example .env >nul

echo.
echo ================================================
echo Installation completed successfully.
echo Run: run_dashboard.bat
echo Then open: http://127.0.0.1:5000
echo ================================================
pause
exit /b 0

:fail
echo.
echo [ERROR] Installation failed. Copy the error above for troubleshooting.
pause
exit /b 1
