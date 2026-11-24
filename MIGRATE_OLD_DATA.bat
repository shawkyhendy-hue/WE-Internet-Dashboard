@echo off
setlocal EnableExtensions
cd /d "%~dp0"

echo ================================================================
echo WE Dashboard - Import Data From Previous Version
echo ================================================================
echo.
set "DATA_DIR=%USERPROFILE%\WE-Dashboard-Data"
if not exist "%DATA_DIR%" mkdir "%DATA_DIR%"
if not exist "%DATA_DIR%\sessions" mkdir "%DATA_DIR%\sessions"
if not exist "%DATA_DIR%\diagnostics" mkdir "%DATA_DIR%\diagnostics"
if not exist "%DATA_DIR%\backups" mkdir "%DATA_DIR%\backups"

echo Persistent destination:
echo   %DATA_DIR%
echo.
set /p "OLD_DIR=Paste the full path of the PREVIOUS dashboard folder: "
if "%OLD_DIR%"=="" goto :cancel
if not exist "%OLD_DIR%" (
  echo [ERROR] Folder not found: %OLD_DIR%
  pause
  exit /b 1
)

echo.
echo Copying database and encryption key...
if exist "%OLD_DIR%\we_dashboard.db" (
  copy /y "%OLD_DIR%\we_dashboard.db" "%DATA_DIR%\we_dashboard.db" >nul
  echo [OK] we_dashboard.db copied.
) else (
  echo [WARNING] we_dashboard.db was not found in the old folder.
)
if exist "%OLD_DIR%\.we_secret.key" (
  copy /y "%OLD_DIR%\.we_secret.key" "%DATA_DIR%\.we_secret.key" >nul
  echo [OK] .we_secret.key copied.
) else (
  echo [WARNING] .we_secret.key was not found. Saved encrypted passwords may need to be entered again.
)

if exist "%OLD_DIR%\sessions" (
  xcopy /E /I /Y "%OLD_DIR%\sessions" "%DATA_DIR%\sessions" >nul
  echo [OK] Saved WE browser sessions copied.
)
if exist "%OLD_DIR%\diagnostics" (
  xcopy /E /I /Y "%OLD_DIR%\diagnostics" "%DATA_DIR%\diagnostics" >nul
)

echo.
echo ================================================================
echo Import completed.
echo Your future versions will use the same persistent folder:
echo %DATA_DIR%
echo ================================================================
explorer "%DATA_DIR%"
pause
exit /b 0

:cancel
echo Import cancelled.
pause
