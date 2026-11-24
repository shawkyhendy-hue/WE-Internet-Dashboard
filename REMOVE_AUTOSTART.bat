@echo off
setlocal
set "LINK=%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup\Internet Dashboard.lnk"
if exist "%LINK%" (
  del /q "%LINK%"
  echo [OK] Internet Dashboard auto start removed.
) else (
  echo Auto start shortcut was not found.
)
pause
