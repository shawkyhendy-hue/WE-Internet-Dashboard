@echo off
set "DATA_DIR=%USERPROFILE%\WE-Dashboard-Data"
if not exist "%DATA_DIR%" mkdir "%DATA_DIR%"
explorer "%DATA_DIR%"
