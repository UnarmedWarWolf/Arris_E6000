@echo off
rem Double-click to build dist\installer\CMTS_Command_Center_Setup_<version>.exe
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0installer\build_installer.ps1"
echo.
pause
