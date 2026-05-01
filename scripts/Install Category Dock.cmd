@echo off
setlocal
cd /d "%~dp0.."
powershell -NoProfile -ExecutionPolicy Bypass -File ".\scripts\install-package.ps1"
echo.
echo Press any key to close.
pause >nul
