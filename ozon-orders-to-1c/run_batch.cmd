@echo off
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0make_batch.ps1"
echo.
echo Press any key to exit...
pause >nul
