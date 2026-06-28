@echo off
echo Starting YARDI download script...
echo.
powershell -ExecutionPolicy Bypass -File "%~dp0download_yardi_outlook_com.ps1"
echo.
echo Script completed. Press any key to close this window.
pause
