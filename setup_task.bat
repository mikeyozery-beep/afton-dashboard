@echo off
echo Setting up Windows Task Scheduler for Daily YARDI Download...
echo.
powershell -ExecutionPolicy Bypass -File "%~dp0setup_daily_task.ps1"
echo.
echo Setup complete. Press any key to close this window.
pause
