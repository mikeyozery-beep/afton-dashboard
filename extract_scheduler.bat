@echo off
echo Extracting worksheets from Scheduler_Reports files...
echo.
python "%~dp0extract_scheduler_reports.py"
echo.
echo Extraction complete. Press any key to close this window.
pause
