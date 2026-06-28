@echo off
echo Splitting YARDI Workbooks into Individual Reports...
echo.
python "%~dp0split_yardi_workbooks.py"
echo.
echo Script completed. Press any key to close this window.
pause
