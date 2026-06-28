@echo off
echo Cleaning up and organizing Dashboard Data folder...
echo This will:
echo   1. Extract Scheduler_Reports worksheets
echo   2. Handle generic-named files (Report One, etc.)
echo   3. Organize ALL reports into folders by type
echo   4. Delete original files
echo.
python "%~dp0cleanup_and_organize.py"
echo.
echo Cleanup complete. Press any key to close this window.
pause
