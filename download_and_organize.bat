@echo off
echo ========================================
echo YARDI DOWNLOAD & ORGANIZE - FULL PROCESS
echo ========================================
echo.
echo This will:
echo   1. Download all files from mikereports@aftonprop.com inbox
echo   2. Extract zips and worksheets
echo   3. Organize by actual report name (A1)
echo.
echo Starting Step 1: Download...
echo.

REM Run download script
call "%~dp0run_download.bat"

echo.
echo ========================================
echo Step 1 complete. Starting Step 2...
echo ========================================
echo.

REM Run organization script
call "%~dp0proper_cleanup.bat"

echo.
echo ========================================
echo COMPLETE - All files downloaded and organized
echo ========================================
echo.
echo Dashboard Data folder is now fully populated and organized!
pause
