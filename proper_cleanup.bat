@echo off
echo Proper Report Extraction and Organization
echo ==========================================
echo This will:
echo   1. Remove old poorly-organized folders
echo   2. Extract all zip files
echo   3. Extract worksheets from multi-sheet Excel files
echo   4. Read first row (A1) to get actual report name
echo   5. Organize into proper report name folders
echo   6. Delete original files
echo.
python "%~dp0proper_extract_organize.py"
echo.
echo Extraction complete. Press any key to close this window.
pause
