@echo off
REM Afton Dashboard - daily auto-update (point Windows Task Scheduler at this file)
cd /d "C:\Afton\Dashboard\extraction_system\outputs\afton-dashboard"
python update_dashboard.py
