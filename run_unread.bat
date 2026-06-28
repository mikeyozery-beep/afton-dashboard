@echo off
powershell -ExecutionPolicy Bypass -File "%~dp0download_unread.ps1" > "%~dp0logs\unread_pull.txt" 2>&1
