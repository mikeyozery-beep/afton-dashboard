# Setup Windows Task Scheduler for YARDI Report Downloader
# Runs daily at 5:00 AM Eastern Time

# Configuration
$TaskName = "YARDI Report Downloader - Daily 5 AM"
$TaskDescription = "Download YARDI Scheduler Reports daily at 5 AM ET"
$ScriptPath = "C:\Afton\Dashboard\download_yardi_reports.py"
$PythonPath = "python"
$ScheduleTime = "05:00"  # 5:00 AM

# Check if running as Administrator
$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")
if (-not $isAdmin) {
    Write-Host "ERROR: This script must run as Administrator" -ForegroundColor Red
    Write-Host "Right-click PowerShell and select 'Run as Administrator'" -ForegroundColor Yellow
    exit 1
}

Write-Host "========================================"
Write-Host "YARDI REPORT DOWNLOADER - SCHEDULER SETUP"
Write-Host "========================================"

# Delete existing task if it exists
Write-Host ""
Write-Host "[1/4] Checking for existing task..."
$existingTask = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
if ($existingTask) {
    Write-Host "  Existing task found. Removing..."
    Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
    Write-Host "  Removed"
}

# Create trigger (daily at 5 AM Eastern)
Write-Host ""
Write-Host "[2/4] Creating daily trigger (5:00 AM Eastern)..."
$trigger = New-ScheduledTaskTrigger -Daily -At $ScheduleTime
Write-Host "  Trigger created"

# Create action (run Python script)
Write-Host ""
Write-Host "[3/4] Creating action (run Python script)..."
$action = New-ScheduledTaskAction -Execute $PythonPath -Argument $ScriptPath -WorkingDirectory "C:\Afton\Dashboard"
Write-Host "  Action created"
Write-Host "  Script: $ScriptPath"
Write-Host "  Working Directory: C:\Afton\Dashboard"

# Create task settings
Write-Host ""
Write-Host "[4/4] Registering task..."
$settings = New-ScheduledTaskSettingsSet -RunOnlyIfNetworkAvailable -StartWhenAvailable -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries
$principal = New-ScheduledTaskPrincipal -UserId $env:USERNAME -LogonType Interactive -RunLevel Highest

Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $trigger -Principal $principal -Settings $settings -Description $TaskDescription | Out-Null
Write-Host "  Task registered"

Write-Host ""
Write-Host "========================================"
Write-Host "SCHEDULER SETUP COMPLETE"
Write-Host "========================================"
Write-Host ""
Write-Host "Task Name: $TaskName"
Write-Host "Schedule: Daily at 5:00 AM Eastern Time"
Write-Host "Script: $ScriptPath"
Write-Host "Status: Enabled"
Write-Host ""
Write-Host "Next Run: Tomorrow at 5:00 AM ET"
Write-Host ""
Write-Host "To test manually, run:"
Write-Host "  python C:\Afton\Dashboard\download_yardi_reports.py"
Write-Host ""
Write-Host "To view task in Task Scheduler, search for:"
Write-Host "  Task Scheduler > Task Scheduler Library > $TaskName"
Write-Host ""
