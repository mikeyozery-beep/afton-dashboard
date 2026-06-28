# Setup Windows Task Scheduler for Afton Dashboard Daily Automation
# Run this in PowerShell as Administrator

# Configuration
$TaskName = "Afton Dashboard - Daily Metrics Extract"
$TaskDescription = "Extract metrics from Excel and push to GitHub daily"
$ScriptPath = "C:\Afton\Dashboard\extract_and_push_to_github.py"
$PythonPath = "python"
$ScheduleTime = "09:00"

# Check if running as Administrator
$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")
if (-not $isAdmin) {
    Write-Host "ERROR: This script must run as Administrator" -ForegroundColor Red
    Write-Host "Right-click PowerShell and select 'Run as Administrator'" -ForegroundColor Yellow
    exit 1
}

Write-Host "========================================"
Write-Host "SETTING UP TASK SCHEDULER"
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

# Create trigger (daily at 9 AM)
Write-Host ""
Write-Host "[2/4] Creating daily trigger (9:00 AM)..."
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
$settings = New-ScheduledTaskSettingsSet -RunOnlyIfNetworkAvailable -StartWhenAvailable
$principal = New-ScheduledTaskPrincipal -UserId $env:USERNAME -LogonType Interactive -RunLevel Highest

Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $trigger -Principal $principal -Settings $settings -Description $TaskDescription | Out-Null
Write-Host "  Task registered"

Write-Host ""
Write-Host "========================================"
Write-Host "TASK SCHEDULER SETUP COMPLETE"
Write-Host "========================================"
Write-Host ""
Write-Host "Task Name: $TaskName"
Write-Host "Schedule: Daily at $ScheduleTime"
Write-Host "Script: $ScriptPath"
Write-Host "Status: Enabled"
Write-Host ""
Write-Host "Next Run: Tomorrow at $ScheduleTime"
Write-Host ""
Write-Host "To test manually, run:"
Write-Host "  python C:\Afton\Dashboard\extract_and_push_to_github.py"
Write-Host ""
