# Setup Windows Task Scheduler for Daily YARDI Download
# Runs C:\Afton\Dashboard\run_download.bat every day at 5:00 AM

$TaskName = "YARDI Daily Download - 5am"
$TaskPath = "\Afton\"
$ScriptPath = "C:\Afton\Dashboard\run_download.bat"
$TaskDescription = "Download YARDI reports from mikereports@aftonprop.com shared mailbox daily at 5:00 AM"

Write-Host "=========================================="
Write-Host "TASK SCHEDULER SETUP"
Write-Host "=========================================="
Write-Host ""
Write-Host "Task Name: $TaskName"
Write-Host "Schedule: Daily at 5:00 AM"
Write-Host "Action: Run $ScriptPath"
Write-Host ""

# Check if task already exists
$existingTask = Get-ScheduledTask -TaskName $TaskName -TaskPath $TaskPath -ErrorAction SilentlyContinue

if ($existingTask) {
    Write-Host "[WARNING] Task '$TaskName' already exists"
    Write-Host "Deleting existing task..."
    Unregister-ScheduledTask -TaskName $TaskName -TaskPath $TaskPath -Confirm:$false
    Write-Host "[OK] Deleted existing task"
    Write-Host ""
}

# Create trigger: Daily at 5:00 AM
$trigger = New-ScheduledTaskTrigger -Daily -At 5:00AM
Write-Host "[OK] Created daily trigger at 5:00 AM"

# Create action: Run both download and organize scripts
# First download, then organize
$downloadScript = "C:\Afton\Dashboard\run_download.bat"
$organizeScript = "python `"C:\Afton\Dashboard\organize_reports.py`""

$action = New-ScheduledTaskAction `
    -Execute "cmd.exe" `
    -Argument "/c `"$downloadScript && $organizeScript`"" `
    -WorkingDirectory "C:\Afton\Dashboard"

Write-Host "[OK] Created task action:"
Write-Host "  1. Download: $downloadScript"
Write-Host "  2. Organize: organize_reports.py"

# Create settings
$settings = New-ScheduledTaskSettingsSet `
    -AllowStartIfOnBatteries `
    -DontStopIfGoingOnBatteries `
    -RunOnlyIfNetworkAvailable `
    -MultipleInstances IgnoreNew `
    -StartWhenAvailable

Write-Host "[OK] Configured task settings"
Write-Host "  - Run with highest privileges"
Write-Host "  - Allow on battery"
Write-Host "  - Require network available"
Write-Host "  - Start if missed"
Write-Host ""

# Register the task
try {
    $task = Register-ScheduledTask `
        -TaskName $TaskName `
        -TaskPath $TaskPath `
        -Trigger $trigger `
        -Action $action `
        -Settings $settings `
        -Description $TaskDescription `
        -RunLevel Highest

    Write-Host "=========================================="
    Write-Host "[SUCCESS] TASK CREATED"
    Write-Host "=========================================="
    Write-Host ""
    Write-Host "Task Details:"
    Write-Host "  Name: $($task.TaskName)"
    Write-Host "  Path: $($task.TaskPath)"
    Write-Host "  State: $($task.State)"
    Write-Host "  Enabled: $($task.Enabled)"
    Write-Host ""
    Write-Host "Schedule: Daily at 5:00 AM"
    Write-Host "Action: $ScriptPath"
    Write-Host ""
    Write-Host "Next Run: Check Task Scheduler for next scheduled time"
    Write-Host ""
    Write-Host "=========================================="
    Write-Host "TASK IS ACTIVE AND WILL RUN DAILY AT 5AM"
    Write-Host "=========================================="
}
catch {
    Write-Host "=========================================="
    Write-Host "[ERROR] FAILED TO CREATE TASK"
    Write-Host "=========================================="
    Write-Host $_.Exception.Message
    exit 1
}
