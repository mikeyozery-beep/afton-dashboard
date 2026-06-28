# Afton Dashboard - scheduled task setup
# RUN THIS AS ADMINISTRATOR (right-click > Run with PowerShell, or from an elevated prompt).
# It repoints the existing 7am task to the new auto-update pipeline and disables the
# obsolete 9am task. Re-runnable / safe to run more than once.

$ErrorActionPreference = "Stop"
$py     = "C:\Users\MichaelOzery\AppData\Local\Python\bin\python.exe"
$repo   = "C:\Afton\Dashboard\extraction_system\outputs\afton-dashboard"
$script = Join-Path $repo "update_dashboard.py"

if (-not (Test-Path $py))     { throw "Python not found at $py" }
if (-not (Test-Path $script)) { throw "update_dashboard.py not found at $script" }

$action = New-ScheduledTaskAction -Execute $py -Argument ('"' + $script + '"') -WorkingDirectory $repo

# Set principal to "Run only when user is logged on" (Interactive) -- most reliable for
# OneDrive-hosted source files, and needs no stored password.
$existing = Get-ScheduledTask -TaskName "Afton Dashboard - Daily"
$userId   = if ($existing.Principal.UserId) { $existing.Principal.UserId } else { "$env:USERDOMAIN\$env:USERNAME" }
$runLevel = if ($existing.Principal.RunLevel) { $existing.Principal.RunLevel } else { "Limited" }
$principal = New-ScheduledTaskPrincipal -UserId $userId -LogonType Interactive -RunLevel $runLevel

# 1) Repoint the daily 7am task to the new pipeline + run-only-when-logged-on (keeps its 7am trigger)
Set-ScheduledTask -TaskName "Afton Dashboard - Daily" -Action $action -Principal $principal | Out-Null
Write-Host "[OK] 'Afton Dashboard - Daily' now runs update_dashboard.py (rebuild + push)." -ForegroundColor Green
Write-Host "[OK] Set to 'Run only when user is logged on' (user: $userId)." -ForegroundColor Green

# 2) Disable the obsolete stale-metrics task
try {
    Disable-ScheduledTask -TaskName "Afton Dashboard - Daily Metrics Extract" | Out-Null
    Write-Host "[OK] Disabled obsolete 'Afton Dashboard - Daily Metrics Extract'." -ForegroundColor Green
} catch {
    Write-Host "[skip] 'Afton Dashboard - Daily Metrics Extract' not found / already gone." -ForegroundColor Yellow
}

Write-Host "`n=== Verification ===" -ForegroundColor Cyan
Get-ScheduledTask -TaskName "Afton Dashboard*" |
    Select-Object TaskName, State,
        @{n="LogonType";e={$_.Principal.LogonType}},
        @{n="Runs";e={$_.Actions.Execute + " " + $_.Actions.Arguments}} |
    Format-List

Write-Host "Done. The live dashboard will refresh each morning at 7:00 AM (while you're logged in)." -ForegroundColor Green
Write-Host "To test now, run:  python `"$script`"" -ForegroundColor Gray
