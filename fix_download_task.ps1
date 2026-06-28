# =====================================================================
# ONE-TIME FIX for the 3 Afton Dashboard scheduled tasks.
# Run this ONCE in an ELEVATED PowerShell (Run as administrator).
#
# Root cause shared by all: tasks launched bare "python", which on this
# PC resolves to the Windows Store alias (a 0-byte stub) that does NOT
# run under Task Scheduler -> 0x80070002. Plus task #1 was elevated
# (can't COM-attach to your normal Outlook) and only downloaded (never
# organized); task #2 pointed at a .bat that does not exist (0x1).
# =====================================================================

$PY = "C:\Users\MichaelOzery\AppData\Local\Python\bin\python.exe"
if (-not (Test-Path $PY)) { Write-Host "[ERROR] python not found at $PY"; exit 1 }

function Show-Task($name) {
    $t = Get-ScheduledTask -TaskName $name
    $t.Actions | ForEach-Object { Write-Host ("  Execute : {0}`n  Args    : {1}`n  WorkDir : {2}" -f $_.Execute, $_.Arguments, $_.WorkingDirectory) }
    Write-Host ("  Logon   : {0}  RunLevel: {1}  State: {2}" -f $t.Principal.LogonType, $t.Principal.RunLevel, $t.State)
}

# ---------------------------------------------------------------------
# TASK 1 - 5 AM: download from mikereports@aftonprop.com + organize
#   - real python  - unified download+organize script  - non-elevated
# ---------------------------------------------------------------------
$n1 = "YARDI Report Downloader - Daily 5 AM"
Write-Host "`n[1/3] Fixing: $n1"
Set-ScheduledTask -TaskName $n1 `
    -Action (New-ScheduledTaskAction -Execute $PY -Argument '"C:\Afton\Dashboard\unified_download_organize.py"' -WorkingDirectory "C:\Afton\Dashboard") `
    -Principal (New-ScheduledTaskPrincipal -UserId "MichaelOzery" -LogonType Interactive -RunLevel Limited) | Out-Null
Show-Task $n1

# ---------------------------------------------------------------------
# TASK 2 - 7 AM: metrics extraction (real entry point = main_extractor.py)
#   - real python  - correct script  - workdir = extraction_system (relative paths)
# ---------------------------------------------------------------------
$n2 = "Afton Dashboard - Daily"
Write-Host "`n[2/3] Fixing: $n2"
Set-ScheduledTask -TaskName $n2 `
    -Action (New-ScheduledTaskAction -Execute $PY -Argument '"C:\Afton\Dashboard\extraction_system\main_extractor.py"' -WorkingDirectory "C:\Afton\Dashboard\extraction_system") `
    -Principal (New-ScheduledTaskPrincipal -UserId "MichaelOzery" -LogonType Interactive -RunLevel Limited) | Out-Null
Show-Task $n2

# ---------------------------------------------------------------------
# TASK 3 - 9 AM: extract metrics + push to GitHub
#   - real python  (GITHUB_TOKEN already set machine-level, git installed)
# ---------------------------------------------------------------------
$n3 = "Afton Dashboard - Daily Metrics Extract"
Write-Host "`n[3/3] Fixing: $n3"
Set-ScheduledTask -TaskName $n3 `
    -Action (New-ScheduledTaskAction -Execute $PY -Argument '"C:\Afton\Dashboard\extract_and_push_to_github.py"' -WorkingDirectory "C:\Afton\Dashboard") `
    -Principal (New-ScheduledTaskPrincipal -UserId "MichaelOzery" -LogonType Interactive -RunLevel Limited) | Out-Null
Show-Task $n3

Write-Host "`n=== ALL THREE TASKS UPDATED ==="
Write-Host "They now use the real python.exe and run non-elevated so they can read your Outlook."
Write-Host "Daily order is preserved: 5AM download+organize -> 7AM extract -> 9AM extract+push."
