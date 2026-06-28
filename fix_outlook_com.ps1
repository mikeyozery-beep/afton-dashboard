# Diagnose and fix Outlook COM registration issues
# Run as Administrator

Write-Host "========================================================================"
Write-Host "OUTLOOK COM DIAGNOSTIC & FIX SCRIPT"
Write-Host "========================================================================"
Write-Host ""

# Check if running as Administrator
$IsAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $IsAdmin) {
    Write-Host "[ERROR] This script must be run as Administrator"
    Write-Host "Please right-click PowerShell and select 'Run as Administrator'"
    exit 1
}

Write-Host "[STEP 1] Check if Outlook is installed..."
$OutlookPaths = @(
    "C:\Program Files\Microsoft Office\Office16\OUTLOOK.EXE",
    "C:\Program Files (x86)\Microsoft Office\Office16\OUTLOOK.EXE",
    "C:\Program Files\Microsoft Office\Office15\OUTLOOK.EXE",
    "C:\Program Files (x86)\Microsoft Office\Office15\OUTLOOK.EXE"
)

$OutlookFound = $false
foreach ($path in $OutlookPaths) {
    if (Test-Path $path) {
        Write-Host "  [OK] Found Outlook at: $path"
        $OutlookFound = $true
        break
    }
}

if (-not $OutlookFound) {
    Write-Host "  [ERROR] Outlook not found in common locations"
    Write-Host "  Please ensure Microsoft Outlook is installed"
    exit 1
}

Write-Host ""
Write-Host "[STEP 2] Check if Outlook is currently running..."
$OutlookRunning = Get-Process OUTLOOK -ErrorAction SilentlyContinue

if ($OutlookRunning) {
    Write-Host "  [OK] Outlook is running (PID: $($OutlookRunning.Id))"
}
else {
    Write-Host "  [WARNING] Outlook is not running"
    Write-Host "  Attempting to start Outlook..."
    try {
        Start-Process -FilePath $path -ErrorAction Stop
        Start-Sleep -Seconds 3
        Write-Host "  [OK] Outlook started"
    }
    catch {
        Write-Host "  [ERROR] Could not start Outlook: $_"
    }
}

Write-Host ""
Write-Host "[STEP 3] Check Outlook COM registration..."
try {
    $TestCom = New-Object -ComObject Outlook.Application -ErrorAction Stop
    Write-Host "  [OK] Outlook COM is accessible"
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($TestCom) | Out-Null
}
catch {
    Write-Host "  [ERROR] Outlook COM is NOT accessible: $_"
    Write-Host ""
    Write-Host "Attempting to re-register Outlook COM..."

    # Get Outlook installation path
    $RegPath = "HKLM:\Software\Microsoft\Windows\CurrentVersion\App Paths\OUTLOOK.EXE"
    $OutlookPath = (Get-ItemProperty $RegPath)."(Default)"

    if ($OutlookPath) {
        Write-Host "  Found Outlook at: $OutlookPath"

        # Run Outlook repair
        Write-Host "  Running: `"$OutlookPath`" /resetnavpane"
        & "$OutlookPath" /resetnavpane

        Start-Sleep -Seconds 2

        # Try COM again
        try {
            $TestCom = New-Object -ComObject Outlook.Application -ErrorAction Stop
            Write-Host "  [OK] Outlook COM now accessible after repair"
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($TestCom) | Out-Null
        }
        catch {
            Write-Host "  [ERROR] COM still not accessible: $_"
        }
    }
}

Write-Host ""
Write-Host "[STEP 4] Restart Outlook..."
try {
    $OutlookProcess = Get-Process OUTLOOK -ErrorAction SilentlyContinue
    if ($OutlookProcess) {
        Write-Host "  Closing Outlook..."
        Stop-Process -Name OUTLOOK -Force -ErrorAction Stop
        Start-Sleep -Seconds 2
        Write-Host "  [OK] Outlook closed"
    }

    Write-Host "  Starting Outlook..."
    Start-Process -FilePath $path -ErrorAction Stop
    Start-Sleep -Seconds 5
    Write-Host "  [OK] Outlook restarted"
}
catch {
    Write-Host "  [ERROR] Could not restart Outlook: $_"
}

Write-Host ""
Write-Host "[STEP 5] Final COM test..."
try {
    $TestCom = New-Object -ComObject Outlook.Application -ErrorAction Stop
    Write-Host "  [SUCCESS] Outlook COM is now working!"
    Write-Host "  You can now run: .\download_yardi_powershell.ps1"
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($TestCom) | Out-Null
}
catch {
    Write-Host "  [ERROR] Outlook COM still not accessible"
    Write-Host "  Error: $_"
    Write-Host ""
    Write-Host "RECOMMENDATION:"
    Write-Host "  1. Open Outlook manually"
    Write-Host "  2. Wait for it to fully load (2-3 seconds)"
    Write-Host "  3. Then run: .\download_yardi_powershell.ps1"
}

Write-Host ""
Write-Host "========================================================================"
Write-Host "DIAGNOSTIC COMPLETE"
Write-Host "========================================================================"
