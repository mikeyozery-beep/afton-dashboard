# PowerShell Test Script for Enhanced Metrics Extraction
# Run this directly in Windows PowerShell from C:\Afton\Dashboard

Write-Host "================================" -ForegroundColor Cyan
Write-Host "AFTON METRICS EXTRACTION TEST" -ForegroundColor Cyan
Write-Host "================================" -ForegroundColor Cyan
Write-Host ""

# Test file paths
$greenbooks = "C:\Users\MichaelOzery\OneDrive - Afton Properties\Old Dropbox\My PC (DESKTOP-5D77V89)\Mike Ozery\Greenbooks"
$financial = "C:\Users\MichaelOzery\OneDrive - Afton Properties\Old Dropbox\My PC (DESKTOP-5D77V89)\Mike Ozery\5. Monthly Financial Reviews"

Write-Host "1. Testing Greenbooks folder access..." -ForegroundColor Yellow
if (Test-Path $greenbooks) {
    Write-Host "   ✓ Greenbooks folder found" -ForegroundColor Green

    $files = @("Occupancy Analysis.xlsx", "DQ Reports.xlsm", "Effective Rent Analysis.xlsx")
    foreach ($file in $files) {
        $path = Join-Path $greenbooks $file
        if (Test-Path $path) {
            Write-Host "   ✓ $file found" -ForegroundColor Green
        } else {
            Write-Host "   ✗ $file NOT found" -ForegroundColor Red
        }
    }
} else {
    Write-Host "   ✗ Greenbooks folder NOT found" -ForegroundColor Red
}

Write-Host ""
Write-Host "2. Testing Financial Reviews folder access..." -ForegroundColor Yellow
if (Test-Path $financial) {
    Write-Host "   ✓ Financial Reviews folder found" -ForegroundColor Green
    $files = Get-ChildItem $financial -Filter "*.xlsx" -ErrorAction SilentlyContinue
    if ($files.Count -gt 0) {
        Write-Host "   ✓ Found $($files.Count) Excel file(s)" -ForegroundColor Green
        Write-Host "   Latest file: $($files[-1].Name)" -ForegroundColor Green
    } else {
        Write-Host "   ✗ No Excel files found in folder" -ForegroundColor Red
    }
} else {
    Write-Host "   ✗ Financial Reviews folder NOT found" -ForegroundColor Red
}

Write-Host ""
Write-Host "3. Running Python extraction script..." -ForegroundColor Yellow
Write-Host ""

# Run the Python script
python extract_metrics_enhanced.py

Write-Host ""
Write-Host "4. Checking output..." -ForegroundColor Yellow
$metricsFile = "C:\Afton\Dashboard\data\metrics.json"
if (Test-Path $metricsFile) {
    Write-Host "   ✓ metrics.json created successfully" -ForegroundColor Green
    Write-Host ""
    Write-Host "   Contents:" -ForegroundColor Cyan
    Get-Content $metricsFile | Write-Host
} else {
    Write-Host "   ✗ metrics.json NOT created" -ForegroundColor Red
}

Write-Host ""
Write-Host "================================" -ForegroundColor Cyan
Write-Host "TEST COMPLETE" -ForegroundColor Cyan
Write-Host "================================" -ForegroundColor Cyan
