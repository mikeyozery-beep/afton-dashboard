# Simple test script
Write-Host "Running Metrics Extraction..." -ForegroundColor Green
Write-Host ""

python extract_metrics_enhanced.py

Write-Host ""
Write-Host "Checking output file..." -ForegroundColor Green
Write-Host ""

if (Test-Path "C:\Afton\Dashboard\data\metrics.json") {
    Write-Host "SUCCESS! metrics.json created." -ForegroundColor Green
    Write-Host ""
    Write-Host "File contents:" -ForegroundColor Cyan
    Get-Content "C:\Afton\Dashboard\data\metrics.json"
} else {
    Write-Host "ERROR: metrics.json not found" -ForegroundColor Red
}
