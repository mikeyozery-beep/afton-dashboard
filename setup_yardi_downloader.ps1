# Setup YARDI Report Downloader dependencies

Write-Host "========================================"
Write-Host "YARDI REPORT DOWNLOADER - SETUP"
Write-Host "========================================"

Write-Host ""
Write-Host "[1/2] Installing pywin32 for Outlook access..."
pip install pywin32 --break-system-packages

if ($LASTEXITCODE -ne 0) {
    Write-Host "ERROR: Failed to install pywin32" -ForegroundColor Red
    exit 1
}

Write-Host ""
Write-Host "[2/2] Configuring pywin32..."
python -m pip install --upgrade pywin32 --break-system-packages
python -m pywin32_postinstall -install

if ($LASTEXITCODE -ne 0) {
    Write-Host "WARNING: pywin32 configuration may need manual setup" -ForegroundColor Yellow
}

Write-Host ""
Write-Host "========================================"
Write-Host "SETUP COMPLETE"
Write-Host "========================================"
Write-Host ""
Write-Host "To run the downloader, use:"
Write-Host "  python C:\Afton\Dashboard\download_yardi_reports.py"
Write-Host ""
Write-Host "Requirements:"
Write-Host "  - Outlook must be running"
Write-Host "  - Must have access to mikereports@aftonprop.com inbox"
Write-Host ""
