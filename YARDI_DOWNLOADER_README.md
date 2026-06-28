# YARDI Scheduler Report Downloader

## Overview

Automatically downloads YARDI Scheduler Reports from Outlook and organizes them by report title + date received.

**Example:**
- Email received: June 25, 2026 at 2:47 AM
- Attachment: `rs_sql_evictions_am_reports_ap-gs.xlsx`
- Workbook title: `rs_sql_evictions_am_reports`
- **Saved as:** `rs_sql_evictions_am_reports 6.25.2026.xlsx`
- **Location:** `C:\Users\MichaelOzery\OneDrive - Afton Properties\Old Dropbox\My PC (DESKTOP-5D77V89)\Mike Ozery\Dashboard Data\`

---

## Setup (One-Time)

### Step 1: Install Dependencies
Run PowerShell as Administrator:
```powershell
C:\Afton\Dashboard\setup_yardi_downloader.ps1
```

This installs:
- **pywin32** - Windows COM bridge for Outlook access
- **openpyxl** - Excel file reading

### Step 2: Verify Outlook
- Make sure Outlook is running
- Ensure you have access to `mikereports@aftonprop.com` inbox
- Test that emails from `cdr@yardi.com` are visible

---

## Usage

### Manual Run
```powershell
cd C:\Afton\Dashboard
python download_yardi_reports.py
```

### Scheduled (Optional)
Add to Windows Task Scheduler to run daily (e.g., at 10:00 AM after reports arrive):

```powershell
# Create scheduled task
$trigger = New-ScheduledTaskTrigger -Daily -At "10:00 AM"
$action = New-ScheduledTaskAction -Execute "python" -Argument "C:\Afton\Dashboard\download_yardi_reports.py" -WorkingDirectory "C:\Afton\Dashboard"
$principal = New-ScheduledTaskPrincipal -UserId $env:USERNAME -LogonType Interactive -RunLevel Highest
Register-ScheduledTask -TaskName "YARDI Report Downloader" -Action $action -Trigger $trigger -Principal $principal -Description "Download YARDI Scheduler Reports daily"
```

---

## What It Does

1. **Connects to Outlook** → Accesses `mikereports@aftonprop.com` inbox
2. **Filters emails** → Looks for messages from `cdr@yardi.com` (YARDI Scheduler)
3. **Processes attachments** → Downloads only Excel files (.xlsx, .xls)
4. **Reads workbook title** → Opens Excel to get the actual report name
5. **Renames files** → Format: `{report_title} {date_received}`
6. **Saves files** → To Dashboard Data folder
7. **Handles duplicates** → Appends number if file already exists (e.g., `report_1.xlsx`, `report_2.xlsx`)

---

## Output

**Location:** `C:\Users\MichaelOzery\OneDrive - Afton Properties\Old Dropbox\My PC (DESKTOP-5D77V89)\Mike Ozery\Dashboard Data\`

**Example Files:**
```
rs_sql_evictions_am_reports 6.25.2026.xlsx
rs_sql_income_statement_all_muvan 6.25.2026.xlsx
ResAnalytics_Box_Score_Summary 6.25.2026.xlsx
rs_sql_rent_roll_muvan 6.25.2026.xlsx
Unit_Occupancy_6413722 6.25.2026.xlsx
... (one file per report per date)
```

---

## Logs

Execution logs are saved to: `C:\Afton\Dashboard\logs\yardi_download_*.log`

Check logs if downloads fail:
```powershell
Get-Content C:\Afton\Dashboard\logs\yardi_download_*.log -Tail 50
```

---

## Troubleshooting

### "ERROR: pywin32 not installed"
**Solution:** Run setup script:
```powershell
C:\Afton\Dashboard\setup_yardi_downloader.ps1
```

### "Cannot connect to Outlook"
**Causes:**
- Outlook not running
- Wrong email address in config
- No access to inbox

**Solution:**
1. Start Outlook
2. Verify `mikereports@aftonprop.com` inbox is accessible
3. Check that emails from `cdr@yardi.com` exist in inbox

### "Could not determine title"
**Cause:** Excel file doesn't have workbook properties set

**Solution:** Script falls back to filename - check log for details

### "Access is denied"
**Cause:** Outlook/Windows permission issue

**Solution:**
1. Restart Outlook
2. Run PowerShell as Administrator
3. Restart computer

---

## Integration with Extraction Pipeline

Once reports are downloaded and organized, they can be fed into the metrics extraction pipeline:

1. ✅ Reports downloaded daily → Dashboard Data folder
2. ⏳ Extract metrics from reports (Phase 2 work)
3. ⏳ Feed into Python extraction script
4. ⏳ Push to GitHub + Dashboard update

---

## Configuration

To modify the script, edit these lines in `download_yardi_reports.py`:

```python
OUTPUT_DIR = Path(r"C:\Users\...")  # Where to save files
OUTLOOK_INBOX = "mikereports@aftonprop.com"  # Email to connect to
SENDER_EMAIL = "cdr@yardi.com"  # YARDI scheduler email
TEMP_DIR = SCRIPT_DIR / "temp_yardi_downloads"  # Temporary working folder
```

---

## Next Steps

1. ✅ Run setup script
2. ✅ Test manual run: `python download_yardi_reports.py`
3. ✅ Verify files appear in Dashboard Data folder
4. ⏳ (Optional) Add to Task Scheduler for automatic daily downloads
5. ⏳ Create extraction scripts to parse downloaded reports
6. ⏳ Feed extracted data into dashboard metrics

