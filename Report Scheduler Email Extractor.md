# Report Scheduler Email Extractor

Automated pipeline that pulls YARDI Scheduler report attachments from the
**mikereports@aftonprop.com** shared Outlook mailbox, organizes them into
per-report folders, and marks the source emails read once they're filed.

Part of the Afton Properties portfolio dashboard system.
Lives in `C:\Afton\Dashboard` (folder name kept as-is — scheduled tasks and
scripts hardcode this path).

---

## What it does

Every morning a Windows scheduled task pulls the overnight YARDI reports out of
the shared mailbox and files them, with no manual steps. "Unread" in the mailbox
means "not yet processed" — an email is only marked read after its attachments
are saved **and** filed into the correct folder.

---

## Pipeline

Scheduled task **"YARDI Report Downloader - Daily 5 AM"** runs
`unified_download_organize.py`, which executes four gated steps:

| Step | Script | Action |
|------|--------|--------|
| 1. Download | `download_yardi_outlook_com.ps1` | Save attachments from **unread** `cdr@yardi.com` emails to the Dashboard Data folder; queue their email IDs in `pending_mark_read.txt`. Does **not** change read state. Extracts `.zip` attachments. Skips exact duplicates. |
| 2. Extract | (built into the orchestrator) | Unzip any remaining `.zip` files. |
| 3. Organize | organize logic (also in `organize_reports.py`) | Split multi-sheet workbooks; file each report into a folder named after its report title (cell `A1`), timestamped. |
| 4. Mark read | `mark_read_pending.ps1` | **Only if step 3 succeeded**, mark the queued emails read and clear the queue. If marking fails, emails stay unread and retry next run. |

---

## Scheduled tasks

| Task | Time | Runs | Purpose |
|------|------|------|---------|
| YARDI Report Downloader - Daily 5 AM | 5:00 AM | `unified_download_organize.py` | Download + organize + mark read |
| Afton Dashboard - Daily | 7:00 AM | `extraction_system\main_extractor.py` | Extract metrics from filed reports |
| Afton Dashboard - Daily Metrics Extract | 9:00 AM | `extract_and_push_to_github.py` | Push metrics to GitHub / dashboard |

All run **non-elevated** (`RunLevel Limited`) as user `MichaelOzery` so they can
read Outlook, using the real Python interpreter (see gotchas).

---

## Key files

- `unified_download_organize.py` — orchestrator (the 5 AM task runs this)
- `download_yardi_outlook_com.ps1` — unread-email downloader + read-queue writer
- `mark_read_pending.ps1` — marks queued emails read (scheduler step 4)
- `mark_read_now.ps1` — one-off: mark already-processed emails read
- `download_unread.ps1` / `run_unread.bat` — one-off: pull attachments from all unread emails
- `organize_reports.py` — standalone organize/file step
- `fix_download_task.ps1` — points the scheduled tasks at the real `python.exe`
- `README.md`, `CLAUDE.md` — project docs

---

## Output location

```
C:\Users\MichaelOzery\OneDrive - Afton Properties\Old Dropbox\My PC (DESKTOP-5D77V89)\Mike Ozery\Dashboard Data\
```
One folder per report type (e.g. `Rent Roll`, `Aged Receivables`,
`rs_sql_income_statement_all_muvan`), files timestamped.

---

## Requirements

- Windows with **Outlook running**, logged in, and the **Mike Reports** shared
  mailbox loaded. Download uses Outlook COM, so it needs an interactive session.
- Python with `pywin32` and `openpyxl`. Use the real interpreter at
  `C:\Users\MichaelOzery\AppData\Local\Python\bin\python.exe`.

---

## Operational gotchas (learned the hard way)

- **Python path:** scheduled tasks must use the real interpreter, not the
  `python` on PATH — that one is the Windows Store alias (a 0-byte stub) and
  fails under Task Scheduler with `0x80070002` (file not found).
- **Outlook COM needs a non-elevated, interactive session.** An elevated shell
  cannot attach to the user's Outlook (UAC isolation). To run an Outlook script
  from an elevated/automation context, register a temporary scheduled task with
  `LogonType Interactive, RunLevel Limited` as user `MichaelOzery`, start it,
  read its transcript log, then unregister it.
- **Unread-based selection:** if someone opens a report email in the shared
  mailbox before 5 AM, Outlook marks it read and that run will skip it.

---

## Run manually

```powershell
# Full pipeline (download + organize + mark read)
& "C:\Users\MichaelOzery\AppData\Local\Python\bin\python.exe" "C:\Afton\Dashboard\unified_download_organize.py"

# Or trigger the scheduled task
Start-ScheduledTask -TaskName "YARDI Report Downloader - Daily 5 AM"

# Organize only (no download)
& "C:\Users\MichaelOzery\AppData\Local\Python\bin\python.exe" "C:\Afton\Dashboard\organize_reports.py"
```

Logs: `C:\Afton\Dashboard\logs\unified_*.log`

---

## Troubleshooting

| Symptom | Likely cause / fix |
|---------|--------------------|
| Task result `0x80070002` | Task using the Store `python` alias — run `fix_download_task.ps1` |
| "Could not find target mailbox" | Outlook not running, or the Mike Reports shared mailbox isn't loaded in the profile |
| COM/RPC error when run manually | You're in an elevated shell — use a non-elevated interactive scheduled task instead |
| Reports re-downloaded repeatedly | Emails not getting marked read — check step 4 / `pending_mark_read.txt` |
| Slow organize | Large workbooks load twice per sheet; normal for big Rent Rolls |

---

## Security

Credentials and tokens are gitignored (`credentials*.json`, `tokens/`,
`o365_tokens/`, `OAUTH_CREDENTIALS_REFERENCE.md`). Keep tokens out of committed
files; the git remote URL has historically embedded a plaintext PAT — rotate it
and store it via a credential manager rather than in the remote URL.
