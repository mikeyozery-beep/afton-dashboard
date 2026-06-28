# Report Scheduler Email Extractor

Automated pipeline that pulls YARDI Scheduler report attachments from the
**mikereports@aftonprop.com** shared Outlook mailbox, organizes them into
per-report folders, and marks the source emails read once they're filed.

Part of the Afton Properties portfolio dashboard system.

## How it works

A Windows scheduled task (**"YARDI Report Downloader - Daily 5 AM"**) runs
`unified_download_organize.py`, which executes four gated steps:

| Step | Script | Action |
|------|--------|--------|
| 1. Download | `download_yardi_outlook_com.ps1` | Saves attachments from **unread** `cdr@yardi.com` emails to the Dashboard Data folder; queues their email IDs in `pending_mark_read.txt`. Does **not** change read state. |
| 2. Extract | (built in) | Unzips any `.zip` attachments. |
| 3. Organize | `organize_reports.py` logic | Splits multi-sheet workbooks and files each report into a folder named after its report title (cell `A1`). |
| 4. Mark read | `mark_read_pending.ps1` | **Only if step 3 succeeded**, marks the queued emails read and clears the queue. If marking fails, emails stay unread and retry next run. |

"Unread" = "not yet processed." An email is marked read strictly after its
attachments are both saved **and** filed into the right folder.

## Key files

- `unified_download_organize.py` — orchestrator (the scheduled task runs this)
- `download_yardi_outlook_com.ps1` — unread-email downloader + read-queue writer
- `mark_read_pending.ps1` — marks queued emails read (scheduler step 4)
- `mark_read_now.ps1` — one-off: mark already-processed emails read
- `organize_reports.py` — standalone organize/file step
- `fix_download_task.ps1` — points the scheduled tasks at the real `python.exe`

## Output location

`C:\Users\MichaelOzery\OneDrive - Afton Properties\Old Dropbox\My PC (DESKTOP-5D77V89)\Mike Ozery\Dashboard Data\`
— one folder per report type, files timestamped.

## Requirements

- Windows with Outlook running, logged in, with the **Mike Reports** shared
  mailbox loaded (download uses Outlook COM, so it needs an interactive session).
- Python (`pywin32`, `openpyxl`). Use the real interpreter at
  `C:\Users\MichaelOzery\AppData\Local\Python\bin\python.exe`, **not** the
  Windows Store `python` alias (it doesn't resolve under Task Scheduler).

## Notes

- Selection is unread-based: if someone opens a report email in the shared
  mailbox before 5 AM, Outlook marks it read and that run will skip it.
- Logs: `logs\unified_*.log`.
