# Report Scheduler Email Extractor

Automation that downloads YARDI report attachments from the
`mikereports@aftonprop.com` shared Outlook mailbox, organizes them into
per-report folders, and marks the source emails read once filed. See `README.md`
for the full pipeline description.

## Entry point
- Scheduled task **"YARDI Report Downloader - Daily 5 AM"** runs
  `unified_download_organize.py` (download → extract → organize → mark read).
- Read-marking is gated on a successful organize step. Emails that aren't fully
  processed stay UNREAD and are retried next run.

## Operational gotchas (learned the hard way)
- **Python path:** scheduled tasks must use the real interpreter
  `C:\Users\MichaelOzery\AppData\Local\Python\bin\python.exe`. The `python` on
  PATH is the Windows Store alias (0-byte stub) and fails under Task Scheduler
  with `0x80070002`.
- **Outlook COM needs a non-elevated, interactive session.** An elevated shell
  cannot attach to the user's Outlook (UAC isolation). To run an Outlook script
  from an elevated/automation context, register a temporary scheduled task with
  `LogonType Interactive, RunLevel Limited` as user `MichaelOzery`, start it,
  read its transcript log, then unregister it.
- Tasks run `RunLevel Limited` (non-elevated) on purpose so they can read Outlook.

## Selection model
- "Unread" = "not yet processed." Selection is unread-based, so manually opening
  a report email before the 5 AM run will cause that run to skip it.

## Never commit
Credentials/tokens are gitignored (`credentials*.json`, `tokens/`,
`o365_tokens/`, `OAUTH_CREDENTIALS_REFERENCE.md`). The git remote URL has
historically contained a plaintext PAT — keep tokens out of committed files.
