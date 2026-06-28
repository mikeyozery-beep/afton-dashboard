# Mark as READ the emails queued by download_yardi_outlook_com.ps1.
# Run ONLY after the organize step has succeeded, so an email is marked read
# strictly once its attachments are saved AND filed into the right folders.
# Consumes pending_mark_read.txt (first line: STORE=<storeID>, then one EntryID per line),
# then deletes it.

$PendingFile = Join-Path $PSScriptRoot "pending_mark_read.txt"
$ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

if (-not (Test-Path $PendingFile)) {
    Write-Host "[$ts] No pending emails to mark read. Nothing to do."
    exit 0
}

$lines = Get-Content $PendingFile | Where-Object { $_.Trim() -ne "" }
$storeID = $null
$entryIDs = New-Object System.Collections.Generic.List[string]
foreach ($line in $lines) {
    if ($line -like "STORE=*") { $storeID = $line.Substring(6) }
    else { $entryIDs.Add($line.Trim()) }
}

if (-not $storeID -or $entryIDs.Count -eq 0) {
    Write-Host "[$ts] Pending file present but no usable entries. Removing it."
    Remove-Item $PendingFile -Force
    exit 0
}

Write-Host "[$ts] Marking $($entryIDs.Count) processed email(s) as read..."

try {
    $outlook = New-Object -ComObject Outlook.Application
    $mapi = $outlook.GetNamespace("MAPI")

    $marked = 0
    $errors = 0
    foreach ($id in $entryIDs) {
        try {
            $item = $mapi.GetItemFromID($id, $storeID)
            if ($item.UnRead) {
                $item.UnRead = $false
                $item.Save()
            }
            $marked++
        }
        catch {
            Write-Host "[$ts]   Could not mark one email (it may have moved/been deleted): $_"
            $errors++
        }
    }

    Write-Host "[$ts] Marked read: $marked  | Errors: $errors"

    # Clear the queue so the same emails aren't re-marked next run
    Remove-Item $PendingFile -Force
    Write-Host "[$ts] Pending queue cleared."
}
catch {
    Write-Host "[$ts] FATAL: $_"
    # Leave the pending file in place so a later run can retry
    exit 1
}

exit 0
