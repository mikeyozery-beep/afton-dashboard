# Download YARDI attachments from Outlook COM
# Uses ONLY the shared mailbox mikereports@aftonprop.com.
#
# Selection is driven by a PROCESSED LEDGER (processed_emails.txt), not by
# read/unread state. We consider every report email from the sender received in
# the last $WindowDays days and skip any whose EntryID is already in the ledger.
# This fixes the old "someone opened the email before 5 AM so it was skipped"
# silent miss - read state no longer decides what gets downloaded.
#
# An email's EntryID is added to the ledger as soon as its attachments are pulled
# to disk, so a later organize/mark-read failure can never cause the same report
# to be downloaded and filed twice.
#
# First run (ledger file absent): seed the ledger with the EntryIDs of every
# already-READ sender email in the window so the historical backlog is treated as
# processed - avoids a one-time duplicate storm - and only the unread ones are
# pulled. From then on the recency clause is active.
#
# Only spreadsheet attachments (.xlsx/.xlsm/.xls) and .zip archives are saved;
# the organize step only understands workbooks, so csv/txt/pdf are intentionally
# NOT downloaded (they would otherwise pile up unfiled and the email would be
# consumed). Records which emails to mark read into pending_mark_read.txt; the
# actual read flip happens in mark_read_pending.ps1 AFTER organize succeeds.

Add-Type -AssemblyName System.IO.Compression

$OutputDir = "C:\Users\MichaelOzery\OneDrive - Afton Properties\Old Dropbox\My PC (DESKTOP-5D77V89)\Mike Ozery\Dashboard Data"
$SenderEmail = "cdr@yardi.com"
$TargetMailbox = "mikereports@aftonprop.com"
$PendingFile = Join-Path $PSScriptRoot "pending_mark_read.txt"
$LedgerFile = Join-Path $PSScriptRoot "processed_emails.txt"
$WindowDays = 4          # how far back to look for not-yet-processed report emails
$LedgerCap = 2000        # keep the ledger bounded (>> the recency window)

$timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
Write-Host "[$timestamp] Starting YARDI attachment download (ledger-based selection)..."
Write-Host "[$timestamp] Target mailbox: $TargetMailbox ONLY"
Write-Host ""

if (-not (Test-Path $OutputDir)) {
    New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null
}

# --- Load the processed-email ledger -------------------------------------------
$ledger = New-Object System.Collections.Generic.HashSet[string]
$ledgerExists = Test-Path $LedgerFile
if ($ledgerExists) {
    foreach ($line in (Get-Content $LedgerFile)) {
        $t = $line.Trim()
        if ($t -ne "") { [void]$ledger.Add($t) }
    }
}
Write-Host "[$timestamp] Processed ledger: $($ledger.Count) email(s) known (exists=$ledgerExists)"

try {
    $outlook = New-Object -ComObject Outlook.Application
    $mapi = $outlook.GetNamespace("MAPI")
    Write-Host "[$timestamp] Connected to Outlook"

    # Find ONLY the shared mailbox (Mike Reports / mikereports@aftonprop.com)
    $inbox = $null
    $storeID = $null
    foreach ($store in $mapi.Stores) {
        $folderName = $store.DisplayName
        Write-Host "[$timestamp] Checking mailbox: $folderName"
        if ($folderName -like "Mike*Reports*" -or $folderName -eq "Mike Reports" -or $folderName -like "*mikereports*") {
            try {
                $inbox = $store.GetRootFolder().Folders("Inbox")
                $storeID = $store.StoreID
                Write-Host "[$timestamp] FOUND TARGET: $folderName"
                break
            }
            catch {
                Write-Host "[$timestamp] Error accessing $folderName, skipping..."
                continue
            }
        }
    }

    if (-not $inbox) {
        Write-Host "[$timestamp] ERROR: Could not find target mailbox"
        exit 1
    }

    # --- Build the candidate set ----------------------------------------------
    # Recency-restricted view of the inbox, plus the unread view as a backstop.
    $cutoff = (Get-Date).AddDays(-$WindowDays).ToString("MM/dd/yyyy HH:mm")
    $recentItems = $inbox.Items.Restrict("[ReceivedTime] >= '$cutoff'")
    $unreadItems = $inbox.Items.Restrict("[UnRead] = true")

    $seen = New-Object System.Collections.Generic.HashSet[string]
    $candidates = New-Object System.Collections.Generic.List[object]

    function Test-IsSender($item) {
        try { return ($item.SenderEmailAddress -like "*$SenderEmail*") } catch { return $false }
    }

    if (-not $ledgerExists) {
        # FIRST RUN: treat the existing read backlog as already-processed.
        Write-Host "[$timestamp] First run - seeding ledger from already-read mail in the window..."
        foreach ($item in $recentItems) {
            if (Test-IsSender $item) {
                try { if (-not $item.UnRead) { [void]$ledger.Add($item.EntryID) } } catch {}
            }
        }
        # Only the unread sender emails are actually new work this first run.
        foreach ($item in $unreadItems) {
            if (-not (Test-IsSender $item)) { continue }
            $id = $item.EntryID
            if ($ledger.Contains($id) -or $seen.Contains($id)) { continue }
            [void]$seen.Add($id); $candidates.Add($item)
        }
    }
    else {
        foreach ($coll in @($unreadItems, $recentItems)) {
            foreach ($item in $coll) {
                if (-not (Test-IsSender $item)) { continue }
                $id = $item.EntryID
                if ($ledger.Contains($id) -or $seen.Contains($id)) { continue }
                [void]$seen.Add($id); $candidates.Add($item)
            }
        }
    }

    $downloaded = 0
    $extracted = 0
    $skipped = 0
    $processed = 0
    $pending = New-Object System.Collections.Generic.List[string]
    $newlyProcessed = New-Object System.Collections.Generic.List[string]

    Write-Host "[$timestamp] Window: last $WindowDays day(s) (since $cutoff)"
    Write-Host "[$timestamp] Candidate report emails to process: $($candidates.Count)"
    Write-Host "[$timestamp] Processing emails from $SenderEmail..."
    Write-Host ""

    foreach ($item in $candidates) {
        if ($item.Attachments.Count -eq 0) {
            # Nothing to pull, but record it so we don't re-examine it forever.
            $newlyProcessed.Add($item.EntryID)
            continue
        }

        $processed++
        $emailDate = $item.ReceivedTime.ToString("yyyy-MM-dd HH:mm:ss")
        $savedFromThisEmail = $false
        $emailHadError = $false

        foreach ($attachment in $item.Attachments) {
            $filename = $attachment.Filename
            $tempPath = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), $filename)

            $attachment.SaveAsFile($tempPath)

            # Handle zip files
            if ($filename -like "*.zip") {
                Write-Host "[$timestamp] [Email $processed] Extracting zip: $filename ($emailDate)"
                try {
                    $zip = [System.IO.Compression.ZipFile]::OpenRead($tempPath)
                    foreach ($entry in $zip.Entries) {
                        if ($entry.Length -gt 0) {
                            $extractPath = Join-Path $OutputDir $entry.Name
                            if (Test-Path $extractPath) {
                                $baseName = [System.IO.Path]::GetFileNameWithoutExtension($entry.Name)
                                $ext = [System.IO.Path]::GetExtension($entry.Name)
                                $dateForFile = $emailDate -replace ':', '-'
                                $counter = 1
                                while (Test-Path $extractPath) {
                                    $newName = "$baseName`_$dateForFile`_$counter$ext"
                                    $extractPath = Join-Path $OutputDir $newName
                                    $counter++
                                }
                            }
                            [System.IO.Compression.ZipFileExtensions]::ExtractToFile($entry, $extractPath, $true)
                            Write-Host "[$timestamp]     [OK]Extracted: $($entry.Name)"
                            $extracted++
                            $savedFromThisEmail = $true
                        }
                    }
                    $zip.Dispose()
                }
                catch {
                    Write-Host "[$timestamp]     [X]Error extracting: $_"
                    $emailHadError = $true
                }
                Remove-Item $tempPath -Force -ErrorAction SilentlyContinue
                continue
            }

            # Only spreadsheet reports - organize only understands workbooks.
            if ($filename -notmatch '\.(xlsx|xlsm|xls)$') {
                Write-Host "[$timestamp] [Email $processed] Skipping (not a workbook): $filename"
                Remove-Item $tempPath -Force -ErrorAction SilentlyContinue
                continue
            }

            $finalPath = Join-Path $OutputDir $filename
            if (Test-Path $finalPath) {
                $existingDate = (Get-Item $finalPath).LastWriteTime.ToString("yyyy-MM-dd HH:mm")
                if ($existingDate -eq $emailDate.Substring(0, 16)) {
                    Write-Host "[$timestamp] [Email $processed] Skipping (exact duplicate): $filename [$emailDate]"
                    $skipped++
                    # Already on disk from a prior run -> still safe to mark this email read
                    $savedFromThisEmail = $true
                    Remove-Item $tempPath -Force -ErrorAction SilentlyContinue
                    continue
                }
                else {
                    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($filename)
                    $ext = [System.IO.Path]::GetExtension($filename)
                    $dateForFile = $emailDate -replace ':', '-'
                    $newName = "$baseName`_$dateForFile$ext"
                    $finalPath = Join-Path $OutputDir $newName
                    Write-Host "[$timestamp] [Email $processed] Saving new version: $newName"
                }
            }

            try {
                Move-Item $tempPath $finalPath -Force
                Write-Host "[$timestamp] [Email $processed] [OK]Downloaded: $filename ($emailDate)"
                $downloaded++
                $savedFromThisEmail = $true
            }
            catch {
                Write-Host "[$timestamp] [Email $processed] [X]Error: $_"
                $emailHadError = $true
                Remove-Item $tempPath -Force -ErrorAction SilentlyContinue
            }
        }

        # Record this email as processed UNLESS a save/extract error occurred - on
        # error we leave it out of the ledger so the next run retries it. Queue it
        # to be marked read (after organize) when something was actually saved.
        if (-not $emailHadError) {
            $newlyProcessed.Add($item.EntryID)
            if ($savedFromThisEmail) { $pending.Add($item.EntryID) }
        }
    }

    # --- Persist the ledger immediately (download-time durability) -------------
    if ($newlyProcessed.Count -gt 0 -or -not $ledgerExists) {
        foreach ($id in $newlyProcessed) { [void]$ledger.Add($id) }
        $all = @($ledger)
        if ($all.Count -gt $LedgerCap) { $all = $all[($all.Count - $LedgerCap)..($all.Count - 1)] }
        Set-Content -Path $LedgerFile -Value $all -Encoding UTF8
    }

    # --- Write the pending-mark-read queue (consumed after organize succeeds) ---
    if ($pending.Count -gt 0) {
        $lines = New-Object System.Collections.Generic.List[string]
        $lines.Add("STORE=$storeID")
        foreach ($id in $pending) { $lines.Add($id) }
        Set-Content -Path $PendingFile -Value $lines -Encoding UTF8
    }
    else {
        # Nothing to mark; clear any stale queue
        if (Test-Path $PendingFile) { Remove-Item $PendingFile -Force }
    }

    Write-Host ""
    Write-Host "[$timestamp] =============================================="
    Write-Host "[$timestamp] DOWNLOAD COMPLETE"
    Write-Host "[$timestamp] =============================================="
    Write-Host "[$timestamp] Downloaded: $downloaded files"
    Write-Host "[$timestamp] Extracted from zips: $extracted files"
    Write-Host "[$timestamp] Skipped (exact duplicates): $skipped files"
    Write-Host "[$timestamp] Processed: $processed candidate emails"
    Write-Host "[$timestamp] Recorded in ledger this run: $($newlyProcessed.Count)"
    Write-Host "[$timestamp] Queued to mark read (after organize): $($pending.Count)"
    Write-Host "[$timestamp] Saved to: $OutputDir"
    Write-Host "[$timestamp] =============================================="
}
catch {
    Write-Host "FATAL ERROR: $_"
    exit 1
}

exit 0
