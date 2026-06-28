# Download YARDI attachments from Outlook COM
# Uses ONLY the shared mailbox mikereports@aftonprop.com
# Selects UNREAD report emails (unread = not yet processed).
# Records which emails were downloaded into pending_mark_read.txt so they can be
# marked READ *after* the organize step succeeds (see mark_read_pending.ps1).
# Does NOT change read/unread state itself. Extracts zips; skips exact duplicates.

Add-Type -AssemblyName System.IO.Compression

$OutputDir = "C:\Users\MichaelOzery\OneDrive - Afton Properties\Old Dropbox\My PC (DESKTOP-5D77V89)\Mike Ozery\Dashboard Data"
$SenderEmail = "cdr@yardi.com"
$TargetMailbox = "mikereports@aftonprop.com"
$PendingFile = Join-Path $PSScriptRoot "pending_mark_read.txt"

$timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
Write-Host "[$timestamp] Starting YARDI attachment download (UNREAD emails only)..."
Write-Host "[$timestamp] Target mailbox: $TargetMailbox ONLY"
Write-Host ""

if (-not (Test-Path $OutputDir)) {
    New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null
}

try {
    $outlook = New-Object -ComObject Outlook.Application
    $mapi = $outlook.GetNamespace("MAPI")
    Write-Host "[$timestamp] Connected to Outlook"

    # Find ONLY the shared mailbox (Mike Reports / mikereports@aftonprop.com)
    $inbox = $null
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

    # Only UNREAD messages (unread = not yet processed)
    $items = $inbox.Items.Restrict("[UnRead] = true")

    $downloaded = 0
    $extracted = 0
    $skipped = 0
    $processed = 0
    $pending = New-Object System.Collections.Generic.List[string]

    Write-Host "[$timestamp] Unread messages in inbox: $($items.Count)"
    Write-Host "[$timestamp] Processing UNREAD emails from $SenderEmail..."
    Write-Host ""

    foreach ($item in $items) {
        if ($item.SenderEmailAddress -notlike "*$SenderEmail*") {
            continue
        }
        if ($item.Attachments.Count -eq 0) {
            continue
        }

        $processed++
        $emailDate = $item.ReceivedTime.ToString("yyyy-MM-dd HH:mm:ss")
        $savedFromThisEmail = $false

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
                }
                Remove-Item $tempPath -Force
                continue
            }

            # Only report-type files
            if ($filename -notmatch '\.(xlsx|xls|csv|txt|pdf)$') {
                Write-Host "[$timestamp] [Email $processed] Skipping (not a report): $filename"
                Remove-Item $tempPath -Force
                continue
            }

            $finalPath = Join-Path $OutputDir $filename
            if (Test-Path $finalPath) {
                $existingDate = (Get-Item $finalPath).LastWriteTime.ToString("yyyy-MM-dd HH:mm")
                if ($existingDate -eq $emailDate.Substring(0, 16)) {
                    Write-Host "[$timestamp] [Email $processed] Skipping (exact duplicate): $filename [$emailDate]"
                    $skipped++
                    # Already downloaded previously -> still safe to mark this email read
                    $savedFromThisEmail = $true
                    Remove-Item $tempPath -Force
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
                Remove-Item $tempPath -Force -ErrorAction SilentlyContinue
            }
        }

        # Queue this email to be marked read AFTER organize succeeds
        if ($savedFromThisEmail) {
            $pending.Add($item.EntryID)
        }
    }

    # Write the pending-mark-read queue (EntryIDs + the store they live in).
    # mark_read_pending.ps1 consumes this only after the organize step succeeds.
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
    Write-Host "[$timestamp] Processed: $processed unread emails"
    Write-Host "[$timestamp] Queued to mark read (after organize): $($pending.Count)"
    Write-Host "[$timestamp] Saved to: $OutputDir"
    Write-Host "[$timestamp] =============================================="
}
catch {
    Write-Host "FATAL ERROR: $_"
    exit 1
}

exit 0
