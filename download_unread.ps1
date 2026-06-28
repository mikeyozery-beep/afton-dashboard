# One-off: download report attachments from UNREAD emails in the
# mikereports@aftonprop.com shared mailbox inbox.
# Does NOT change read/unread state. Exact-duplicate files are skipped.

Start-Transcript -Path "C:\Afton\Dashboard\logs\unread_pull.txt" -Force | Out-Null

Add-Type -AssemblyName System.IO.Compression

$OutputDir = "C:\Users\MichaelOzery\OneDrive - Afton Properties\Old Dropbox\My PC (DESKTOP-5D77V89)\Mike Ozery\Dashboard Data"

$timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
Write-Host "[$timestamp] Downloading attachments from UNREAD emails in mikereports@aftonprop.com..."
Write-Host ""

if (-not (Test-Path $OutputDir)) {
    New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null
}

try {
    $outlook = New-Object -ComObject Outlook.Application
    $mapi = $outlook.GetNamespace("MAPI")
    Write-Host "[$timestamp] Connected to Outlook"

    # Locate the shared mailbox inbox
    $inbox = $null
    foreach ($store in $mapi.Stores) {
        $folderName = $store.DisplayName
        if ($folderName -like "Mike*Reports*" -or $folderName -eq "Mike Reports" -or $folderName -like "*mikereports*") {
            try {
                $inbox = $store.GetRootFolder().Folders("Inbox")
                Write-Host "[$timestamp] FOUND TARGET: $folderName"
                break
            }
            catch { continue }
        }
    }

    if (-not $inbox) {
        Write-Host "[$timestamp] ERROR: Could not find target mailbox"
        exit 1
    }

    # Restrict to unread messages only
    $unread = $inbox.Items.Restrict("[UnRead] = true")
    Write-Host "[$timestamp] Unread messages in inbox: $($unread.Count)"
    Write-Host ""

    $downloaded = 0
    $extracted = 0
    $skipped = 0
    $processed = 0

    foreach ($item in $unread) {
        if ($item.Attachments.Count -eq 0) { continue }

        $processed++
        $emailDate = $item.ReceivedTime.ToString("yyyy-MM-dd HH:mm:ss")
        $senderAddr = $item.SenderEmailAddress
        Write-Host "[$timestamp] [Unread $processed] $($item.Subject) | $senderAddr | $emailDate"

        foreach ($attachment in $item.Attachments) {
            $filename = $attachment.Filename
            $tempPath = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), $filename)
            $attachment.SaveAsFile($tempPath)

            # Zip files: extract report members
            if ($filename -like "*.zip") {
                Write-Host "[$timestamp]   Extracting zip: $filename"
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
                            Write-Host "[$timestamp]     Extracted: $($entry.Name)"
                            $extracted++
                        }
                    }
                    $zip.Dispose()
                }
                catch { Write-Host "[$timestamp]     Error extracting: $_" }
                Remove-Item $tempPath -Force
                continue
            }

            # Only report-type files
            if ($filename -notmatch '\.(xlsx|xls|csv|txt|pdf)$') {
                Remove-Item $tempPath -Force
                continue
            }

            $finalPath = Join-Path $OutputDir $filename
            if (Test-Path $finalPath) {
                $existingDate = (Get-Item $finalPath).LastWriteTime.ToString("yyyy-MM-dd HH:mm")
                if ($existingDate -eq $emailDate.Substring(0, 16)) {
                    Write-Host "[$timestamp]   Skipping (exact duplicate): $filename"
                    $skipped++
                    Remove-Item $tempPath -Force
                    continue
                }
                else {
                    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($filename)
                    $ext = [System.IO.Path]::GetExtension($filename)
                    $dateForFile = $emailDate -replace ':', '-'
                    $finalPath = Join-Path $OutputDir "$baseName`_$dateForFile$ext"
                }
            }

            try {
                Move-Item $tempPath $finalPath -Force
                Write-Host "[$timestamp]   Downloaded: $filename"
                $downloaded++
            }
            catch {
                Write-Host "[$timestamp]   Error: $_"
                Remove-Item $tempPath -Force -ErrorAction SilentlyContinue
            }
        }
    }

    Write-Host ""
    Write-Host "[$timestamp] =============================================="
    Write-Host "[$timestamp] UNREAD DOWNLOAD COMPLETE"
    Write-Host "[$timestamp] Unread emails with attachments: $processed"
    Write-Host "[$timestamp] Downloaded: $downloaded | Extracted: $extracted | Skipped dup: $skipped"
    Write-Host "[$timestamp] (Read/unread state was NOT changed)"
    Write-Host "[$timestamp] =============================================="
}
catch {
    Write-Host "FATAL ERROR: $_"
    Stop-Transcript | Out-Null
    exit 1
}

Stop-Transcript | Out-Null
exit 0
