# ONE-OFF: mark the already-extracted unread report emails as read.
# Targets unread emails in the Mike Reports inbox that have attachments and are
# from the YARDI report sender (cdr@yardi.com) - i.e. the ones already downloaded
# and organized in this session. Read/unread is the only thing changed.

Start-Transcript -Path "C:\Afton\Dashboard\logs\mark_read_now.txt" -Force | Out-Null

$SenderEmail = "cdr@yardi.com"
$ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
Write-Host "[$ts] Marking processed unread report emails as read..."

try {
    $outlook = New-Object -ComObject Outlook.Application
    $mapi = $outlook.GetNamespace("MAPI")

    $inbox = $null
    foreach ($store in $mapi.Stores) {
        $n = $store.DisplayName
        if ($n -like "Mike*Reports*" -or $n -eq "Mike Reports" -or $n -like "*mikereports*") {
            try { $inbox = $store.GetRootFolder().Folders("Inbox"); break } catch { continue }
        }
    }
    if (-not $inbox) { Write-Host "[$ts] ERROR: target mailbox not found"; Stop-Transcript | Out-Null; exit 1 }

    $unread = $inbox.Items.Restrict("[UnRead] = true")
    Write-Host "[$ts] Unread before: $($unread.Count)"

    # Collect first (marking read removes items from the restricted collection mid-loop)
    $toMark = @()
    foreach ($item in $unread) {
        if ($item.Attachments.Count -gt 0 -and $item.SenderEmailAddress -like "*$SenderEmail*") {
            $toMark += $item
        }
    }

    $marked = 0
    foreach ($item in $toMark) {
        try { $item.UnRead = $false; $item.Save(); $marked++ }
        catch { Write-Host "[$ts]   Error marking one: $_" }
    }

    $after = $inbox.Items.Restrict("[UnRead] = true")
    Write-Host "[$ts] Marked read: $marked"
    Write-Host "[$ts] Unread after: $($after.Count)"
}
catch {
    Write-Host "[$ts] FATAL: $_"
    Stop-Transcript | Out-Null
    exit 1
}

Stop-Transcript | Out-Null
exit 0
