# Bulk Download YARDI Reports from mikereports@aftonprop.com Outlook
# PowerShell approach for more reliable Outlook access
# CRITICAL: Uses mikereports@aftonprop.com ONLY - never primary email

param(
    [switch]$DownloadAll = $true
)

# Configuration
$OutputDir = "C:\Users\MichaelOzery\OneDrive - Afton Properties\Old Dropbox\My PC (DESKTOP-5D77V89)\Mike Ozery\Dashboard Data"
$TargetMailbox = "mikereports@aftonprop.com"
$SenderEmail = "cdr@yardi.com"
$LogFile = "$PSScriptRoot\logs\yardi_powershell_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"

# Create output directory
if (-not (Test-Path $OutputDir)) {
    New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null
}

# Create logs directory
$LogDir = "$PSScriptRoot\logs"
if (-not (Test-Path $LogDir)) {
    New-Item -ItemType Directory -Path $LogDir -Force | Out-Null
}

function Write-Log {
    param([string]$Message)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $output = "$timestamp - $Message"
    Write-Host $output
    Add-Content -Path $LogFile -Value $output
}

Write-Log "========================================================================"
Write-Log "BULK DOWNLOAD YARDI ATTACHMENTS (PowerShell)"
Write-Log "CRITICAL: Using $TargetMailbox mailbox ONLY"
Write-Log "========================================================================"

$Outlook = $null

try {
    Write-Log ""
    Write-Log "[OUTLOOK CONNECTION]"
    Write-Log "Connecting to Outlook..."

    # Create Outlook COM object
    $Outlook = New-Object -ComObject Outlook.Application
    Write-Log "  Successfully created Outlook COM object"

    $Namespace = $Outlook.GetNamespace("MAPI")
    Write-Log "  Got MAPI namespace"

    # List all stores to find the target mailbox
    Write-Log "  Searching for mailbox: $TargetMailbox"
    Write-Log "  Available mailboxes:"

    $TargetInbox = $null

    foreach ($Store in $Namespace.Stores) {
        $StoreName = $Store.DisplayName
        Write-Log "    - $StoreName"

        if ($StoreName -like "*$TargetMailbox*" -or $StoreName -like "*mikereports*" -or $StoreName -like "*aftonprop*") {
            Write-Log "    [MATCH] Found target mailbox!"
            try {
                $RootFolder = $Store.GetRootFolder()
                $TargetInbox = $RootFolder.Folders.Item("Inbox")
                Write-Log "  [OK] Connected to: $StoreName > Inbox"
                break
            }
            catch {
                Write-Log "    Error getting inbox: $_"
                continue
            }
        }
    }

    if ($null -eq $TargetInbox) {
        Write-Log "[ERROR] Could not find mailbox for $TargetMailbox"
        Write-Log "  Make sure this mailbox is configured in Outlook"
        exit 1
    }

    # Process emails
    Write-Log ""
    Write-Log "Searching for emails from $SenderEmail..."

    $Items = $TargetInbox.Items
    $Items.Sort("[ReceivedTime]", $false)

    Write-Log "Total emails in inbox: $($Items.Count)"
    Write-Log "Processing emails..."

    $Downloaded = 0
    $Skipped = 0
    $Errors = 0
    $Processed = 0

    foreach ($Item in $Items) {
        try {
            # Check sender
            if ($Item.SenderEmailAddress -notlike "*$SenderEmail*") {
                continue
            }

            # Check attachments
            if ($Item.Attachments.Count -eq 0) {
                continue
            }

            $Processed++
            $EmailDate = $Item.ReceivedTime
            $DateStr = $EmailDate.ToString("MM.dd.yyyy")

            Write-Log ""
            Write-Log "[Email $Processed] $($EmailDate.ToString('yyyy-MM-dd HH:mm:ss')) - $($Item.Subject)"
            Write-Log "    Attachments: $($Item.Attachments.Count)"

            foreach ($Attachment in $Item.Attachments) {
                try {
                    $Filename = $Attachment.Filename

                    # Only download Excel files
                    if ($Filename -notmatch '\.(xlsx|xls)$') {
                        Write-Log "    Skipping (not Excel): $Filename"
                        continue
                    }

                    $FilePath = Join-Path $OutputDir $Filename

                    # Check if already exists
                    if (Test-Path $FilePath) {
                        Write-Log "    Skipping (already exists): $Filename"
                        $Skipped++
                        continue
                    }

                    # Download
                    $Attachment.SaveAsFile($FilePath)
                    Write-Log "    [OK] Downloaded: $Filename"
                    $Downloaded++

                }
                catch {
                    Write-Log "    Error processing attachment: $_"
                    $Errors++
                }
            }

        }
        catch {
            Write-Log "  Error processing email: $_"
            $Errors++
        }
    }

    Write-Log ""
    Write-Log "========================================================================"
    Write-Log "DOWNLOAD COMPLETE"
    Write-Log "========================================================================"
    Write-Log "Downloaded: $Downloaded attachments"
    Write-Log "Skipped (already exist): $Skipped"
    Write-Log "Errors: $Errors"
    Write-Log "Processed: $Processed YARDI emails"
    Write-Log "Saved to: $OutputDir"
    Write-Log "========================================================================"

    if ($Errors -eq 0) {
        exit 0
    }
    else {
        exit 1
    }

}
catch {
    Write-Log "Fatal error: $_"
    Write-Log $_.Exception.StackTrace
    exit 1
}
finally {
    if ($Outlook) {
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Outlook) | Out-Null
    }
}
