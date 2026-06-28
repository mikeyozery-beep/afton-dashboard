# Download YARDI attachments from shared mailbox using Exchange Online PowerShell
# CRITICAL: Uses mikereports@aftonprop.com shared mailbox, NEVER primary email

param(
    [string]$UserEmail = "mozery@aftonprop.com",
    [string]$SharedMailbox = "mikereports@aftonprop.com"
)

Write-Host "========================================================================"
Write-Host "DOWNLOAD YARDI FROM SHARED MAILBOX (Exchange Online)"
Write-Host "CRITICAL: Using $SharedMailbox via $UserEmail"
Write-Host "========================================================================"
Write-Host ""

# Check if ExchangeOnlineManagement module is installed
Write-Host "[STEP 1] Checking Exchange Online PowerShell module..."
$module = Get-Module -ListAvailable -Name ExchangeOnlineManagement

if (-not $module) {
    Write-Host "ERROR: ExchangeOnlineManagement module not installed"
    Write-Host ""
    Write-Host "Install it with:"
    Write-Host '  Install-Module -Name ExchangeOnlineManagement -Force'
    Write-Host ""
    exit 1
}

Write-Host "[OK] Module found"
Write-Host ""

# Import the module
Write-Host "[STEP 2] Importing Exchange Online module..."
Import-Module ExchangeOnlineManagement -Force
Write-Host "[OK] Module imported"
Write-Host ""

# Connect to Exchange Online
Write-Host "[STEP 3] Connecting to Exchange Online..."
Write-Host "Please sign in with: $UserEmail"
Write-Host "When prompted, complete MFA or enter your password"
Write-Host ""

try {
    Connect-ExchangeOnline -UserPrincipalName $UserEmail -ShowBanner:$false
    Write-Host "[OK] Connected to Exchange Online"
}
catch {
    Write-Host "ERROR: Could not connect to Exchange Online"
    Write-Host $_.Exception.Message
    exit 1
}

Write-Host ""
Write-Host "[STEP 4] Accessing shared mailbox: $SharedMailbox"

# Create output directory
$OutputDir = "C:\Users\MichaelOzery\OneDrive - Afton Properties\Old Dropbox\My PC (DESKTOP-5D77V89)\Mike Ozery\Dashboard Data"
if (-not (Test-Path $OutputDir)) {
    New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null
}

Write-Host "Output directory: $OutputDir"
Write-Host ""

try {
    # Search for emails from YARDI in the shared mailbox
    Write-Host "[STEP 5] Searching for YARDI emails..."

    $SenderEmail = "cdr@yardi.com"

    # Get emails from shared mailbox
    $emails = Search-Mailbox -Identity $SharedMailbox -SearchQuery "from:$SenderEmail" -TargetMailbox $SharedMailbox -TargetFolder "SearchResults" -LogOnly

    Write-Host "Search complete. Getting email details..."
    Write-Host ""

    # Get the messages with attachments
    $downloaded = 0
    $skipped = 0
    $errors = 0

    Write-Host "[STEP 6] Processing emails..."

    # Get all messages from YARDI sender in shared mailbox
    $messages = Get-Mailbox -Identity $SharedMailbox | Get-MailboxFolderStatistics -FolderScope All |
                Where-Object { $_.FolderPath -eq "/Inbox" } |
                Get-Mailbox -Identity $SharedMailbox

    # Alternative: Use Get-InboxRule or search via cmdlets
    # This is a simpler approach - iterate through Inbox items

    $mailbox = Get-Mailbox -Identity $SharedMailbox
    Write-Host "Mailbox: $($mailbox.DisplayName)"

    # Use Search-Mailbox to find messages with attachments from YARDI
    $query = "from:$SenderEmail hasattachments:true"

    $mailItems = Search-Mailbox -Identity $SharedMailbox -SearchQuery $query -IncludeUnsearchableItems $false

    if ($mailItems.Count -eq 0) {
        Write-Host "[WARNING] No emails found from $SenderEmail with attachments"
    }
    else {
        Write-Host "[OK] Found $($mailItems.Count) emails with attachments"

        # Note: Search-Mailbox doesn't directly return attachment data
        # We need to use Get-MailboxFolderPermission and other methods

        Write-Host ""
        Write-Host "[NOTE] Exchange Online PowerShell has limitations for bulk attachment downloads"
        Write-Host "Recommended alternative: Use Outlook desktop client or web interface"
    }

    Write-Host ""
    Write-Host "========================================================================"
    Write-Host "Connection test complete"
    Write-Host "========================================================================"
    Write-Host ""
    Write-Host "STATUS: Connected to $SharedMailbox successfully"
    Write-Host ""
    Write-Host "IMPORTANT:"
    Write-Host "Due to Exchange Online limitations, for bulk downloading attachments,"
    Write-Host "please use one of these methods:"
    Write-Host ""
    Write-Host "Option 1: Outlook Desktop Client"
    Write-Host "  - Add shared mailbox to Outlook"
    Write-Host "  - Select emails from cdr@yardi.com"
    Write-Host "  - Right-click > Save As..."
    Write-Host ""
    Write-Host "Option 2: Outlook Web Access"
    Write-Host "  - Go to https://outlook.office365.com"
    Write-Host "  - Open shared mailbox"
    Write-Host "  - Select emails and download"
    Write-Host ""
    Write-Host "Option 3: Use Python with O365 library"
    Write-Host "  - Run: pip install O365 --break-system-packages"
    Write-Host "  - Use OAuth to connect (one-time setup)"
    Write-Host ""

}
catch {
    Write-Host "ERROR: $($_.Exception.Message)"
    $errors++
}

# Disconnect
Write-Host ""
Write-Host "[CLEANUP] Disconnecting from Exchange Online..."
Disconnect-ExchangeOnline -Confirm:$false
Write-Host "[OK] Disconnected"

Write-Host ""
Write-Host "========================================================================"
Write-Host "Done"
Write-Host "========================================================================"
