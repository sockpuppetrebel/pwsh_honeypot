# Check if the Exchange Online PowerShell module is installed
if (!(Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
    Write-Warning "[WARNING] ExchangeOnlineManagement module is not installed."
    Write-Host "[INFO] Please install it with: Install-Module -Name ExchangeOnlineManagement -Force"
    exit 1
}

# Import the module
Import-Module ExchangeOnlineManagement

# Connect to Exchange Online upfront (before any user prompts)
try {
    # Check if already connected
    Get-Command Get-Mailbox -ErrorAction Stop | Out-Null
    Write-Host "[INFO] Already connected to Exchange Online" -ForegroundColor Green
} catch {
    Write-Host "[INFO] Connecting to Exchange Online..." -ForegroundColor Yellow
    Write-Host "[INFO] A popup window may appear for authentication" -ForegroundColor Yellow
    try {
        Connect-ExchangeOnline -ShowBanner:$false
        Write-Host "[SUCCESS] Connected to Exchange Online" -ForegroundColor Green
    } catch {
        Write-Warning "[ERROR] Failed to connect to Exchange Online: $($_.Exception.Message)"
        exit 1
    }
}

# Function to display header
function Write-Header {
    Write-Host "`n=================================" -ForegroundColor Cyan
    Write-Host "Distribution List Member Addition" -ForegroundColor Cyan
    Write-Host "=================================`n" -ForegroundColor Cyan
}

# Display header
Write-Header

# Prompt for distribution list alias
Write-Host "Enter the Distribution List alias/name:" -ForegroundColor Yellow
$GroupName = Read-Host

# Validate the distribution list exists
try {
    $dlInfo = Get-DistributionGroup -Identity $GroupName -ErrorAction Stop
    Write-Host "[SUCCESS] Found distribution list: $($dlInfo.DisplayName)" -ForegroundColor Green
    Write-Host "[INFO] Current member count: $((Get-DistributionGroupMember -Identity $GroupName).Count)" -ForegroundColor Cyan
} catch {
    Write-Warning "[ERROR] Distribution list '$GroupName' not found: $($_.Exception.Message)"
    exit 1
}

# Prompt for member input
Write-Host "`nEnter email addresses to add (one per line):" -ForegroundColor Yellow
Write-Host "Press Enter twice when finished:" -ForegroundColor Gray
Write-Host "(You can paste multiple lines at once)`n" -ForegroundColor Gray

# Collect multi-line input with robust handling
$emails = @()
$emptyLineCount = 0

do {
    try {
        $line = Read-Host
        if ($line.Trim() -eq "") {
            $emptyLineCount++
        } else {
            $emptyLineCount = 0
            # Validate email format
            if ($line -match '^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$') {
                $emails += $line.Trim()
            } else {
                Write-Warning "[WARNING] Invalid email format: $line (skipping)"
            }
        }
    }
    catch {
        # Handle any input interruption
        break
    }
} while ($emptyLineCount -lt 2)

# Check if any emails were provided
if ($emails.Count -eq 0) {
    Write-Warning "[WARNING] No valid email addresses provided. Exiting."
    exit 0
}

# Display summary
Write-Host "`n=================================" -ForegroundColor Cyan
Write-Host "Summary:" -ForegroundColor Cyan
Write-Host "Distribution List: $($dlInfo.DisplayName)" -ForegroundColor White
Write-Host "Emails to add: $($emails.Count)" -ForegroundColor White
Write-Host "=================================" -ForegroundColor Cyan

# Display emails to be added
Write-Host "`nEmails to be added:" -ForegroundColor Yellow
foreach ($email in $emails) {
    Write-Host "  - $email" -ForegroundColor Gray
}

# Confirmation prompt with robust handling
Write-Host ""
do {
    Write-Host "Proceed with adding these members? (Y/N): " -ForegroundColor Yellow -NoNewline
    $confirm = Read-Host
    $confirm = $confirm.Trim().ToUpper()
    
    if ($confirm -eq 'N' -or $confirm -eq 'NO') {
        Write-Host "[INFO] Operation cancelled by user." -ForegroundColor Yellow
        exit 0
    }
    elseif ($confirm -eq 'Y' -or $confirm -eq 'YES') {
        break
    }
    else {
        Write-Host "Please enter Y or N" -ForegroundColor Red
    }
} while ($true)

# Process additions
Write-Host "`n[INFO] Processing additions..." -ForegroundColor Cyan
$successCount = 0
$failureCount = 0

foreach ($email in $emails) {
    try {
        Add-DistributionGroupMember -Identity $GroupName -Member $email -ErrorAction Stop
        Write-Host "[SUCCESS] Added $email to $($dlInfo.DisplayName)" -ForegroundColor Green
        $successCount++
    } catch {
        if ($_.Exception.Message -match "already a member") {
            Write-Host "[INFO] $email is already a member" -ForegroundColor Yellow
        } else {
            Write-Warning "[ERROR] Failed to add $email - $($_.Exception.Message)"
            $failureCount++
        }
    }
}

# Display final summary
Write-Host "`n=================================" -ForegroundColor Cyan
Write-Host "Operation Complete:" -ForegroundColor Cyan
Write-Host "Successfully added: $successCount" -ForegroundColor Green
Write-Host "Failed to add: $failureCount" -ForegroundColor $(if ($failureCount -gt 0) { "Red" } else { "Green" })
Write-Host "New member count: $((Get-DistributionGroupMember -Identity $GroupName).Count)" -ForegroundColor Cyan
Write-Host "=================================" -ForegroundColor Cyan

