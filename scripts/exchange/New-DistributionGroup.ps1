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
    Write-Host "Distribution Group Creation" -ForegroundColor Cyan
    Write-Host "=================================`n" -ForegroundColor Cyan
}

# Display header
Write-Header

# Prompt for primary SMTP address
Write-Host "Enter the Primary SMTP Address (e.g., team-name@optimizely.com):" -ForegroundColor Yellow
$PrimarySmtpAddress = Read-Host

# Validate email format
if ([string]::IsNullOrWhiteSpace($PrimarySmtpAddress)) {
    Write-Warning "[ERROR] Primary SMTP address cannot be empty. Exiting."
    exit 1
}

if (-not ($PrimarySmtpAddress -match '^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$')) {
    Write-Warning "[ERROR] Invalid email format: $PrimarySmtpAddress"
    exit 1
}

# Check if email address already exists
try {
    $existingRecipient = Get-Recipient -Identity $PrimarySmtpAddress -ErrorAction Stop
    Write-Warning "[ERROR] Email address '$PrimarySmtpAddress' is already in use by:"
    Write-Host "  Recipient: $($existingRecipient.DisplayName)"
    Write-Host "  Type: $($existingRecipient.RecipientType)"
    exit 1
} catch {
    # Email doesn't exist, which is what we want
    Write-Host "[SUCCESS] Email address '$PrimarySmtpAddress' is available" -ForegroundColor Green
}

# Prompt for display name
Write-Host "`nEnter the Display Name:" -ForegroundColor Yellow
$DisplayName = Read-Host

# Validate display name
if ([string]::IsNullOrWhiteSpace($DisplayName)) {
    Write-Warning "[ERROR] Display name cannot be empty. Exiting."
    exit 1
}

# Extract alias from email address (part before @)
$Alias = ($PrimarySmtpAddress -split '@')[0]

# Create group name (default to display name)
$GroupName = $DisplayName

# Check if group already exists
try {
    $existingGroup = Get-DistributionGroup -Identity $GroupName -ErrorAction Stop
    Write-Warning "[ERROR] Distribution Group '$GroupName' already exists."
    Write-Host "[INFO] Existing group details:"
    Write-Host "  Display Name: $($existingGroup.DisplayName)"
    Write-Host "  Primary SMTP: $($existingGroup.PrimarySmtpAddress)"
    Write-Host "  Alias: $($existingGroup.Alias)"
    exit 1
} catch {
    # Group doesn't exist, which is what we want
    Write-Host "[SUCCESS] Group name '$GroupName' is available" -ForegroundColor Green
}

# Set group type to Distribution (simplified)
$GroupType = "Distribution"

# Display summary
Write-Host "`n=================================" -ForegroundColor Cyan
Write-Host "Summary:" -ForegroundColor Cyan
Write-Host "Name: $GroupName" -ForegroundColor White
Write-Host "Display Name: $DisplayName" -ForegroundColor White
Write-Host "Alias: $Alias" -ForegroundColor White
Write-Host "Primary SMTP: $PrimarySmtpAddress" -ForegroundColor White
Write-Host "Type: $GroupType" -ForegroundColor White
Write-Host "=================================" -ForegroundColor Cyan

# Confirmation prompt with robust handling
Write-Host ""
do {
    Write-Host "Create this Distribution Group? (Y/N): " -ForegroundColor Yellow -NoNewline
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

# Create the distribution group
Write-Host "`n[INFO] Creating Distribution Group..." -ForegroundColor Cyan

try {
    $newGroup = New-DistributionGroup -Name $GroupName `
        -DisplayName $DisplayName `
        -Alias $Alias `
        -PrimarySmtpAddress $PrimarySmtpAddress `
        -Type $GroupType `
        -ErrorAction Stop
    
    Write-Host "[SUCCESS] Distribution Group created successfully!" -ForegroundColor Green
    
    # Display created group details
    Write-Host "`n=================================" -ForegroundColor Cyan
    Write-Host "Created Group Details:" -ForegroundColor Cyan
    Write-Host "Name: $($newGroup.Name)" -ForegroundColor White
    Write-Host "Display Name: $($newGroup.DisplayName)" -ForegroundColor White
    Write-Host "Alias: $($newGroup.Alias)" -ForegroundColor White
    Write-Host "Primary SMTP: $($newGroup.PrimarySmtpAddress)" -ForegroundColor White
    Write-Host "Type: $($newGroup.GroupType)" -ForegroundColor White
    Write-Host "Identity: $($newGroup.Identity)" -ForegroundColor Gray
    Write-Host "=================================" -ForegroundColor Cyan
    
} catch {
    Write-Warning "[ERROR] Failed to create Distribution Group: $($_.Exception.Message)"
    exit 1
}

# Ask if user wants to add members
Write-Host ""
do {
    Write-Host "Add members to this Distribution Group now? (Y/N): " -ForegroundColor Yellow -NoNewline
    $addMembers = Read-Host
    $addMembers = $addMembers.Trim().ToUpper()
    
    if ($addMembers -eq 'N' -or $addMembers -eq 'NO') {
        Write-Host "[INFO] Distribution Group created without members. You can add members later." -ForegroundColor Yellow
        break
    }
    elseif ($addMembers -eq 'Y' -or $addMembers -eq 'YES') {
        # Prompt for member input
        Write-Host "`nEnter email addresses to add (one per line):" -ForegroundColor Yellow
        Write-Host "Press Enter twice when finished:" -ForegroundColor Gray
        Write-Host "(You can paste multiple lines at once)`n" -ForegroundColor Gray
        
        # Collect multi-line input with robust handling
        $inputLines = @()
        $emptyLineCount = 0
        
        do {
            try {
                $line = Read-Host
                if ($line.Trim() -eq "") {
                    $emptyLineCount++
                } else {
                    $emptyLineCount = 0
                    $inputLines += $line
                }
            }
            catch {
                # Handle any input interruption
                break
            }
        } while ($emptyLineCount -lt 2)
        
        # Process all input lines and split by spaces/newlines
        $allEmails = @()
        foreach ($inputLine in $inputLines) {
            # Split by spaces and filter out empty entries
            $splitEmails = $inputLine -split '\s+' | Where-Object { $_.Trim() -ne "" }
            $allEmails += $splitEmails
        }
        
        # Validate and deduplicate emails
        $emails = @()
        $emailSet = @{}
        
        foreach ($email in $allEmails) {
            $cleanEmail = $email.Trim()
            
            # Skip if empty
            if ($cleanEmail -eq "") { continue }
            
            # Validate email format
            if ($cleanEmail -match '^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$') {
                # Check for duplicates (case insensitive)
                $lowerEmail = $cleanEmail.ToLower()
                if (-not $emailSet.ContainsKey($lowerEmail)) {
                    $emails += $cleanEmail
                    $emailSet[$lowerEmail] = $true
                }
            } else {
                Write-Warning "[WARNING] Invalid email format: $cleanEmail (skipping)"
            }
        }
        
        # Check if any emails were provided
        if ($emails.Count -eq 0) {
            Write-Warning "[WARNING] No valid email addresses provided."
            continue
        }
        
        # Display emails to be added
        Write-Host "`nEmails to be added:" -ForegroundColor Yellow
        foreach ($email in $emails) {
            Write-Host "  - $email" -ForegroundColor Gray
        }
        
        # Confirmation for adding members
        Write-Host ""
        do {
            Write-Host "Add these $($emails.Count) members to the group? (Y/N): " -ForegroundColor Yellow -NoNewline
            $confirmMembers = Read-Host
            $confirmMembers = $confirmMembers.Trim().ToUpper()
            
            if ($confirmMembers -eq 'N' -or $confirmMembers -eq 'NO') {
                Write-Host "[INFO] Members not added." -ForegroundColor Yellow
                break
            }
            elseif ($confirmMembers -eq 'Y' -or $confirmMembers -eq 'YES') {
                # Process member additions
                Write-Host "`n[INFO] Adding members..." -ForegroundColor Cyan
                $successCount = 0
                $failureCount = 0
                
                foreach ($email in $emails) {
                    try {
                        Add-DistributionGroupMember -Identity $GroupName -Member $email -ErrorAction Stop
                        Write-Host "[SUCCESS] Added $email to $($newGroup.DisplayName)" -ForegroundColor Green
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
                
                # Display member addition summary
                Write-Host "`n=================================" -ForegroundColor Cyan
                Write-Host "Member Addition Complete:" -ForegroundColor Cyan
                Write-Host "Successfully added: $successCount" -ForegroundColor Green
                Write-Host "Failed to add: $failureCount" -ForegroundColor $(if ($failureCount -gt 0) { "Red" } else { "Green" })
                Write-Host "Total members: $((Get-DistributionGroupMember -Identity $GroupName).Count)" -ForegroundColor Cyan
                Write-Host "=================================" -ForegroundColor Cyan
                break
            }
            else {
                Write-Host "Please enter Y or N" -ForegroundColor Red
            }
        } while ($true)
        break
    }
    else {
        Write-Host "Please enter Y or N" -ForegroundColor Red
    }
} while ($true)

