#Requires -Modules ExchangeOnlineManagement

<#
.SYNOPSIS
    Adds a user to multiple distribution groups and shared mailboxes
    
.DESCRIPTION
    This script adds a specified user to multiple recipients including:
    - Distribution groups (as member)
    - Mail-enabled security groups (as member)
    - Shared mailboxes (as delegate with permissions)
    
.PARAMETER UserToAdd
    The email address of the user to add
    
.PARAMETER Recipients
    Array of recipient email addresses (groups and shared mailboxes)
    
.PARAMETER MailboxPermission
    Permission level for shared mailboxes (FullAccess or ReadPermission)
    
.PARAMETER AutoMapping
    Whether to auto-map shared mailboxes in Outlook
    
.EXAMPLE
    .\Add-UserToMultipleGroups.ps1 -UserToAdd "first.last@optimizely.com"
    
.EXAMPLE
    .\Add-UserToMultipleGroups.ps1 -UserToAdd "user@domain.com" -MailboxPermission "ReadPermission"
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)]
    [string]$UserToAdd,
    
    [Parameter(Mandatory=$false)]
    [string[]]$Recipients = @(
        "cit-aws@ektron.com",
        "ektronmanagedservices-aws@ektron.com",
        "ektaz1@ektron.com",
        "EktronSupport-AWS@ektron.com",
        "aws.idiotest@episerver.net",
        "managedservices_rnd@ektron.com",
        "aws.welcome-gehc@episerver.net",
        "aws.welcome-trayio@episerver.net",
        "aws.welcome-trayio-hubspot@episerver.net",
        "aws.welcome-webproofing2@episerver.net"
    ),
    
    [Parameter(Mandatory=$false)]
    [ValidateSet("FullAccess", "ReadPermission")]
    [string]$MailboxPermission = "FullAccess",
    
    [Parameter(Mandatory=$false)]
    [bool]$AutoMapping = $true
)

# Connect to Exchange Online if not already connected
$connectionStatus = Get-PSSession | Where-Object {$_.ConfigurationName -eq "Microsoft.Exchange" -and $_.State -eq "Opened"}
if (-not $connectionStatus) {
    Write-Host "Connecting to Exchange Online..." -ForegroundColor Yellow
    try {
        Connect-ExchangeOnline -ShowBanner:$false
        Write-Host "Connected to Exchange Online" -ForegroundColor Green
    }
    catch {
        Write-Error "Failed to connect to Exchange Online: $_"
        exit 1
    }
}

# Verify the user exists
Write-Host "`nVerifying user: $UserToAdd" -ForegroundColor Yellow
$userMailbox = Get-Mailbox -Identity $UserToAdd -ErrorAction SilentlyContinue
if (-not $userMailbox) {
    Write-Error "User $UserToAdd not found or doesn't have a mailbox"
    exit 1
}
Write-Host "User verified: $($userMailbox.DisplayName)" -ForegroundColor Green

# Process results tracking
$results = @{
    GroupsAdded = @()
    GroupsFailed = @()
    MailboxesAdded = @()
    MailboxesFailed = @()
    NotFound = @()
    Skipped = @()
}

Write-Host "`nProcessing recipients..." -ForegroundColor Yellow

foreach ($recipient in $Recipients) {
    try {
        # Get recipient details
        $recipientObject = Get-Recipient -Identity $recipient -ErrorAction SilentlyContinue
        
        if (-not $recipientObject) {
            Write-Host "  ✗ Not found: $recipient" -ForegroundColor Red
            $results.NotFound += $recipient
            continue
        }
        
        Write-Host "`n  Processing: $recipient" -ForegroundColor Cyan
        Write-Host "  Type: $($recipientObject.RecipientType)" -ForegroundColor Gray
        
        switch ($recipientObject.RecipientType) {
            { $_ -in "MailUniversalDistributionGroup", "MailUniversalSecurityGroup", "MailNonUniversalGroup", "DynamicDistributionGroup" } {
                # Check if already a member
                $existingMembers = Get-DistributionGroupMember -Identity $recipient -ErrorAction SilentlyContinue | 
                    Where-Object { $_.PrimarySmtpAddress -eq $UserToAdd }
                
                if ($existingMembers) {
                    Write-Host "  → ALREADY MEMBER: User is already a member of this distribution group" -ForegroundColor Yellow
                    $results.Skipped += "$recipient (already member)"
                }
                else {
                    # Add to distribution group
                    Add-DistributionGroupMember -Identity $recipient -Member $UserToAdd -ErrorAction Stop
                    Write-Host "  ✓ ADDED AS MEMBER: User can now receive emails sent to this group" -ForegroundColor Green
                    $results.GroupsAdded += $recipient
                }
            }
            
            "SharedMailbox" {
                # Check existing permissions
                $existingPermissions = Get-MailboxPermission -Identity $recipient -User $UserToAdd -ErrorAction SilentlyContinue
                $existingSendAs = Get-RecipientPermission -Identity $recipient -Trustee $UserToAdd -ErrorAction SilentlyContinue
                
                $hasFullAccess = $existingPermissions -and ($existingPermissions.AccessRights -contains "FullAccess")
                $hasSendAs = $existingSendAs -and ($existingSendAs.AccessRights -contains "SendAs")
                
                if ($hasFullAccess -and $hasSendAs) {
                    Write-Host "  → ALREADY HAS ACCESS: User already has full access and send-as permissions" -ForegroundColor Yellow
                    $results.Skipped += "$recipient (already has full access)"
                }
                else {
                    $addedPermissions = @()
                    
                    # Add mailbox permission if not exists
                    if (-not $hasFullAccess) {
                        Add-MailboxPermission -Identity $recipient -User $UserToAdd -AccessRights $MailboxPermission -AutoMapping:$AutoMapping -ErrorAction Stop | Out-Null
                        $addedPermissions += $MailboxPermission
                    }
                    
                    # Add Send As permission if FullAccess and not exists
                    if ($MailboxPermission -eq "FullAccess" -and -not $hasSendAs) {
                        Add-RecipientPermission -Identity $recipient -Trustee $UserToAdd -AccessRights SendAs -Confirm:$false -ErrorAction Stop | Out-Null
                        $addedPermissions += "SendAs"
                    }
                    
                    if ($addedPermissions.Count -gt 0) {
                        Write-Host "  ✓ ADDED SHARED MAILBOX ACCESS: $($addedPermissions -join ', ') permissions granted" -ForegroundColor Green
                        if ($AutoMapping) {
                            Write-Host "    → Mailbox will auto-appear in Outlook (15-60 min delay)" -ForegroundColor Cyan
                        }
                        $results.MailboxesAdded += $recipient
                    }
                    else {
                        Write-Host "  → ALREADY HAS PARTIAL ACCESS: Some permissions already exist" -ForegroundColor Yellow
                        $results.Skipped += "$recipient (partial access exists)"
                    }
                }
            }
            
            "UserMailbox" {
                # Check existing SendOnBehalf permissions
                $mailboxInfo = Get-Mailbox -Identity $recipient -ErrorAction SilentlyContinue
                $existingDelegates = $mailboxInfo.GrantSendOnBehalfTo | Where-Object { $_ -like "*$UserToAdd*" }
                
                if ($existingDelegates) {
                    Write-Host "  → ALREADY DELEGATE: User already has SendOnBehalf permission" -ForegroundColor Yellow
                    $results.Skipped += "$recipient (already delegate)"
                }
                else {
                    # Add as delegate with SendOnBehalf permission
                    Set-Mailbox -Identity $recipient -GrantSendOnBehalfTo @{Add=$UserToAdd} -ErrorAction Stop
                    Write-Host "  ✓ ADDED AS DELEGATE: User can send emails on behalf of this mailbox" -ForegroundColor Green
                    $results.MailboxesAdded += "$recipient (delegate)"
                }
            }
            
            default {
                Write-Host "  → Unsupported recipient type: $($recipientObject.RecipientType)" -ForegroundColor Yellow
                $results.Skipped += "$recipient ($($recipientObject.RecipientType))"
            }
        }
    }
    catch {
        Write-Host "  ✗ Error: $($_.Exception.Message)" -ForegroundColor Red
        
        # Categorize the error
        if ($recipientObject.RecipientType -like "*Group*") {
            $results.GroupsFailed += "$recipient - $($_.Exception.Message)"
        }
        else {
            $results.MailboxesFailed += "$recipient - $($_.Exception.Message)"
        }
    }
}

# Display summary
Write-Host "`n=== SUMMARY ===" -ForegroundColor Cyan
Write-Host "User: $UserToAdd" -ForegroundColor White
Write-Host "`nGroups:" -ForegroundColor Yellow
Write-Host "  Added: $($results.GroupsAdded.Count)" -ForegroundColor Green
Write-Host "  Failed: $($results.GroupsFailed.Count)" -ForegroundColor Red

Write-Host "`nMailboxes:" -ForegroundColor Yellow
Write-Host "  Added: $($results.MailboxesAdded.Count)" -ForegroundColor Green
Write-Host "  Failed: $($results.MailboxesFailed.Count)" -ForegroundColor Red

Write-Host "`nOther:" -ForegroundColor Yellow
Write-Host "  Not Found: $($results.NotFound.Count)" -ForegroundColor Red
Write-Host "  Skipped: $($results.Skipped.Count)" -ForegroundColor Yellow

# Show details if any failures
if ($results.GroupsFailed.Count -gt 0) {
    Write-Host "`nFailed Groups:" -ForegroundColor Red
    $results.GroupsFailed | ForEach-Object { Write-Host "  - $_" -ForegroundColor Red }
}

if ($results.MailboxesFailed.Count -gt 0) {
    Write-Host "`nFailed Mailboxes:" -ForegroundColor Red
    $results.MailboxesFailed | ForEach-Object { Write-Host "  - $_" -ForegroundColor Red }
}

if ($results.NotFound.Count -gt 0) {
    Write-Host "`nNot Found:" -ForegroundColor Red
    $results.NotFound | ForEach-Object { Write-Host "  - $_" -ForegroundColor Red }
}

Write-Host "`n=== NEXT STEPS ===" -ForegroundColor Cyan
Write-Host "1. Verify group memberships in Exchange Admin Center" -ForegroundColor White
Write-Host "2. Have user check Outlook for shared mailbox access" -ForegroundColor White
Write-Host "3. Test sending emails to distribution groups" -ForegroundColor White
Write-Host "4. For shared mailboxes, allow 15-60 minutes for auto-mapping" -ForegroundColor White