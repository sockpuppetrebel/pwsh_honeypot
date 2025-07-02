<#
.SYNOPSIS
    Complete user offboarding checklist with automated actions.

.DESCRIPTION
    This script performs a comprehensive offboarding process including:
    - Disabling the user account
    - Resetting password
    - Removing from all groups
    - Converting mailbox to shared
    - Removing licenses
    - Generating offboarding report
    - Setting out-of-office message
    - Forwarding email to manager

.PARAMETER UserPrincipalName
    The UPN of the user to offboard

.PARAMETER ManagerEmail
    Email address to forward mail to (optional)

.PARAMETER PreserveMailbox
    Convert mailbox to shared instead of removing

.PARAMETER TicketNumber
    HR ticket number for tracking

.EXAMPLE
    .\Complete-UserOffboarding.ps1 -UserPrincipalName "john.doe@company.com" -ManagerEmail "manager@company.com" -TicketNumber "HR-2024-001"

.NOTES
    Author: HR Automation Script
    Date: 2025-07-02
    Version: 1.0
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory=$true)]
    [string]$UserPrincipalName,
    
    [Parameter(Mandatory=$false)]
    [string]$ManagerEmail,
    
    [Parameter(Mandatory=$false)]
    [switch]$PreserveMailbox = $true,
    
    [Parameter(Mandatory=$false)]
    [string]$TicketNumber = "N/A",
    
    [Parameter(Mandatory=$false)]
    [switch]$WhatIf
)

# Initialize logging
$LogPath = "$PSScriptRoot\Logs\Offboarding"
if (-not (Test-Path $LogPath)) {
    New-Item -ItemType Directory -Path $LogPath -Force | Out-Null
}
$SessionId = Get-Date -Format 'yyyyMMdd-HHmmss'
$LogFile = "$LogPath\Offboarding-$($UserPrincipalName.Replace('@','_'))-$SessionId.log"
Start-Transcript -Path $LogFile

# Initialize results tracking
$results = @{
    UserPrincipalName = $UserPrincipalName
    TicketNumber = $TicketNumber
    StartTime = Get-Date
    Steps = @()
}

function Add-OffboardingStep {
    param(
        [string]$Step,
        [string]$Status,
        [string]$Details = ""
    )
    
    $results.Steps += @{
        Step = $Step
        Status = $Status
        Details = $Details
        Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    }
    
    $color = switch ($Status) {
        "Success" { "Green" }
        "Failed" { "Red" }
        "Skipped" { "Yellow" }
        "InProgress" { "Cyan" }
        default { "White" }
    }
    
    Write-Host "[$Status] $Step $(if ($Details) {": $Details"})" -ForegroundColor $color
}

Write-Host "`n=== USER OFFBOARDING PROCESS ===" -ForegroundColor Green
Write-Host "User: $UserPrincipalName" -ForegroundColor Yellow
Write-Host "Ticket: $TicketNumber" -ForegroundColor Yellow
Write-Host "Started: $(Get-Date)" -ForegroundColor Yellow
if ($WhatIf) {
    Write-Host "MODE: What-If (No changes will be made)" -ForegroundColor Cyan
}
Write-Host "="*50 -ForegroundColor Green

try {
    # Connect to required services
    Write-Host "`nConnecting to Microsoft services..." -ForegroundColor Cyan
    
    # Connect to Microsoft Graph
    $mgContext = Get-MgContext -ErrorAction SilentlyContinue
    if (-not $mgContext) {
        Connect-MgGraph -Scopes "User.ReadWrite.All", "Group.ReadWrite.All", "Directory.ReadWrite.All", "Mail.Send" -NoWelcome
    }
    
    # Connect to Exchange Online
    $exoSession = Get-PSSession | Where-Object {$_.ConfigurationName -eq "Microsoft.Exchange" -and $_.State -eq "Opened"}
    if (-not $exoSession) {
        Connect-ExchangeOnline -ShowBanner:$false
    }
    
    # Step 1: Get User Information
    Write-Host "`n[Step 1] Retrieving user information..." -ForegroundColor Cyan
    $user = Get-MgUser -UserId $UserPrincipalName -Property * -ErrorAction Stop
    $mailbox = Get-Mailbox -Identity $UserPrincipalName -ErrorAction SilentlyContinue
    
    Add-OffboardingStep -Step "Retrieve User Info" -Status "Success" -Details $user.DisplayName
    
    # Step 2: Disable User Account
    Write-Host "`n[Step 2] Disabling user account..." -ForegroundColor Cyan
    if ($PSCmdlet.ShouldProcess($UserPrincipalName, "Disable account")) {
        Update-MgUser -UserId $user.Id -AccountEnabled:$false
        Add-OffboardingStep -Step "Disable Account" -Status "Success"
    } else {
        Add-OffboardingStep -Step "Disable Account" -Status "Skipped" -Details "WhatIf mode"
    }
    
    # Step 3: Reset Password
    Write-Host "`n[Step 3] Resetting password..." -ForegroundColor Cyan
    $newPassword = [System.Web.Security.Membership]::GeneratePassword(16, 4)
    if ($PSCmdlet.ShouldProcess($UserPrincipalName, "Reset password")) {
        $passwordProfile = @{
            Password = $newPassword
            ForceChangePasswordNextSignIn = $true
        }
        Update-MgUser -UserId $user.Id -PasswordProfile $passwordProfile
        Add-OffboardingStep -Step "Reset Password" -Status "Success" -Details "Password reset to random value"
    } else {
        Add-OffboardingStep -Step "Reset Password" -Status "Skipped" -Details "WhatIf mode"
    }
    
    # Step 4: Block Sign-in
    Write-Host "`n[Step 4] Blocking sign-in..." -ForegroundColor Cyan
    if ($PSCmdlet.ShouldProcess($UserPrincipalName, "Block sign-in")) {
        Revoke-MgUserSignInSession -UserId $user.Id
        Add-OffboardingStep -Step "Block Sign-in" -Status "Success"
    } else {
        Add-OffboardingStep -Step "Block Sign-in" -Status "Skipped" -Details "WhatIf mode"
    }
    
    # Step 5: Set Out-of-Office Message
    Write-Host "`n[Step 5] Setting out-of-office message..." -ForegroundColor Cyan
    if ($mailbox -and $PSCmdlet.ShouldProcess($UserPrincipalName, "Set OOF message")) {
        $oofMessage = @"
Thank you for your email. I am no longer with the company.

For immediate assistance, please contact:
$(if ($ManagerEmail) { "- My former manager: $ManagerEmail" } else { "- The main office" })

Thank you,
$($user.DisplayName)
"@
        Set-MailboxAutoReplyConfiguration -Identity $UserPrincipalName `
            -AutoReplyState Enabled `
            -InternalMessage $oofMessage `
            -ExternalMessage $oofMessage
        Add-OffboardingStep -Step "Set Out-of-Office" -Status "Success"
    } else {
        Add-OffboardingStep -Step "Set Out-of-Office" -Status "Skipped" -Details $(if (-not $mailbox) {"No mailbox"} else {"WhatIf mode"})
    }
    
    # Step 6: Forward Email to Manager
    Write-Host "`n[Step 6] Setting email forwarding..." -ForegroundColor Cyan
    if ($mailbox -and $ManagerEmail -and $PSCmdlet.ShouldProcess($UserPrincipalName, "Forward email to $ManagerEmail")) {
        Set-Mailbox -Identity $UserPrincipalName -ForwardingAddress $ManagerEmail -DeliverToMailboxAndForward $true
        Add-OffboardingStep -Step "Email Forwarding" -Status "Success" -Details "Forwarding to $ManagerEmail"
    } else {
        Add-OffboardingStep -Step "Email Forwarding" -Status "Skipped" -Details $(if (-not $ManagerEmail) {"No manager specified"} else {"WhatIf mode"})
    }
    
    # Step 7: Remove from All Groups
    Write-Host "`n[Step 7] Removing from groups..." -ForegroundColor Cyan
    $groups = Get-MgUserMemberOf -UserId $user.Id -All
    $groupCount = 0
    
    foreach ($group in $groups | Where-Object { $_.'@odata.type' -eq '#microsoft.graph.group' }) {
        if ($PSCmdlet.ShouldProcess($group.AdditionalProperties.displayName, "Remove user from group")) {
            try {
                Remove-MgGroupMemberByRef -GroupId $group.Id -DirectoryObjectId $user.Id
                $groupCount++
            } catch {
                # Skip errors for dynamic groups
            }
        }
    }
    Add-OffboardingStep -Step "Remove from Groups" -Status "Success" -Details "Removed from $groupCount groups"
    
    # Step 8: Remove Licenses
    Write-Host "`n[Step 8] Removing licenses..." -ForegroundColor Cyan
    $licenses = Get-MgUserLicenseDetail -UserId $user.Id
    if ($licenses -and $PSCmdlet.ShouldProcess($UserPrincipalName, "Remove all licenses")) {
        $licensesToRemove = $licenses | ForEach-Object { $_.SkuId }
        Set-MgUserLicense -UserId $user.Id -AddLicenses @() -RemoveLicenses $licensesToRemove
        Add-OffboardingStep -Step "Remove Licenses" -Status "Success" -Details "$($licenses.Count) licenses removed"
    } else {
        Add-OffboardingStep -Step "Remove Licenses" -Status "Skipped" -Details $(if (-not $licenses) {"No licenses"} else {"WhatIf mode"})
    }
    
    # Step 9: Convert Mailbox to Shared (if requested)
    Write-Host "`n[Step 9] Converting mailbox..." -ForegroundColor Cyan
    if ($mailbox -and $PreserveMailbox -and $PSCmdlet.ShouldProcess($UserPrincipalName, "Convert to shared mailbox")) {
        Set-Mailbox -Identity $UserPrincipalName -Type Shared
        Add-OffboardingStep -Step "Convert Mailbox" -Status "Success" -Details "Converted to shared mailbox"
    } else {
        Add-OffboardingStep -Step "Convert Mailbox" -Status "Skipped" -Details $(if (-not $PreserveMailbox) {"Not requested"} else {"WhatIf mode"})
    }
    
    # Step 10: Hide from GAL
    Write-Host "`n[Step 10] Hiding from address lists..." -ForegroundColor Cyan
    if ($mailbox -and $PSCmdlet.ShouldProcess($UserPrincipalName, "Hide from GAL")) {
        Set-Mailbox -Identity $UserPrincipalName -HiddenFromAddressListsEnabled $true
        Add-OffboardingStep -Step "Hide from GAL" -Status "Success"
    } else {
        Add-OffboardingStep -Step "Hide from GAL" -Status "Skipped" -Details "WhatIf mode"
    }
    
    # Generate final report
    $results.EndTime = Get-Date
    $results.Duration = $results.EndTime - $results.StartTime
    
    $reportPath = "$LogPath\OffboardingReport-$($UserPrincipalName.Replace('@','_'))-$SessionId.json"
    $results | ConvertTo-Json -Depth 10 | Out-File $reportPath
    
    # Display summary
    Write-Host "`n$('='*50)" -ForegroundColor Green
    Write-Host "OFFBOARDING SUMMARY" -ForegroundColor Green
    Write-Host "$('='*50)" -ForegroundColor Green
    Write-Host "User: $($user.DisplayName) ($UserPrincipalName)"
    Write-Host "Ticket: $TicketNumber"
    Write-Host "Duration: $($results.Duration.TotalMinutes.ToString('0.0')) minutes"
    Write-Host "`nSteps Completed:"
    
    $results.Steps | ForEach-Object {
        $color = switch ($_.Status) {
            "Success" { "Green" }
            "Failed" { "Red" }
            "Skipped" { "Yellow" }
            default { "White" }
        }
        Write-Host "  [$($_.Status)] $($_.Step)" -ForegroundColor $color
    }
    
    Write-Host "`nReports saved to:"
    Write-Host "  - Log: $LogFile" -ForegroundColor Yellow
    Write-Host "  - Report: $reportPath" -ForegroundColor Yellow
    
    if ($WhatIf) {
        Write-Host "`nNOTE: This was a What-If run. No actual changes were made." -ForegroundColor Cyan
    }
    
} catch {
    Write-Host "`nERROR: $($_.Exception.Message)" -ForegroundColor Red
    Add-OffboardingStep -Step "Process Error" -Status "Failed" -Details $_.Exception.Message
} finally {
    Stop-Transcript
}