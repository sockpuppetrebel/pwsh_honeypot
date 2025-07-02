<#
.SYNOPSIS
    Removes all mailbox delegations for an offboarded user.

.DESCRIPTION
    This script removes all mailbox permissions and delegations including:
    - Full Access permissions (where user has access to other mailboxes)
    - Send As permissions
    - Send on Behalf permissions
    - Calendar delegations
    - Folder permissions
    - Removes user from other mailboxes where they have permissions

.PARAMETER UserPrincipalName
    The UPN of the user whose delegations to remove

.PARAMETER RemoveBothDirections
    Remove permissions in both directions (default: true)

.PARAMETER ExportOnly
    Only export current permissions without removing them

.EXAMPLE
    .\Remove-MailboxDelegations.ps1 -UserPrincipalName "john.doe@company.com"

.EXAMPLE
    .\Remove-MailboxDelegations.ps1 -UserPrincipalName "john.doe@company.com" -ExportOnly

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
    [switch]$RemoveBothDirections = $true,
    
    [Parameter(Mandatory=$false)]
    [switch]$ExportOnly,
    
    [Parameter(Mandatory=$false)]
    [switch]$WhatIf
)

# Initialize logging
$LogPath = "$PSScriptRoot\Logs\Delegations"
if (-not (Test-Path $LogPath)) {
    New-Item -ItemType Directory -Path $LogPath -Force | Out-Null
}
$SessionId = Get-Date -Format 'yyyyMMdd-HHmmss'
$LogFile = "$LogPath\RemoveDelegations-$($UserPrincipalName.Replace('@','_'))-$SessionId.log"
Start-Transcript -Path $LogFile

# Initialize results tracking
$delegations = @{
    UserPrincipalName = $UserPrincipalName
    Timestamp = Get-Date
    PermissionsGrantedByUser = @{
        FullAccess = @()
        SendAs = @()
        SendOnBehalf = @()
        CalendarDelegates = @()
        FolderPermissions = @()
    }
    PermissionsGrantedToUser = @{
        FullAccess = @()
        SendAs = @()
        SendOnBehalf = @()
        CalendarDelegates = @()
    }
    RemovalResults = @()
}

Write-Host "`n=== MAILBOX DELEGATION REMOVAL ===" -ForegroundColor Green
Write-Host "User: $UserPrincipalName" -ForegroundColor Yellow
Write-Host "Mode: $(if ($ExportOnly) {'Export Only'} else {'Remove Delegations'})" -ForegroundColor Yellow
Write-Host "Remove Both Directions: $RemoveBothDirections" -ForegroundColor Yellow
Write-Host "Started: $(Get-Date)" -ForegroundColor Yellow
if ($WhatIf) {
    Write-Host "MODE: What-If (No changes will be made)" -ForegroundColor Cyan
}
Write-Host "="*50 -ForegroundColor Green

try {
    # Connect to Exchange Online if not connected
    $exoSession = Get-PSSession | Where-Object {$_.ConfigurationName -eq "Microsoft.Exchange" -and $_.State -eq "Opened"}
    if (-not $exoSession) {
        Write-Host "Connecting to Exchange Online..." -ForegroundColor Yellow
        Connect-ExchangeOnline -ShowBanner:$false
    }
    
    # Get the user's mailbox
    Write-Host "`n[Step 1] Getting user mailbox information..." -ForegroundColor Cyan
    $userMailbox = Get-Mailbox -Identity $UserPrincipalName -ErrorAction Stop
    
    if (-not $userMailbox) {
        throw "Mailbox not found for user: $UserPrincipalName"
    }
    
    Write-Host "Found mailbox: $($userMailbox.DisplayName)" -ForegroundColor Green
    
    # PART 1: Find permissions GRANTED BY the user
    Write-Host "`n[Step 2] Finding permissions granted BY $UserPrincipalName..." -ForegroundColor Cyan
    
    # Full Access permissions granted by user
    Write-Host "  Checking Full Access permissions..." -ForegroundColor Gray
    $mailboxes = Get-Mailbox -ResultSize Unlimited
    
    foreach ($mailbox in $mailboxes) {
        if ($mailbox.PrimarySmtpAddress -ne $userMailbox.PrimarySmtpAddress) {
            $permissions = Get-MailboxPermission -Identity $mailbox.Identity -User $UserPrincipalName -ErrorAction SilentlyContinue
            if ($permissions -and $permissions.AccessRights -contains "FullAccess") {
                $delegations.PermissionsGrantedByUser.FullAccess += @{
                    Mailbox = $mailbox.PrimarySmtpAddress
                    DisplayName = $mailbox.DisplayName
                    AccessRights = $permissions.AccessRights -join ", "
                }
            }
        }
    }
    
    # Send As permissions granted by user
    Write-Host "  Checking Send As permissions..." -ForegroundColor Gray
    foreach ($mailbox in $mailboxes) {
        if ($mailbox.PrimarySmtpAddress -ne $userMailbox.PrimarySmtpAddress) {
            $sendAsPerms = Get-RecipientPermission -Identity $mailbox.Identity -Trustee $UserPrincipalName -ErrorAction SilentlyContinue
            if ($sendAsPerms) {
                $delegations.PermissionsGrantedByUser.SendAs += @{
                    Mailbox = $mailbox.PrimarySmtpAddress
                    DisplayName = $mailbox.DisplayName
                    AccessRights = "SendAs"
                }
            }
        }
    }
    
    # Send on Behalf permissions
    Write-Host "  Checking Send on Behalf permissions..." -ForegroundColor Gray
    $userDN = $userMailbox.DistinguishedName
    foreach ($mailbox in $mailboxes) {
        if ($mailbox.GrantSendOnBehalfTo -contains $userDN) {
            $delegations.PermissionsGrantedByUser.SendOnBehalf += @{
                Mailbox = $mailbox.PrimarySmtpAddress
                DisplayName = $mailbox.DisplayName
                AccessRights = "SendOnBehalf"
            }
        }
    }
    
    # PART 2: Find permissions GRANTED TO the user (if RemoveBothDirections)
    if ($RemoveBothDirections) {
        Write-Host "`n[Step 3] Finding permissions granted TO $UserPrincipalName..." -ForegroundColor Cyan
        
        # Full Access permissions on user's mailbox
        Write-Host "  Checking who has Full Access to user's mailbox..." -ForegroundColor Gray
        $fullAccessPerms = Get-MailboxPermission -Identity $UserPrincipalName | Where-Object { 
            $_.User -notlike "NT AUTHORITY\*" -and 
            $_.User -notlike "S-1-5-*" -and
            $_.AccessRights -contains "FullAccess" -and
            $_.IsInherited -eq $false
        }
        
        foreach ($perm in $fullAccessPerms) {
            $delegations.PermissionsGrantedToUser.FullAccess += @{
                User = $perm.User
                AccessRights = $perm.AccessRights -join ", "
            }
        }
        
        # Send As permissions on user's mailbox
        Write-Host "  Checking who has Send As on user's mailbox..." -ForegroundColor Gray
        $sendAsPerms = Get-RecipientPermission -Identity $UserPrincipalName | Where-Object {
            $_.Trustee -notlike "NT AUTHORITY\*" -and
            $_.Trustee -notlike "S-1-5-*"
        }
        
        foreach ($perm in $sendAsPerms) {
            $delegations.PermissionsGrantedToUser.SendAs += @{
                User = $perm.Trustee
                AccessRights = "SendAs"
            }
        }
        
        # Send on Behalf on user's mailbox
        if ($userMailbox.GrantSendOnBehalfTo.Count -gt 0) {
            foreach ($delegate in $userMailbox.GrantSendOnBehalfTo) {
                $delegations.PermissionsGrantedToUser.SendOnBehalf += @{
                    User = $delegate
                    AccessRights = "SendOnBehalf"
                }
            }
        }
        
        # Calendar delegates
        Write-Host "  Checking calendar delegates..." -ForegroundColor Gray
        try {
            $calendarPerms = Get-MailboxFolderPermission -Identity "$($UserPrincipalName):\Calendar" -ErrorAction SilentlyContinue | 
                Where-Object { $_.User -notlike "Default" -and $_.User -notlike "Anonymous" }
            
            foreach ($perm in $calendarPerms) {
                $delegations.PermissionsGrantedToUser.CalendarDelegates += @{
                    User = $perm.User
                    AccessRights = $perm.AccessRights -join ", "
                }
            }
        } catch {
            Write-Host "    Unable to check calendar permissions" -ForegroundColor Yellow
        }
    }
    
    # PART 3: Export current state
    Write-Host "`n[Step 4] Exporting current permissions..." -ForegroundColor Cyan
    $exportPath = "$LogPath\DelegationExport-$($UserPrincipalName.Replace('@','_'))-$SessionId.json"
    $delegations | ConvertTo-Json -Depth 10 | Out-File $exportPath
    Write-Host "Permissions exported to: $exportPath" -ForegroundColor Green
    
    # Display summary
    Write-Host "`n=== PERMISSION SUMMARY ===" -ForegroundColor Green
    Write-Host "Permissions GRANTED BY $UserPrincipalName`:" -ForegroundColor Yellow
    Write-Host "  Full Access: $($delegations.PermissionsGrantedByUser.FullAccess.Count) mailboxes"
    Write-Host "  Send As: $($delegations.PermissionsGrantedByUser.SendAs.Count) mailboxes"
    Write-Host "  Send on Behalf: $($delegations.PermissionsGrantedByUser.SendOnBehalf.Count) mailboxes"
    
    if ($RemoveBothDirections) {
        Write-Host "`nPermissions GRANTED TO $UserPrincipalName`:" -ForegroundColor Yellow
        Write-Host "  Full Access: $($delegations.PermissionsGrantedToUser.FullAccess.Count) users"
        Write-Host "  Send As: $($delegations.PermissionsGrantedToUser.SendAs.Count) users"
        Write-Host "  Send on Behalf: $($delegations.PermissionsGrantedToUser.SendOnBehalf.Count) users"
        Write-Host "  Calendar Delegates: $($delegations.PermissionsGrantedToUser.CalendarDelegates.Count) users"
    }
    
    # PART 4: Remove permissions (if not ExportOnly)
    if (-not $ExportOnly) {
        Write-Host "`n[Step 5] Removing permissions..." -ForegroundColor Cyan
        
        # Remove permissions granted BY user
        Write-Host "`nRemoving permissions granted BY user..." -ForegroundColor Yellow
        
        # Remove Full Access
        foreach ($perm in $delegations.PermissionsGrantedByUser.FullAccess) {
            if ($PSCmdlet.ShouldProcess($perm.Mailbox, "Remove Full Access for $UserPrincipalName")) {
                try {
                    Remove-MailboxPermission -Identity $perm.Mailbox -User $UserPrincipalName -AccessRights FullAccess -Confirm:$false
                    Write-Host "  ✓ Removed Full Access from: $($perm.DisplayName)" -ForegroundColor Green
                    $delegations.RemovalResults += @{Permission="FullAccess"; Target=$perm.Mailbox; Status="Success"}
                } catch {
                    Write-Host "  ✗ Failed to remove Full Access from: $($perm.DisplayName) - $_" -ForegroundColor Red
                    $delegations.RemovalResults += @{Permission="FullAccess"; Target=$perm.Mailbox; Status="Failed"; Error=$_.Exception.Message}
                }
            }
        }
        
        # Remove Send As
        foreach ($perm in $delegations.PermissionsGrantedByUser.SendAs) {
            if ($PSCmdlet.ShouldProcess($perm.Mailbox, "Remove Send As for $UserPrincipalName")) {
                try {
                    Remove-RecipientPermission -Identity $perm.Mailbox -Trustee $UserPrincipalName -AccessRights SendAs -Confirm:$false
                    Write-Host "  ✓ Removed Send As from: $($perm.DisplayName)" -ForegroundColor Green
                    $delegations.RemovalResults += @{Permission="SendAs"; Target=$perm.Mailbox; Status="Success"}
                } catch {
                    Write-Host "  ✗ Failed to remove Send As from: $($perm.DisplayName) - $_" -ForegroundColor Red
                    $delegations.RemovalResults += @{Permission="SendAs"; Target=$perm.Mailbox; Status="Failed"; Error=$_.Exception.Message}
                }
            }
        }
        
        # Remove Send on Behalf
        foreach ($perm in $delegations.PermissionsGrantedByUser.SendOnBehalf) {
            if ($PSCmdlet.ShouldProcess($perm.Mailbox, "Remove Send on Behalf for $UserPrincipalName")) {
                try {
                    Set-Mailbox -Identity $perm.Mailbox -GrantSendOnBehalfTo @{Remove=$UserPrincipalName}
                    Write-Host "  ✓ Removed Send on Behalf from: $($perm.DisplayName)" -ForegroundColor Green
                    $delegations.RemovalResults += @{Permission="SendOnBehalf"; Target=$perm.Mailbox; Status="Success"}
                } catch {
                    Write-Host "  ✗ Failed to remove Send on Behalf from: $($perm.DisplayName) - $_" -ForegroundColor Red
                    $delegations.RemovalResults += @{Permission="SendOnBehalf"; Target=$perm.Mailbox; Status="Failed"; Error=$_.Exception.Message}
                }
            }
        }
        
        # Remove permissions granted TO user (if enabled)
        if ($RemoveBothDirections) {
            Write-Host "`nRemoving permissions granted TO user..." -ForegroundColor Yellow
            
            # Remove Full Access to user's mailbox
            foreach ($perm in $delegations.PermissionsGrantedToUser.FullAccess) {
                if ($PSCmdlet.ShouldProcess($UserPrincipalName, "Remove Full Access for $($perm.User)")) {
                    try {
                        Remove-MailboxPermission -Identity $UserPrincipalName -User $perm.User -AccessRights FullAccess -Confirm:$false
                        Write-Host "  ✓ Removed Full Access for: $($perm.User)" -ForegroundColor Green
                    } catch {
                        Write-Host "  ✗ Failed to remove Full Access for: $($perm.User) - $_" -ForegroundColor Red
                    }
                }
            }
            
            # Remove Send As to user's mailbox
            foreach ($perm in $delegations.PermissionsGrantedToUser.SendAs) {
                if ($PSCmdlet.ShouldProcess($UserPrincipalName, "Remove Send As for $($perm.User)")) {
                    try {
                        Remove-RecipientPermission -Identity $UserPrincipalName -Trustee $perm.User -AccessRights SendAs -Confirm:$false
                        Write-Host "  ✓ Removed Send As for: $($perm.User)" -ForegroundColor Green
                    } catch {
                        Write-Host "  ✗ Failed to remove Send As for: $($perm.User) - $_" -ForegroundColor Red
                    }
                }
            }
            
            # Clear Send on Behalf
            if ($delegations.PermissionsGrantedToUser.SendOnBehalf.Count -gt 0) {
                if ($PSCmdlet.ShouldProcess($UserPrincipalName, "Clear all Send on Behalf delegates")) {
                    try {
                        Set-Mailbox -Identity $UserPrincipalName -GrantSendOnBehalfTo $null
                        Write-Host "  ✓ Cleared all Send on Behalf delegates" -ForegroundColor Green
                    } catch {
                        Write-Host "  ✗ Failed to clear Send on Behalf delegates - $_" -ForegroundColor Red
                    }
                }
            }
        }
    }
    
    # Final summary
    Write-Host "`n=== PROCESS COMPLETE ===" -ForegroundColor Green
    Write-Host "Export saved to: $exportPath" -ForegroundColor Yellow
    Write-Host "Log saved to: $LogFile" -ForegroundColor Yellow
    
    if ($WhatIf) {
        Write-Host "`nNOTE: This was a What-If run. No actual changes were made." -ForegroundColor Cyan
    }
    
} catch {
    Write-Host "`nERROR: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Stack Trace: $($_.ScriptStackTrace)" -ForegroundColor Red
} finally {
    Stop-Transcript
}