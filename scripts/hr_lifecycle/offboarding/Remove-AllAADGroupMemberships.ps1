<#
.SYNOPSIS
    Removes all Azure AD group memberships from a user during offboarding.

.DESCRIPTION
    This script removes a user from all Azure AD groups (excluding dynamic groups).
    It creates a backup of all group memberships before removal for audit purposes.
    Includes options to exclude certain critical groups from removal.

.PARAMETER UserPrincipalName
    The UPN of the user to offboard (e.g., user@company.com)

.PARAMETER ExcludeGroups
    Array of group names or ObjectIds to exclude from removal (e.g., critical security groups)

.PARAMETER WhatIf
    Shows what groups would be removed without actually removing them

.EXAMPLE
    .\Remove-AllAADGroupMemberships.ps1 -UserPrincipalName "john.doe@company.com"

.EXAMPLE
    .\Remove-AllAADGroupMemberships.ps1 -UserPrincipalName "john.doe@company.com" -ExcludeGroups @("All Users", "VPN Access") -WhatIf

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
    [string[]]$ExcludeGroups = @(),
    
    [Parameter(Mandatory=$false)]
    [switch]$WhatIf
)

# Start transcript for logging
$LogPath = "$PSScriptRoot\Logs"
if (-not (Test-Path $LogPath)) {
    New-Item -ItemType Directory -Path $LogPath -Force | Out-Null
}
$LogFile = "$LogPath\Remove-AADGroups-$(Get-Date -Format 'yyyyMMdd-HHmmss').log"
Start-Transcript -Path $LogFile

Write-Host "=== Azure AD Group Membership Removal ===" -ForegroundColor Green
Write-Host "User: $UserPrincipalName" -ForegroundColor Yellow
Write-Host "Timestamp: $(Get-Date)" -ForegroundColor Yellow
if ($WhatIf) {
    Write-Host "MODE: What-If (No changes will be made)" -ForegroundColor Cyan
}

try {
    # Connect to Microsoft Graph if not already connected
    $mgContext = Get-MgContext -ErrorAction SilentlyContinue
    if (-not $mgContext) {
        Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Yellow
        Connect-MgGraph -Scopes "User.Read.All", "Group.ReadWrite.All", "Directory.ReadWrite.All" -NoWelcome
    }
    
    # Get the user
    Write-Host "`nStep 1: Retrieving user information..." -ForegroundColor Cyan
    $user = Get-MgUser -UserId $UserPrincipalName -ErrorAction Stop
    
    if (-not $user) {
        throw "User '$UserPrincipalName' not found!"
    }
    
    Write-Host "Found user: $($user.DisplayName) (ObjectId: $($user.Id))" -ForegroundColor Green
    
    # Get all group memberships
    Write-Host "`nStep 2: Retrieving current group memberships..." -ForegroundColor Cyan
    $groupMemberships = Get-MgUserMemberOf -UserId $user.Id -All -Property @('id', 'displayName', 'groupTypes', 'membershipRule')
    
    # Filter to only groups (exclude other directory objects)
    $groups = $groupMemberships | Where-Object { $_.'@odata.type' -eq '#microsoft.graph.group' }
    
    Write-Host "User is a member of $($groups.Count) groups" -ForegroundColor Green
    
    # Create backup of group memberships
    Write-Host "`nStep 3: Creating backup of group memberships..." -ForegroundColor Cyan
    $BackupFile = "$LogPath\GroupBackup-$($user.UserPrincipalName.Replace('@','_'))-$(Get-Date -Format 'yyyyMMdd-HHmmss').json"
    
    $groupDetails = @()
    foreach ($group in $groups) {
        $groupInfo = Get-MgGroup -GroupId $group.Id -Property @('id', 'displayName', 'groupTypes', 'membershipRule', 'mail', 'securityEnabled', 'mailEnabled')
        $groupDetails += @{
            ObjectId = $groupInfo.Id
            DisplayName = $groupInfo.DisplayName
            Mail = $groupInfo.Mail
            GroupTypes = $groupInfo.GroupTypes
            SecurityEnabled = $groupInfo.SecurityEnabled
            MailEnabled = $groupInfo.MailEnabled
            MembershipRule = $groupInfo.MembershipRule
            IsDynamic = ($groupInfo.GroupTypes -contains "DynamicMembership")
        }
    }
    
    $backupData = @{
        User = @{
            UserPrincipalName = $user.UserPrincipalName
            DisplayName = $user.DisplayName
            ObjectId = $user.Id
        }
        BackupDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        TotalGroups = $groups.Count
        Groups = $groupDetails
    }
    
    $backupData | ConvertTo-Json -Depth 10 | Out-File $BackupFile
    Write-Host "Backup saved to: $BackupFile" -ForegroundColor Green
    
    # Process group removals
    Write-Host "`nStep 4: Processing group removals..." -ForegroundColor Cyan
    
    $removedCount = 0
    $skippedCount = 0
    $failedCount = 0
    $dynamicCount = 0
    
    foreach ($group in $groupDetails) {
        $groupName = $group.DisplayName
        
        # Skip dynamic groups
        if ($group.IsDynamic) {
            Write-Host "  ! Skipping dynamic group: $groupName (cannot manually remove members)" -ForegroundColor Yellow
            $dynamicCount++
            continue
        }
        
        # Check if group is in exclude list
        if ($ExcludeGroups -contains $groupName -or $ExcludeGroups -contains $group.ObjectId) {
            Write-Host "  - Skipping excluded group: $groupName" -ForegroundColor Yellow
            $skippedCount++
            continue
        }
        
        # Remove user from group
        if ($PSCmdlet.ShouldProcess($groupName, "Remove user from group")) {
            try {
                Remove-MgGroupMemberByRef -GroupId $group.ObjectId -DirectoryObjectId $user.Id -ErrorAction Stop
                Write-Host "  ✓ Removed from: $groupName" -ForegroundColor Green
                $removedCount++
            }
            catch {
                Write-Host "  ✗ Failed to remove from: $groupName - Error: $_" -ForegroundColor Red
                $failedCount++
            }
        }
        else {
            Write-Host "  [WhatIf] Would remove from: $groupName" -ForegroundColor Cyan
        }
    }
    
    # Generate summary report
    Write-Host "`nStep 5: Generating summary report..." -ForegroundColor Cyan
    
    $summaryReport = @"
=== Group Removal Summary ===
User: $($user.DisplayName) ($UserPrincipalName)
Total Groups: $($groups.Count)
Removed: $removedCount
Skipped (Excluded): $skippedCount
Skipped (Dynamic): $dynamicCount
Failed: $failedCount

Backup File: $BackupFile
Log File: $LogFile

"@
    
    if ($WhatIf) {
        $summaryReport += "`nNOTE: This was a What-If run. No actual changes were made."
    }
    
    Write-Host $summaryReport -ForegroundColor Green
    
    # Save summary to file
    $SummaryFile = "$LogPath\RemovalSummary-$($user.UserPrincipalName.Replace('@','_'))-$(Get-Date -Format 'yyyyMMdd-HHmmss').txt"
    $summaryReport | Out-File $SummaryFile
    
    # Optional: Send notification email
    if (-not $WhatIf -and $removedCount -gt 0) {
        Write-Host "`nConsider sending notification to IT team about completed group removals." -ForegroundColor Yellow
    }
    
    Write-Host "`nGroup removal process completed!" -ForegroundColor Green
    
}
catch {
    Write-Host "`nERROR: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Stack Trace: $($_.ScriptStackTrace)" -ForegroundColor Red
}
finally {
    Stop-Transcript
}