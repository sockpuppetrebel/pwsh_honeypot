#Requires -Modules Microsoft.Graph.Groups

<#
.SYNOPSIS
    Copies members and owners from one M365 group to another
    
.DESCRIPTION
    This script creates a new Microsoft 365 group and copies all members and owners from an existing group.
    Useful for fixing incorrectly configured groups or migrating groups to new domains.
    
.PARAMETER SourceGroupEmail
    Email address of the source group to copy members from
    
.PARAMETER NewGroupDisplayName
    Display name for the new group
    
.PARAMETER NewGroupMailNickname
    Mail nickname (alias) for the new group
    
.PARAMETER GroupVisibility
    Visibility of the new group (Private or Public). Default is Private
    
.PARAMETER WhatIf
    Shows what would happen if the script runs without making actual changes
    
.EXAMPLE
    .\Copy-M365GroupMembers.ps1 -SourceGroupEmail "oldgroup@domain.com" -NewGroupDisplayName "New Group Name" -NewGroupMailNickname "newgroup"
    
.EXAMPLE
    .\Copy-M365GroupMembers.ps1 -SourceGroupEmail "erg-mentallyforward@episerver99.onmicrosoft.com" -NewGroupDisplayName "ERG Mentally Forward" -NewGroupMailNickname "erg-mentallyforward" -WhatIf
    
.NOTES
    Author: Jason Slater
    Version: 1.0
    Date: January 2025
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory=$true)]
    [ValidateNotNullOrEmpty()]
    [string]$SourceGroupEmail,
    
    [Parameter(Mandatory=$true)]
    [ValidateNotNullOrEmpty()]
    [string]$NewGroupDisplayName,
    
    [Parameter(Mandatory=$true)]
    [ValidateNotNullOrEmpty()]
    [string]$NewGroupMailNickname,
    
    [Parameter(Mandatory=$false)]
    [ValidateSet("Private", "Public")]
    [string]$GroupVisibility = "Private",
    
    [Parameter(Mandatory=$false)]
    [switch]$SkipGroupCreation
)

# Function to write colored output
function Write-ColorOutput {
    param(
        [string]$Message,
        [string]$ForegroundColor = "White"
    )
    Write-Host $Message -ForegroundColor $ForegroundColor
}

# Initialize error handling
$ErrorActionPreference = "Stop"
$script:hasErrors = $false

try {
    # Connect to Microsoft Graph if not already connected
    $context = Get-MgContext
    if (-not $context) {
        Write-ColorOutput "Connecting to Microsoft Graph..." -ForegroundColor Yellow
        Connect-MgGraph -Scopes "Group.ReadWrite.All", "Directory.ReadWrite.All" -NoWelcome
    }
    
    Write-ColorOutput "`n=== M365 GROUP MEMBER COPY UTILITY ===" -ForegroundColor Cyan
    Write-ColorOutput "Source Group: $SourceGroupEmail" -ForegroundColor White
    Write-ColorOutput "New Group Name: $NewGroupDisplayName" -ForegroundColor White
    Write-ColorOutput "New Group Alias: $NewGroupMailNickname" -ForegroundColor White
    
    # Get source group
    Write-ColorOutput "`nRetrieving source group..." -ForegroundColor Yellow
    $sourceGroup = Get-MgGroup -Filter "mail eq '$SourceGroupEmail'" -ErrorAction SilentlyContinue
    
    if (-not $sourceGroup) {
        throw "Source group with email '$SourceGroupEmail' not found"
    }
    
    $sourceGroupId = $sourceGroup.Id
    Write-ColorOutput "Source group found: $($sourceGroup.DisplayName) (ID: $sourceGroupId)" -ForegroundColor Green
    
    # Create new group if not skipping
    $newGroup = $null
    $newGroupId = $null
    
    if (-not $SkipGroupCreation) {
        if ($PSCmdlet.ShouldProcess("Microsoft 365", "Create new group '$NewGroupDisplayName'")) {
            Write-ColorOutput "`nCreating new group..." -ForegroundColor Yellow
            
            $groupParams = @{
                displayName     = $NewGroupDisplayName
                mailNickname    = $NewGroupMailNickname
                mailEnabled     = $true
                securityEnabled = $false
                groupTypes      = @("Unified")
                visibility      = $GroupVisibility
                description     = "Created from $($sourceGroup.DisplayName) on $(Get-Date -Format 'yyyy-MM-dd')"
            }
            
            try {
                $newGroup = New-MgGroup -BodyParameter $groupParams
                $newGroupId = $newGroup.Id
                Write-ColorOutput "New group created successfully (ID: $newGroupId)" -ForegroundColor Green
                
                # Wait for group to be fully provisioned
                Write-ColorOutput "Waiting for group provisioning..." -ForegroundColor Yellow
                Start-Sleep -Seconds 10
            }
            catch {
                if ($_.Exception.Message -like "*already exists*") {
                    Write-ColorOutput "Group with alias '$NewGroupMailNickname' already exists. Use -SkipGroupCreation to add members to existing group." -ForegroundColor Red
                    throw $_
                }
                else {
                    throw $_
                }
            }
        }
    }
    else {
        # Find existing group
        Write-ColorOutput "`nFinding existing target group..." -ForegroundColor Yellow
        $newGroup = Get-MgGroup -Filter "mailNickname eq '$NewGroupMailNickname'" -ErrorAction SilentlyContinue
        
        if (-not $newGroup) {
            throw "Target group with mail nickname '$NewGroupMailNickname' not found"
        }
        
        $newGroupId = $newGroup.Id
        Write-ColorOutput "Target group found: $($newGroup.DisplayName) (ID: $newGroupId)" -ForegroundColor Green
    }
    
    # Get members and owners from source group
    Write-ColorOutput "`nRetrieving source group members and owners..." -ForegroundColor Yellow
    $members = Get-MgGroupMember -GroupId $sourceGroupId -All
    $owners = Get-MgGroupOwner -GroupId $sourceGroupId -All
    
    Write-ColorOutput "Found $($members.Count) members and $($owners.Count) owners" -ForegroundColor White
    
    # Get current members of target group to avoid duplicates
    $existingMembers = Get-MgGroupMember -GroupId $newGroupId -All | Select-Object -ExpandProperty Id
    $existingOwners = Get-MgGroupOwner -GroupId $newGroupId -All | Select-Object -ExpandProperty Id
    
    # Add owners first (owners are automatically members)
    if ($owners.Count -gt 0) {
        Write-ColorOutput "`nAdding owners to new group..." -ForegroundColor Yellow
        $addedOwners = 0
        $skippedOwners = 0
        
        foreach ($owner in $owners) {
            if ($owner.Id -in $existingOwners) {
                Write-Verbose "Owner $($owner.Id) already exists, skipping"
                $skippedOwners++
                continue
            }
            
            if ($PSCmdlet.ShouldProcess($owner.Id, "Add as owner to new group")) {
                try {
                    New-MgGroupOwnerByRef -GroupId $newGroupId -BodyParameter @{
                        "@odata.id" = "https://graph.microsoft.com/v1.0/directoryObjects/$($owner.Id)"
                    }
                    $addedOwners++
                    Write-Verbose "Added owner: $($owner.Id)"
                }
                catch {
                    Write-Warning "Failed to add owner $($owner.Id): $($_.Exception.Message)"
                    $script:hasErrors = $true
                }
            }
        }
        
        Write-ColorOutput "Added $addedOwners new owners (skipped $skippedOwners existing)" -ForegroundColor Green
    }
    
    # Add members
    if ($members.Count -gt 0) {
        Write-ColorOutput "`nAdding members to new group..." -ForegroundColor Yellow
        $addedMembers = 0
        $skippedMembers = 0
        
        # Refresh existing members list to include newly added owners
        $existingMembers = Get-MgGroupMember -GroupId $newGroupId -All | Select-Object -ExpandProperty Id
        
        foreach ($member in $members) {
            if ($member.Id -in $existingMembers) {
                Write-Verbose "Member $($member.Id) already exists, skipping"
                $skippedMembers++
                continue
            }
            
            if ($PSCmdlet.ShouldProcess($member.Id, "Add as member to new group")) {
                try {
                    New-MgGroupMemberByRef -GroupId $newGroupId -BodyParameter @{
                        "@odata.id" = "https://graph.microsoft.com/v1.0/directoryObjects/$($member.Id)"
                    }
                    $addedMembers++
                    Write-Verbose "Added member: $($member.Id)"
                }
                catch {
                    Write-Warning "Failed to add member $($member.Id): $($_.Exception.Message)"
                    $script:hasErrors = $true
                }
            }
        }
        
        Write-ColorOutput "Added $addedMembers new members (skipped $skippedMembers existing)" -ForegroundColor Green
    }
    
    # Summary
    Write-ColorOutput "`n=== COPY SUMMARY ===" -ForegroundColor Cyan
    Write-ColorOutput "Source Group: $($sourceGroup.DisplayName)" -ForegroundColor White
    Write-ColorOutput "Target Group: $($newGroup.DisplayName)" -ForegroundColor White
    Write-ColorOutput "Total Members Processed: $($members.Count)" -ForegroundColor White
    Write-ColorOutput "Total Owners Processed: $($owners.Count)" -ForegroundColor White
    
    if ($script:hasErrors) {
        Write-ColorOutput "`nSome errors occurred during the copy process. Check warnings above." -ForegroundColor Yellow
    }
    else {
        Write-ColorOutput "`nAll members and owners copied successfully!" -ForegroundColor Green
    }
    
    # Next steps
    Write-ColorOutput "`n=== NEXT STEPS ===" -ForegroundColor Cyan
    Write-ColorOutput "1. Verify the new group in Outlook or Teams" -ForegroundColor White
    Write-ColorOutput "2. Test email delivery to: $NewGroupMailNickname@$($newGroup.Mail.Split('@')[1])" -ForegroundColor White
    Write-ColorOutput "3. Once verified, you can update the old group or remove it" -ForegroundColor White
    Write-ColorOutput "4. To swap aliases, use Set-MgGroup to update mailNickname" -ForegroundColor White
    
}
catch {
    Write-Error "Script failed: $_"
    exit 1
}
finally {
    # Disconnect if we connected in this script
    if (-not $context) {
        Disconnect-MgGraph -ErrorAction SilentlyContinue
    }
}