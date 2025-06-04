#Requires -Modules ExchangeOnlineManagement

<#
.SYNOPSIS
    Sets exact membership for a distribution list based on provided names/UPNs
    
.DESCRIPTION
    This script replaces the entire membership of a distribution list with the exact
    list provided. It will remove existing members not in the new list and add any
    missing members. Perfect for manager requests to sync DL membership with their team list.
    
.PARAMETER DistributionList
    Email address or name of the distribution list to modify
    
.PARAMETER Members
    Array of member identifiers (UPNs or Display Names)
    
.PARAMETER InputType
    Type of input provided: "UPN" or "DisplayName"
    
.PARAMETER WhatIf
    Shows what changes would be made without actually making them
    
.EXAMPLE
    .\Set-DistributionListMembership.ps1
    (Interactive mode - prompts for all inputs)
    
.EXAMPLE
    .\Set-DistributionListMembership.ps1 -DistributionList "team@company.com" -Members @("first.last@company.com") -InputType "UPN"
    
.NOTES
    Author: System Administrator
    Version: 1.0
    Requires: Exchange Online Management with appropriate permissions
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory=$false)]
    [string]$DistributionList,
    
    [Parameter(Mandatory=$false)]
    [string[]]$Members,
    
    [Parameter(Mandatory=$false)]
    [ValidateSet("UPN", "DisplayName")]
    [string]$InputType,
    
    [Parameter(Mandatory=$false)]
    [switch]$WhatIf
)

# Function to write colored output
function Write-ColorOutput {
    param(
        [string]$Message,
        [string]$ForegroundColor = "White"
    )
    Write-Host $Message -ForegroundColor $ForegroundColor
}

# Function to resolve display name to UPN using Graph
function Resolve-DisplayNameToUPN {
    param([string]$DisplayName)
    
    try {
        # Try exact match first
        $user = Get-MgUser -Filter "DisplayName eq '$DisplayName'" -ConsistencyLevel eventual -ErrorAction SilentlyContinue
        if ($user) {
            return $user.UserPrincipalName
        }
        
        # Try startswith
        $users = Get-MgUser -Filter "startswith(DisplayName,'$DisplayName')" -ConsistencyLevel eventual -ErrorAction SilentlyContinue
        if ($users.Count -eq 1) {
            return $users[0].UserPrincipalName
        }
        
        # Try contains search
        $allUsers = Get-MgUser -All -Property DisplayName,UserPrincipalName -ErrorAction SilentlyContinue
        $matchedUsers = $allUsers | Where-Object { $_.DisplayName -like "*$DisplayName*" }
        
        if ($matchedUsers.Count -eq 1) {
            return $matchedUsers[0].UserPrincipalName
        } elseif ($matchedUsers.Count -gt 1) {
            Write-ColorOutput "  Multiple matches for '$DisplayName':" -ForegroundColor Yellow
            $matchedUsers | ForEach-Object { Write-ColorOutput "    - $($_.DisplayName) ($($_.UserPrincipalName))" -ForegroundColor Gray }
            return $null
        }
        
        return $null
    }
    catch {
        Write-Warning "Error resolving '$DisplayName': $_"
        return $null
    }
}

# Connect to Exchange Online if not already connected
$exoSession = Get-PSSession | Where-Object {$_.ConfigurationName -eq "Microsoft.Exchange" -and $_.State -eq "Opened"}
if (-not $exoSession) {
    Write-ColorOutput "Connecting to Exchange Online..." -ForegroundColor Yellow
    try {
        Connect-ExchangeOnline -ShowBanner:$false
        Write-ColorOutput "Connected to Exchange Online" -ForegroundColor Green
    }
    catch {
        Write-Error "Failed to connect to Exchange Online: $_"
        exit 1
    }
}

Write-ColorOutput "`n=== DISTRIBUTION LIST MEMBERSHIP MANAGER ===" -ForegroundColor Cyan

# Interactive mode if no parameters provided
if (-not $DistributionList) {
    Write-ColorOutput "`nEnter the distribution list (email or name): " -ForegroundColor Yellow -NoNewline
    $DistributionList = Read-Host
}

# Validate distribution list exists
Write-ColorOutput "`nValidating distribution list..." -ForegroundColor Yellow
try {
    $dlInfo = Get-DistributionGroup -Identity $DistributionList -ErrorAction Stop
    Write-ColorOutput "✓ Found: $($dlInfo.DisplayName) ($($dlInfo.PrimarySmtpAddress))" -ForegroundColor Green
}
catch {
    Write-Error "Distribution list not found: $DistributionList"
    exit 1
}

# Get current members
Write-ColorOutput "`nRetrieving current members..." -ForegroundColor Yellow
$currentMembers = Get-DistributionGroupMember -Identity $DistributionList | Select-Object DisplayName, PrimarySmtpAddress
Write-ColorOutput "Current member count: $($currentMembers.Count)" -ForegroundColor White

if ($currentMembers.Count -gt 0) {
    Write-ColorOutput "Current members:" -ForegroundColor Gray
    $currentMembers | ForEach-Object { Write-ColorOutput "  - $($_.DisplayName) ($($_.PrimarySmtpAddress))" -ForegroundColor Gray }
}

# Interactive input if not provided
if (-not $InputType) {
    Write-ColorOutput "`nWhat type of list do you have?" -ForegroundColor Yellow
    Write-ColorOutput "1. Email addresses (UPNs)" -ForegroundColor White
    Write-ColorOutput "2. Display names (Full names)" -ForegroundColor White
    $choice = Read-Host "Enter choice (1 or 2)"
    
    $InputType = switch ($choice) {
        "1" { "UPN" }
        "2" { "DisplayName" }
        default { 
            Write-Error "Invalid choice. Please run script again."
            exit 1
        }
    }
}

if (-not $Members) {
    Write-ColorOutput "`nPaste your list of $($InputType.ToLower())s below and press Enter twice when done:" -ForegroundColor Yellow
    Write-ColorOutput "(Can be separated by newlines, commas, or semicolons)" -ForegroundColor Gray
    Write-Host ""
    
    # Collect input
    $inputLines = @()
    do {
        $line = Read-Host
        if ($line.Trim() -ne "") {
            $inputLines += $line
        }
    } while ($line.Trim() -ne "")
    
    if ($inputLines.Count -eq 0) {
        Write-Error "No members provided. Exiting."
        exit 1
    }
    
    # Parse the input
    $allText = $inputLines -join "`n"
    $Members = $allText -split "`n|`r`n|,|;" | 
        ForEach-Object { $_.Trim() } | 
        Where-Object { $_ -ne "" -and $_ -notmatch "^\s*$" }
}

Write-ColorOutput "`nParsed $($Members.Count) desired members:" -ForegroundColor Green
$Members | ForEach-Object { Write-ColorOutput "  - $_" -ForegroundColor White }

# Process members based on input type
$resolvedMembers = @()

if ($InputType -eq "DisplayName") {
    Write-ColorOutput "`nResolving display names to UPNs..." -ForegroundColor Yellow
    
    # Connect to Microsoft Graph for name resolution
    $mgContext = Get-MgContext
    if (-not $mgContext) {
        try {
            Connect-MgGraph -Scopes "User.Read.All" -NoWelcome
        }
        catch {
            Write-Error "Failed to connect to Microsoft Graph for name resolution: $_"
            exit 1
        }
    }
    
    foreach ($displayName in $Members) {
        Write-ColorOutput "Resolving: $displayName" -ForegroundColor Cyan
        $upn = Resolve-DisplayNameToUPN -DisplayName $displayName
        
        if ($upn) {
            $resolvedMembers += $upn
            Write-ColorOutput "  ✓ $upn" -ForegroundColor Green
        } else {
            Write-ColorOutput "  ✗ Could not resolve" -ForegroundColor Red
        }
    }
} else {
    # UPNs provided directly
    $resolvedMembers = $Members
}

if ($resolvedMembers.Count -eq 0) {
    Write-Error "No valid members could be resolved. Exiting."
    exit 1
}

Write-ColorOutput "`nResolved members: $($resolvedMembers.Count)" -ForegroundColor Green

# Compare current vs desired membership
$currentUPNs = $currentMembers.PrimarySmtpAddress
$desiredUPNs = $resolvedMembers

$toRemove = $currentUPNs | Where-Object { $_ -notin $desiredUPNs }
$toAdd = $desiredUPNs | Where-Object { $_ -notin $currentUPNs }
$staying = $currentUPNs | Where-Object { $_ -in $desiredUPNs }

# Display change summary
Write-ColorOutput "`n=== CHANGE SUMMARY ===" -ForegroundColor Cyan
Write-ColorOutput "Members staying: $($staying.Count)" -ForegroundColor Green
Write-ColorOutput "Members to remove: $($toRemove.Count)" -ForegroundColor Red
Write-ColorOutput "Members to add: $($toAdd.Count)" -ForegroundColor Yellow

if ($staying.Count -gt 0) {
    Write-ColorOutput "`nStaying (no change):" -ForegroundColor Green
    $staying | ForEach-Object { Write-ColorOutput "  ✓ $_" -ForegroundColor Green }
}

if ($toRemove.Count -gt 0) {
    Write-ColorOutput "`nTo be removed:" -ForegroundColor Red
    $toRemove | ForEach-Object { Write-ColorOutput "  ✗ $_" -ForegroundColor Red }
}

if ($toAdd.Count -gt 0) {
    Write-ColorOutput "`nTo be added:" -ForegroundColor Yellow
    $toAdd | ForEach-Object { Write-ColorOutput "  + $_" -ForegroundColor Yellow }
}

# Confirmation
if (-not $WhatIf -and ($toRemove.Count -gt 0 -or $toAdd.Count -gt 0)) {
    Write-ColorOutput "`nProceed with these changes? (Y/N): " -ForegroundColor Yellow -NoNewline
    $confirm = Read-Host
    if ($confirm -ne 'Y' -and $confirm -ne 'y') {
        Write-ColorOutput "Operation cancelled by user." -ForegroundColor Yellow
        exit 0
    }
}

# Execute changes
$removeErrors = 0
$addErrors = 0

if ($WhatIf) {
    Write-ColorOutput "`n[WHATIF] Would make the above changes" -ForegroundColor Cyan
} else {
    Write-ColorOutput "`nExecuting changes..." -ForegroundColor Yellow
    
    # Remove members
    foreach ($member in $toRemove) {
        try {
            if ($PSCmdlet.ShouldProcess($member, "Remove from $($dlInfo.DisplayName)")) {
                Remove-DistributionGroupMember -Identity $DistributionList -Member $member -Confirm:$false -ErrorAction Stop
                Write-ColorOutput "  ✓ Removed: $member" -ForegroundColor Green
            }
        }
        catch {
            Write-ColorOutput "  ✗ Failed to remove $member`: $_" -ForegroundColor Red
            $removeErrors++
        }
    }
    
    # Add members
    foreach ($member in $toAdd) {
        try {
            if ($PSCmdlet.ShouldProcess($member, "Add to $($dlInfo.DisplayName)")) {
                Add-DistributionGroupMember -Identity $DistributionList -Member $member -ErrorAction Stop
                Write-ColorOutput "  ✓ Added: $member" -ForegroundColor Green
            }
        }
        catch {
            Write-ColorOutput "  ✗ Failed to add $member`: $_" -ForegroundColor Red
            $addErrors++
        }
    }
}

# Final summary
Write-ColorOutput "`n=== FINAL SUMMARY ===" -ForegroundColor Cyan
Write-ColorOutput "Distribution List: $($dlInfo.DisplayName)" -ForegroundColor White

if ($WhatIf) {
    Write-ColorOutput "Mode: WHATIF - No changes made" -ForegroundColor Cyan
} else {
    Write-ColorOutput "Removed: $($toRemove.Count - $removeErrors) of $($toRemove.Count)" -ForegroundColor $(if ($removeErrors -eq 0) { "Green" } else { "Yellow" })
    Write-ColorOutput "Added: $($toAdd.Count - $addErrors) of $($toAdd.Count)" -ForegroundColor $(if ($addErrors -eq 0) { "Green" } else { "Yellow" })
    
    if ($removeErrors -gt 0 -or $addErrors -gt 0) {
        Write-ColorOutput "Some operations failed. Check output above for details." -ForegroundColor Red
    } else {
        Write-ColorOutput "All operations completed successfully!" -ForegroundColor Green
    }
}

Write-ColorOutput "`nFinal member count: $($resolvedMembers.Count)" -ForegroundColor White