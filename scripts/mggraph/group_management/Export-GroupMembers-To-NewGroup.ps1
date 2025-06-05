#Requires -Modules Microsoft.Graph.Groups, Microsoft.Graph.Users

<#
.SYNOPSIS
    Exports members from one Microsoft 365 group and adds them to another group.

.DESCRIPTION
    This script retrieves all members from a source Microsoft 365 group and adds them to a destination group.
    It handles pagination, creates a CSV export of members, and provides detailed progress reporting.

.PARAMETER SourceGroupEmail
    The email address of the source group to export members from.

.PARAMETER DestinationGroupEmail
    The email address of the destination group to add members to.

.PARAMETER ExportPath
    Optional path for the CSV export file. Defaults to ./output/mggraph/group-members-export-{timestamp}.csv

.EXAMPLE
    .\Export-GroupMembers-To-NewGroup.ps1 -SourceGroupEmail "erg-group@company.onmicrosoft.com" -DestinationGroupEmail "erg-group-new@company.com"

.NOTES
    Author: Jason Slater
    Version: 1.0
    Date: 2025-06-05
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory = $true)]
    [ValidatePattern('^[^@]+@[^@]+\.[^@]+$')]
    [string]$SourceGroupEmail,
    
    [Parameter(Mandatory = $true)]
    [ValidatePattern('^[^@]+@[^@]+\.[^@]+$')]
    [string]$DestinationGroupEmail,
    
    [Parameter(Mandatory = $false)]
    [string]$ExportPath = ""
)

# Set up paths
$timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
if ([string]::IsNullOrEmpty($ExportPath)) {
    $outputDir = Join-Path $PSScriptRoot "..\..\..\output\mggraph"
    if (-not (Test-Path $outputDir)) {
        New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
    }
    $ExportPath = Join-Path $outputDir "group-members-export-$timestamp.csv"
}

$errorLogPath = Join-Path (Split-Path $ExportPath) "migration-errors-$timestamp.log"

# Connect to Microsoft Graph
Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Yellow
try {
    Connect-MgGraph -Scopes "Group.Read.All", "Group.ReadWrite.All", "User.Read.All" -NoWelcome
}
catch {
    Write-Error "Failed to connect to Microsoft Graph: $_"
    exit 1
}

try {
    # Get source group
    Write-Host "Getting source group: $SourceGroupEmail" -ForegroundColor Yellow
    $sourceGroup = Get-MgGroup -Filter "mail eq '$SourceGroupEmail'" -Property Id,DisplayName,Mail
    
    if (-not $sourceGroup) {
        throw "Source group not found: $SourceGroupEmail"
    }
    Write-Host "Found source group: $($sourceGroup.DisplayName)" -ForegroundColor Green
    
    # Get destination group
    Write-Host "Getting destination group: $DestinationGroupEmail" -ForegroundColor Yellow
    $destGroup = Get-MgGroup -Filter "mail eq '$DestinationGroupEmail'" -Property Id,DisplayName,Mail
    
    if (-not $destGroup) {
        throw "Destination group not found: $DestinationGroupEmail"
    }
    Write-Host "Found destination group: $($destGroup.DisplayName)" -ForegroundColor Green
    
    # Get all members from source group
    Write-Host "`nFetching members from source group..." -ForegroundColor Yellow
    $sourceMembers = @()
    $pageSize = 999
    
    # Initial request
    $response = Get-MgGroupMember -GroupId $sourceGroup.Id -Top $pageSize -ConsistencyLevel eventual -CountVariable memberCount
    $sourceMembers += $response
    
    # Handle pagination
    while ($response.'@odata.nextLink') {
        Write-Host "." -NoNewline
        $response = Invoke-MgGraphRequest -Uri $response.'@odata.nextLink' -Method GET
        $sourceMembers += $response.value
    }
    
    Write-Host "`nFound $($sourceMembers.Count) members in source group" -ForegroundColor Green
    
    # Export member details to CSV
    Write-Host "`nExporting member details to CSV..." -ForegroundColor Yellow
    $memberDetails = @()
    
    foreach ($member in $sourceMembers) {
        # Get user details
        try {
            $user = Get-MgUser -UserId $member.Id -Property DisplayName,UserPrincipalName,Mail -ErrorAction Stop
            $memberDetails += [PSCustomObject]@{
                DisplayName = $user.DisplayName
                UserPrincipalName = $user.UserPrincipalName
                Email = $user.Mail
                Id = $user.Id
                Type = "User"
            }
        }
        catch {
            # Member might be a group or service principal
            $memberDetails += [PSCustomObject]@{
                DisplayName = $member.AdditionalProperties.displayName
                UserPrincipalName = "N/A"
                Email = $member.AdditionalProperties.mail
                Id = $member.Id
                Type = $member.AdditionalProperties.'@odata.type' -replace '#microsoft.graph.', ''
            }
        }
    }
    
    $memberDetails | Export-Csv -Path $ExportPath -NoTypeInformation
    Write-Host "Exported member list to: $ExportPath" -ForegroundColor Green
    
    # Get existing members of destination group to avoid duplicates
    Write-Host "`nChecking existing members in destination group..." -ForegroundColor Yellow
    $existingMembers = @()
    $response = Get-MgGroupMember -GroupId $destGroup.Id -Top $pageSize
    $existingMembers += $response
    
    while ($response.'@odata.nextLink') {
        $response = Invoke-MgGraphRequest -Uri $response.'@odata.nextLink' -Method GET
        $existingMembers += $response.value
    }
    
    $existingMemberIds = @{}
    foreach ($member in $existingMembers) {
        $existingMemberIds[$member.Id] = $true
    }
    
    # Add members to destination group
    Write-Host "`nAdding members to destination group..." -ForegroundColor Yellow
    Write-Host "Legend: . = added, s = skipped (already exists), x = failed" -ForegroundColor Cyan
    
    $successCount = 0
    $skipCount = 0
    $failureCount = 0
    $errors = @()
    
    foreach ($member in $sourceMembers) {
        if ($PSCmdlet.ShouldProcess("$($member.Id)", "Add to group $($destGroup.DisplayName)")) {
            try {
                if ($existingMemberIds.ContainsKey($member.Id)) {
                    $skipCount++
                    Write-Host "s" -NoNewline -ForegroundColor Yellow
                    continue
                }
                
                # Add member to group
                New-MgGroupMember -GroupId $destGroup.Id -DirectoryObjectId $member.Id -ErrorAction Stop
                $successCount++
                Write-Host "." -NoNewline -ForegroundColor Green
            }
            catch {
                $failureCount++
                Write-Host "x" -NoNewline -ForegroundColor Red
                
                # Get user info for error logging
                $userInfo = $memberDetails | Where-Object { $_.Id -eq $member.Id }
                if ($userInfo) {
                    $errors += "Failed to add $($userInfo.UserPrincipalName): $_"
                }
                else {
                    $errors += "Failed to add member $($member.Id): $_"
                }
            }
        }
    }
    
    Write-Host ""
    
    # Save error log if there were failures
    if ($errors.Count -gt 0) {
        $errors | Out-File -FilePath $errorLogPath
        Write-Host "`nError log saved to: $errorLogPath" -ForegroundColor Yellow
    }
    
    # Print summary
    Write-Host "`nMigration Summary:" -ForegroundColor Cyan
    Write-Host "=================" -ForegroundColor Cyan
    Write-Host "Total members in source group: $($sourceMembers.Count)" -ForegroundColor White
    Write-Host "Successfully added: $successCount members" -ForegroundColor Green
    Write-Host "Skipped (already exists): $skipCount members" -ForegroundColor Yellow
    if ($failureCount -gt 0) {
        Write-Host "Failed to add: $failureCount members" -ForegroundColor Red
        Write-Host "See error log: $errorLogPath" -ForegroundColor Red
    }
    
    Write-Host "`nMember export saved to: $ExportPath" -ForegroundColor Green
}
catch {
    Write-Error "Script failed: $_"
}
finally {
    # Disconnect from Microsoft Graph
    Disconnect-MgGraph | Out-Null
}