#Requires -Modules ExchangeOnlineManagement

<#
.SYNOPSIS
    Quick command to add a member to a distribution group
    
.DESCRIPTION
    Simple script to add a single member to a distribution group.
    Connects to Exchange Online if not already connected.
    
.PARAMETER GroupEmail
    Email address of the distribution group
    
.PARAMETER MemberEmail
    Email address of the member to add
    
.EXAMPLE
    .\Add-DistributionGroupMember-Quick.ps1 -GroupEmail "askus@optimizely.com" -MemberEmail "first.last@optimizely.com"
    
.EXAMPLE
    .\Add-DistributionGroupMember-Quick.ps1 -GroupEmail "team@optimizely.com" -MemberEmail "john.smith@optimizely.com"
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)]
    [string]$GroupEmail,
    
    [Parameter(Mandatory=$true)]
    [string]$MemberEmail
)

# Check if already connected to Exchange Online
$connectionStatus = Get-PSSession | Where-Object {$_.ConfigurationName -eq "Microsoft.Exchange" -and $_.State -eq "Opened"}
if (-not $connectionStatus) {
    Write-Host "Connecting to Exchange Online..." -ForegroundColor Yellow
    Connect-ExchangeOnline -ShowBanner:$false
}

Write-Host "Adding $MemberEmail to $GroupEmail..." -ForegroundColor Yellow

try {
    Add-DistributionGroupMember -Identity $GroupEmail -Member $MemberEmail -ErrorAction Stop
    Write-Host "Successfully added $MemberEmail to $GroupEmail" -ForegroundColor Green
}
catch {
    if ($_.Exception.Message -like "*already a member*") {
        Write-Host "$MemberEmail is already a member of $GroupEmail" -ForegroundColor Yellow
    }
    else {
        Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
    }
}