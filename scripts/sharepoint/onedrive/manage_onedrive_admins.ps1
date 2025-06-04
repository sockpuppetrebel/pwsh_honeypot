#Requires -Modules Microsoft.Online.SharePoint.PowerShell

<#
.SYNOPSIS
    Manage OneDrive Site Collection Administrators across all OneDrive sites
.DESCRIPTION
    This script provides three functions:
    1. Add SPO admin to all OneDrive sites as site collection admin
    2. Remove specific user information from all OneDrive sites
    3. Remove SPO admin from all OneDrive sites
.NOTES
    Requires SharePoint Online Management Shell module
    Install with: Install-Module -Name Microsoft.Online.SharePoint.PowerShell
#>

# Function 1: Add SPO admin to all OneDrive site admin
function Add-SPOAdminToAllOneDrive {
    param(
        [Parameter(Mandatory=$true)]
        [string]$AdminURL,
        
        [Parameter(Mandatory=$true)]
        [string]$AdminName
    )
    
    Write-Host "Adding $AdminName as Site Collection Admin to all OneDrive sites..." -ForegroundColor Green
    
    # Connect to SharePoint Online
    Connect-SPOService -url $AdminURL
    
    $Sites = Get-SPOSite -IncludePersonalSite $true -Limit all -Filter "Url -like '-my.sharepoint.com/personal/'"
    
    Foreach ($Site in $Sites) {
        Write-host "Adding Site Collection Admin for: $($Site.URL)"
        Set-SPOUser -site $Site -LoginName $AdminName -IsSiteCollectionAdmin $True
    }
}

# Function 2: Remove user information for affected users from all OneDrive sites
function Remove-UserFromAllOneDrive {
    param(
        [Parameter(Mandatory=$true)]
        [string]$AdminURL,
        
        [Parameter(Mandatory=$true)]
        [string]$UserToRemove
    )
    
    Write-Host "Removing $UserToRemove from all OneDrive sites..." -ForegroundColor Yellow
    
    # Connect to SharePoint Online
    Connect-SPOService -Url $AdminURL
    
    # Get all OneDrive site URLs
    $oneDriveSites = Get-SPOSite -IncludePersonalSite $true -Limit all -Filter "Url -like '-my.sharepoint.com/personal/'"
    
    # Loop through each OneDrive site and remove the user
    foreach ($oneDriveSite in $oneDriveSites) {
        try {
            # Remove the user as a site collection administrator (if applicable)
            Remove-SPOUser -Site $oneDriveSite.Url -LoginName $UserToRemove
            Write-Host "Removed $UserToRemove from $($oneDriveSite.Url)"
        } catch {
            Write-Host "Error removing $UserToRemove from $($oneDriveSite.Url): $_" -ForegroundColor Red
        }
    }
}

# Function 3: Remove SPO admin from all OneDrive admin
function Remove-SPOAdminFromAllOneDrive {
    param(
        [Parameter(Mandatory=$true)]
        [string]$AdminURL,
        
        [Parameter(Mandatory=$true)]
        [string]$AdminName
    )
    
    Write-Host "Removing $AdminName as Site Collection Admin from all OneDrive sites..." -ForegroundColor Yellow
    
    # Connect to SharePoint Online
    Connect-SPOService -url $AdminURL
    
    $Sites = Get-SPOSite -IncludePersonalSite $true -Limit all -Filter "Url -like '-my.sharepoint.com/personal/'"
    
    Foreach ($Site in $Sites) {
        Write-host "Removing Site Collection Admin for: $($Site.URL)"
        Set-SPOUser -site $Site -LoginName $AdminName -IsSiteCollectionAdmin $False
    }
}

# Example usage:
<#
# Variables for your tenant
$AdminURL = "https://Yourtenant-admin.sharepoint.com/"
$AdminName = "admin@M365x09110239.onmicrosoft.com"
$UserToRemove = "user@m365x29843443.onmicrosoft.com"

# 1. Add admin to all OneDrive sites
Add-SPOAdminToAllOneDrive -AdminURL $AdminURL -AdminName $AdminName

# 2. Remove specific user from all OneDrive sites
Remove-UserFromAllOneDrive -AdminURL $AdminURL -UserToRemove $UserToRemove

# 3. Remove admin from all OneDrive sites
Remove-SPOAdminFromAllOneDrive -AdminURL $AdminURL -AdminName $AdminName
#>