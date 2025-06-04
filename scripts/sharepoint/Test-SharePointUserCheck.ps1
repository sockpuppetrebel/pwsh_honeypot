# SharePoint User Permission Check Script
# Use this to validate current user permissions before and after changes

param(
    [Parameter(Mandatory=$true)]
    [string]$AdminUrl = "https://episerver99-admin.sharepoint.com/",
    
    [Parameter(Mandatory=$true)]
    [string]$UserToCheck = "first.last@optimizely.com",
    
    [Parameter()]
    [string[]]$SiteUrls = @(),  # Specific sites to check, empty = check all
    
    [Parameter()]
    [switch]$ExportResults = $true,
    
    [Parameter()]
    [switch]$CheckUserProfile = $true
)

# Import module
Import-Module Microsoft.Online.SharePoint.PowerShell -Force

# Connect
Write-Host "Connecting to SharePoint Online..." -ForegroundColor Yellow
Connect-SPOService -Url $AdminUrl

$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$resultsFile = "UserPermissionCheck_$timestamp.csv"
$results = @()

Write-Host "`nChecking permissions for user: $UserToCheck" -ForegroundColor Cyan

# Check user profile if requested
if ($CheckUserProfile) {
    Write-Host "`nChecking user profile..." -ForegroundColor Yellow
    try {
        # Check if user exists in tenant
        $userProfile = Get-SPOUser -Site $AdminUrl -LoginName $UserToCheck
        Write-Host "✓ User Profile Found:" -ForegroundColor Green
        Write-Host "  Display Name: $($userProfile.DisplayName)" -ForegroundColor White
        Write-Host "  Login Name: $($userProfile.LoginName)" -ForegroundColor White
        Write-Host "  User ID: $($userProfile.UserId)" -ForegroundColor White
        Write-Host "  Is Site Admin: $($userProfile.IsSiteAdmin)" -ForegroundColor White
    } catch {
        Write-Host "✗ User profile not found or error: $_" -ForegroundColor Red
    }
    
    # Check for any cross-tenant scenarios
    try {
        Write-Host "`nChecking for external/guest status..." -ForegroundColor Yellow
        $externalUsers = Get-SPOExternalUser -Position 0 -PageSize 50 | Where-Object { $_.Email -eq $UserToCheck }
        if ($externalUsers) {
            Write-Host "⚠ User found in external users list!" -ForegroundColor Yellow
            Write-Host "  This might indicate cross-tenant or guest access issues" -ForegroundColor Yellow
        } else {
            Write-Host "✓ User is not in external users list (internal user)" -ForegroundColor Green
        }
    } catch {
        Write-Host "Could not check external user status" -ForegroundColor Gray
    }
}

# Get sites to check
$sitesToCheck = @()
if ($SiteUrls.Count -gt 0) {
    # Check specific sites
    foreach ($url in $SiteUrls) {
        try {
            $site = Get-SPOSite -Identity $url
            $sitesToCheck += $site
        } catch {
            Write-Host "✗ Could not retrieve site: $url" -ForegroundColor Red
        }
    }
} else {
    # Get sample of sites
    Write-Host "`nRetrieving sites for permission check..." -ForegroundColor Yellow
    $allSites = Get-SPOSite -Limit 50  # Limited for quick check
    $personalSites = Get-SPOSite -IncludePersonalSite $true -Limit All | 
                     Where-Object { $_.Url -like "*-my.sharepoint.com/personal/*" } | 
                     Select-Object -First 10
    $sitesToCheck = $allSites + $personalSites
}

Write-Host "`nChecking $($sitesToCheck.Count) sites..." -ForegroundColor Yellow

foreach ($site in $sitesToCheck) {
    Write-Host "`nChecking: $($site.Url)" -NoNewline
    
    $result = [PSCustomObject]@{
        SiteUrl = $site.Url
        SiteTitle = $site.Title
        UserFound = $false
        Permissions = "Not Found"
        Groups = ""
        LastChecked = Get-Date
        Error = ""
    }
    
    try {
        # Get all users on the site
        $siteUsers = Get-SPOUser -Site $site.Url -Limit All | Where-Object { $_.LoginName -eq $UserToCheck }
        
        if ($siteUsers) {
            $result.UserFound = $true
            $result.Permissions = $siteUsers[0].UserType
            $result.Groups = ($siteUsers[0].Groups -join ", ")
            Write-Host " - USER FOUND!" -ForegroundColor Yellow
            Write-Host "    Permissions: $($result.Permissions)" -ForegroundColor White
            if ($result.Groups) {
                Write-Host "    Groups: $($result.Groups)" -ForegroundColor White
            }
        } else {
            Write-Host " - User not found" -ForegroundColor Gray
        }
    } catch {
        $result.Error = $_.Exception.Message
        Write-Host " - Error: $_" -ForegroundColor Red
    }
    
    $results += $result
}

# Summary
$sitesWithAccess = ($results | Where-Object { $_.UserFound -eq $true }).Count
Write-Host "`n========== SUMMARY ==========" -ForegroundColor Cyan
Write-Host "Total sites checked: $($results.Count)" -ForegroundColor White
Write-Host "Sites with user access: $sitesWithAccess" -ForegroundColor $(if ($sitesWithAccess -gt 0) { 'Yellow' } else { 'Green' })
Write-Host "Sites without user: $($results.Count - $sitesWithAccess)" -ForegroundColor White

if ($ExportResults) {
    $results | Export-Csv -Path $resultsFile -NoTypeInformation
    Write-Host "`nResults exported to: $resultsFile" -ForegroundColor Green
}

# Show sites where user has access
if ($sitesWithAccess -gt 0) {
    Write-Host "`nSites where user has access:" -ForegroundColor Yellow
    $results | Where-Object { $_.UserFound -eq $true } | ForEach-Object {
        Write-Host "  - $($_.SiteUrl)" -ForegroundColor White
        if ($_.Groups) {
            Write-Host "    Groups: $($_.Groups)" -ForegroundColor Gray
        }
    }
}

Write-Host "`nPermission check complete!" -ForegroundColor Green