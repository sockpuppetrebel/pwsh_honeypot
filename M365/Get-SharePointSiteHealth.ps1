#Requires -Modules PnP.PowerShell

<#
.SYNOPSIS
    Comprehensive SharePoint Online site health analysis and reporting
.DESCRIPTION
    Analyzes SharePoint sites for storage usage, permissions, activity, and compliance.
    Identifies sites that may need attention for governance or optimization.
.PARAMETER TenantUrl
    SharePoint tenant URL (e.g., https://contoso.sharepoint.com)
.PARAMETER IncludePersonalSites
    Include OneDrive for Business sites in analysis
.PARAMETER ExportToExcel
    Export detailed results to Excel format
.PARAMETER DaysInactive
    Days to consider a site inactive (default: 90)
.EXAMPLE
    .\Get-SharePointSiteHealth.ps1 -TenantUrl "https://contoso.sharepoint.com" -IncludePersonalSites -ExportToExcel
#>

param(
    [Parameter(Mandatory = $true)]
    [string]$TenantUrl,
    
    [Parameter(Mandatory = $false)]
    [switch]$IncludePersonalSites,
    
    [Parameter(Mandatory = $false)]
    [switch]$ExportToExcel,
    
    [Parameter(Mandatory = $false)]
    [int]$DaysInactive = 90
)

# Connect to SharePoint Online
try {
    Connect-PnPOnline -Url $TenantUrl -Interactive
    Write-Host "Connected to SharePoint Online" -ForegroundColor Green
}
catch {
    Write-Error "Failed to connect to SharePoint Online: $($_.Exception.Message)"
    exit 1
}

Write-Host "=== SHAREPOINT SITE HEALTH ANALYSIS ===" -ForegroundColor Cyan

# Get all sites
Write-Host "Retrieving site collections..." -ForegroundColor Yellow
$sites = Get-PnPTenantSite -IncludeOneDriveSites:$IncludePersonalSites

if ($sites.Count -eq 0) {
    Write-Host "No sites found" -ForegroundColor Yellow
    exit 0
}

Write-Host "Analyzing $($sites.Count) sites..." -ForegroundColor Yellow

$siteHealthData = @()
$counter = 0

foreach ($site in $sites) {
    $counter++
    Write-Progress -Activity "Analyzing sites" -Status "$($site.Title) ($counter/$($sites.Count))" -PercentComplete (($counter / $sites.Count) * 100)
    
    try {
        # Connect to the specific site
        Connect-PnPOnline -Url $site.Url -Interactive
        
        # Get site information
        $siteInfo = Get-PnPSite -Includes Usage
        $web = Get-PnPWeb -Includes Created, LastItemModifiedDate
        
        # Calculate storage metrics
        $storageUsedGB = [math]::Round($site.StorageUsageCurrent / 1024, 2)
        $storageQuotaGB = [math]::Round($site.StorageQuota / 1024, 2)
        $storagePercentUsed = if ($site.StorageQuota -gt 0) {
            [math]::Round(($site.StorageUsageCurrent / $site.StorageQuota) * 100, 2)
        } else { 0 }
        
        # Get lists and libraries
        $lists = Get-PnPList | Where-Object Hidden -eq $false
        $documentLibraries = $lists | Where-Object BaseTemplate -eq "DocumentLibrary"
        
        # Calculate activity metrics
        $daysSinceLastActivity = if ($web.LastItemModifiedDate) {
            (Get-Date) - $web.LastItemModifiedDate | Select-Object -ExpandProperty Days
        } else {
            $null
        }
        
        # Get site users
        $siteUsers = Get-PnPUser | Where-Object PrincipalType -eq "User"
        $siteGroups = Get-PnPGroup
        
        # Analyze permissions
        $uniquePermissions = 0
        try {
            $roleAssignments = Get-PnPRoleAssignment
            $uniquePermissions = $roleAssignments.Count
        }
        catch {
            # Permission enumeration failed
        }
        
        # Check for external sharing
        $externalSharingEnabled = $site.SharingCapability -ne "Disabled"
        
        # Site health assessment
        $healthIssues = @()
        
        if ($storagePercentUsed -gt 90) {
            $healthIssues += "Storage nearly full ($storagePercentUsed%)"
        }
        
        if ($daysSinceLastActivity -and $daysSinceLastActivity -gt $DaysInactive) {
            $healthIssues += "Inactive for $daysSinceLastActivity days"
        }
        
        if ($uniquePermissions -gt 50) {
            $healthIssues += "Complex permissions ($uniquePermissions unique assignments)"
        }
        
        if ($siteUsers.Count -eq 0) {
            $healthIssues += "No active users"
        }
        
        $healthScore = 100
        $healthScore -= ($healthIssues.Count * 20)
        $healthScore = [math]::Max($healthScore, 0)
        
        $siteHealth = [PSCustomObject]@{
            Title = $site.Title
            Url = $site.Url
            Owner = $site.Owner
            Template = $site.Template
            Created = $web.Created
            LastActivity = $web.LastItemModifiedDate
            DaysSinceLastActivity = $daysSinceLastActivity
            IsActive = ($daysSinceLastActivity -lt $DaysInactive)
            
            # Storage
            StorageUsedGB = $storageUsedGB
            StorageQuotaGB = $storageQuotaGB
            StoragePercentUsed = $storagePercentUsed
            
            # Content
            ListCount = $lists.Count
            DocumentLibraryCount = $documentLibraries.Count
            
            # Users and Permissions
            UserCount = $siteUsers.Count
            GroupCount = $siteGroups.Count
            UniquePermissions = $uniquePermissions
            ExternalSharingEnabled = $externalSharingEnabled
            SharingCapability = $site.SharingCapability
            
            # Compliance
            LockState = $site.LockState
            AllowSelfServiceUpgrade = $site.AllowSelfServiceUpgrade
            
            # Health Assessment
            HealthScore = $healthScore
            HealthIssues = ($healthIssues -join "; ")
            NeedsAttention = ($healthIssues.Count -gt 0)
            
            # Additional Properties
            PWAEnabled = $site.PWAEnabled
            SandboxedCodeActivationCapability = $site.SandboxedCodeActivationCapability
            ConditionalAccessPolicy = $site.ConditionalAccessPolicy
        }
        
        $siteHealthData += $siteHealth
    }
    catch {
        Write-Host "Error analyzing site $($site.Url): $($_.Exception.Message)" -ForegroundColor Yellow
        
        # Add basic info even if detailed analysis fails
        $basicHealth = [PSCustomObject]@{
            Title = $site.Title
            Url = $site.Url
            Owner = $site.Owner
            Template = $site.Template
            HealthScore = 0
            HealthIssues = "Analysis failed: $($_.Exception.Message)"
            NeedsAttention = $true
        }
        
        $siteHealthData += $basicHealth
    }
}

Write-Progress -Activity "Analyzing sites" -Completed

# Generate summary report
Write-Host "`n--- SITE HEALTH SUMMARY ---" -ForegroundColor Cyan

$totalSites = $siteHealthData.Count
$activeSites = ($siteHealthData | Where-Object IsActive -eq $true).Count
$inactiveSites = ($siteHealthData | Where-Object IsActive -eq $false).Count
$sitesNeedingAttention = ($siteHealthData | Where-Object NeedsAttention -eq $true).Count

Write-Host "Total Sites: $totalSites" -ForegroundColor White
Write-Host "Active Sites: $activeSites" -ForegroundColor Green
Write-Host "Inactive Sites (>$DaysInactive days): $inactiveSites" -ForegroundColor Yellow
Write-Host "Sites Needing Attention: $sitesNeedingAttention" -ForegroundColor Red

# Storage summary
$totalStorageUsed = ($siteHealthData | Where-Object StorageUsedGB | Measure-Object StorageUsedGB -Sum).Sum
$totalStorageQuota = ($siteHealthData | Where-Object StorageQuotaGB | Measure-Object StorageQuotaGB -Sum).Sum

Write-Host "`nStorage Overview:" -ForegroundColor Yellow
Write-Host "Total Storage Used: $([math]::Round($totalStorageUsed, 2)) GB" -ForegroundColor White
Write-Host "Total Storage Quota: $([math]::Round($totalStorageQuota, 2)) GB" -ForegroundColor White
Write-Host "Overall Storage Usage: $([math]::Round(($totalStorageUsed / $totalStorageQuota) * 100, 2))%" -ForegroundColor White

# Top issues
Write-Host "`n--- TOP SITES NEEDING ATTENTION ---" -ForegroundColor Red
$sitesNeedingAttention = $siteHealthData | Where-Object NeedsAttention -eq $true | Sort-Object HealthScore
$sitesNeedingAttention | Select-Object -First 10 | Select-Object Title, Url, HealthScore, HealthIssues | Format-Table -Wrap

# Inactive sites
$inactiveSitesList = $siteHealthData | Where-Object { $_.DaysSinceLastActivity -gt $DaysInactive } | Sort-Object DaysSinceLastActivity -Descending
if ($inactiveSitesList) {
    Write-Host "`n--- INACTIVE SITES ---" -ForegroundColor Yellow
    $inactiveSitesList | Select-Object -First 10 | Select-Object Title, Owner, DaysSinceLastActivity, StorageUsedGB | Format-Table
}

# Storage usage analysis
$highStorageSites = $siteHealthData | Where-Object StoragePercentUsed -gt 80 | Sort-Object StoragePercentUsed -Descending
if ($highStorageSites) {
    Write-Host "`n--- HIGH STORAGE USAGE SITES ---" -ForegroundColor Yellow
    $highStorageSites | Select-Object Title, StorageUsedGB, StorageQuotaGB, StoragePercentUsed | Format-Table
}

# Permission complexity analysis
$complexPermissionSites = $siteHealthData | Where-Object UniquePermissions -gt 20 | Sort-Object UniquePermissions -Descending
if ($complexPermissionSites) {
    Write-Host "`n--- SITES WITH COMPLEX PERMISSIONS ---" -ForegroundColor Yellow
    $complexPermissionSites | Select-Object -First 10 | Select-Object Title, UserCount, UniquePermissions, ExternalSharingEnabled | Format-Table
}

# Recommendations
Write-Host "`n--- RECOMMENDATIONS ---" -ForegroundColor Cyan

$recommendations = @()

if ($inactiveSites -gt 0) {
    $recommendations += "Review $inactiveSites inactive sites for archival or deletion"
}

if ($highStorageSites.Count -gt 0) {
    $recommendations += "Optimize storage for $($highStorageSites.Count) sites approaching quota limits"
}

if ($complexPermissionSites.Count -gt 0) {
    $recommendations += "Simplify permissions for $($complexPermissionSites.Count) sites with complex sharing"
}

$externalSharingSites = ($siteHealthData | Where-Object ExternalSharingEnabled -eq $true).Count
if ($externalSharingSites -gt 0) {
    $recommendations += "Review external sharing settings on $externalSharingSites sites"
}

if ($recommendations.Count -gt 0) {
    $recommendations | ForEach-Object { Write-Host "• $_" -ForegroundColor Cyan }
} else {
    Write-Host "• All sites appear to be in good health!" -ForegroundColor Green
}

# Export results
if ($ExportToExcel) {
    $timestamp = Get-Date -Format "yyyy-MM-dd-HHmm"
    $outputPath = ".\SharePoint-Site-Health-$timestamp.xlsx"
    
    if (Get-Module -ListAvailable -Name ImportExcel) {
        # Main health data
        $siteHealthData | Export-Excel -Path $outputPath -WorksheetName "Site Health Analysis" -AutoSize -FreezeTopRow
        
        # Summary statistics
        $summaryData = @(
            [PSCustomObject]@{ Metric = "Total Sites"; Value = $totalSites }
            [PSCustomObject]@{ Metric = "Active Sites"; Value = $activeSites }
            [PSCustomObject]@{ Metric = "Inactive Sites"; Value = $inactiveSites }
            [PSCustomObject]@{ Metric = "Sites Needing Attention"; Value = $sitesNeedingAttention }
            [PSCustomObject]@{ Metric = "Total Storage Used (GB)"; Value = [math]::Round($totalStorageUsed, 2) }
            [PSCustomObject]@{ Metric = "Total Storage Quota (GB)"; Value = [math]::Round($totalStorageQuota, 2) }
            [PSCustomObject]@{ Metric = "Storage Usage %"; Value = [math]::Round(($totalStorageUsed / $totalStorageQuota) * 100, 2) }
        )
        
        $summaryData | Export-Excel -Path $outputPath -WorksheetName "Summary" -AutoSize
        
        # Sites needing attention
        if ($sitesNeedingAttention.Count -gt 0) {
            $sitesNeedingAttention | Export-Excel -Path $outputPath -WorksheetName "Needs Attention" -AutoSize
        }
        
        Write-Host "`nDetailed report exported to: $outputPath" -ForegroundColor Green
    } else {
        Write-Host "ImportExcel module not available. Exporting to CSV..." -ForegroundColor Yellow
        $siteHealthData | Export-Csv -Path ".\SharePoint-Site-Health-$timestamp.csv" -NoTypeInformation
        Write-Host "Report exported to: .\SharePoint-Site-Health-$timestamp.csv" -ForegroundColor Green
    }
}

Write-Host "`n--- NEXT STEPS ---" -ForegroundColor Cyan
Write-Host "• Set up automated governance policies for site lifecycle management" -ForegroundColor White
Write-Host "• Implement storage quotas and alerts for high-usage sites" -ForegroundColor White
Write-Host "• Review and simplify complex permission structures" -ForegroundColor White
Write-Host "• Archive or delete inactive sites to optimize tenant resources" -ForegroundColor White

Disconnect-PnPOnline
Write-Host "Site health analysis complete" -ForegroundColor Green