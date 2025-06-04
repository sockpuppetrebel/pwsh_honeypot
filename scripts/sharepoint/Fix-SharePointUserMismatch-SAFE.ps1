# SAFE TEST VERSION - Fix SharePoint User ID Mismatch Script
# This script includes -WhatIf mode and safety checks for testing

param(
    [Parameter(Mandatory=$true)]
    [string]$AdminUrl = "https://episerver99-admin.sharepoint.com/",
    
    [Parameter(Mandatory=$true)]
    [string]$UserToRemove = "first.last@optimizely.com",
    
    [Parameter()]
    [switch]$WhatIf = $true,  # DEFAULT TO WHATIF MODE FOR SAFETY
    
    [Parameter()]
    [switch]$Force = $false,  # Must explicitly set to run for real
    
    [Parameter()]
    [int]$LimitSites = 0,  # Limit number of sites to process (0 = all)
    
    [Parameter()]
    [string[]]$TestSiteUrls = @(),  # Specific sites to test on
    
    [Parameter()]
    [switch]$SkipOneDrive = $false,  # Skip personal OneDrive sites
    
    [Parameter()]
    [switch]$GenerateReport = $true  # Generate detailed pre-execution report
)

# Safety check
if (-not $WhatIf -and -not $Force) {
    Write-Host "`nSAFETY CHECK: This script will remove user permissions!" -ForegroundColor Red
    Write-Host "To run in WhatIf mode (recommended first): Add -WhatIf" -ForegroundColor Yellow
    Write-Host "To run for real: Add -Force parameter" -ForegroundColor Yellow
    Write-Host "`nExiting for safety..." -ForegroundColor Red
    exit 1
}

# Import required module
Import-Module Microsoft.Online.SharePoint.PowerShell -Force -ErrorAction Stop

# Connect to SharePoint Online
Write-Host "Connecting to SharePoint Online..." -ForegroundColor Yellow
try {
    Connect-SPOService -Url $AdminUrl -ErrorAction Stop
    Write-Host "Successfully connected to SharePoint Online" -ForegroundColor Green
} catch {
    Write-Host "ERROR: Failed to connect to SharePoint Online: $_" -ForegroundColor Red
    exit 1
}

# Create detailed report file
$reportFile = "SharePointUserAnalysis_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
$logFile = "SharePointUserRemoval_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"

# Display current mode
if ($WhatIf) {
    Write-Host "`n*** RUNNING IN WHATIF MODE - NO CHANGES WILL BE MADE ***" -ForegroundColor Cyan -BackgroundColor DarkBlue
    Write-Host "This is a simulation to show what would happen" -ForegroundColor Cyan
} else {
    Write-Host "`n*** RUNNING IN LIVE MODE - CHANGES WILL BE MADE ***" -ForegroundColor Yellow -BackgroundColor DarkRed
    Write-Host "User permissions will be removed!" -ForegroundColor Yellow
    
    # Final confirmation
    Write-Host "`nAre you ABSOLUTELY SURE you want to proceed? Type 'YES' to continue: " -NoNewline -ForegroundColor Red
    $confirm = Read-Host
    if ($confirm -ne 'YES') {
        Write-Host "Cancelled by user" -ForegroundColor Yellow
        exit 0
    }
}

Write-Host "`nStarting analysis for user: $UserToRemove" -ForegroundColor Cyan

# Function to write to log
function Write-Log {
    param($Message, $Type = "INFO")
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "$timestamp [$Type] $Message" | Out-File -FilePath $logFile -Append
    
    switch ($Type) {
        "SUCCESS" { Write-Host $Message -ForegroundColor Green }
        "ERROR" { Write-Host $Message -ForegroundColor Red }
        "WARNING" { Write-Host $Message -ForegroundColor Yellow }
        "WHATIF" { Write-Host "[WHATIF] $Message" -ForegroundColor Cyan }
        default { Write-Host $Message }
    }
}

# Pre-flight validation
Write-Host "`nPerforming pre-flight checks..." -ForegroundColor Yellow
Write-Log "Starting pre-flight validation"

# Check if user exists in tenant
try {
    $tenantUser = Get-SPOUser -Site $AdminUrl -LoginName $UserToRemove -ErrorAction Stop
    Write-Host "✓ User found in tenant: $($tenantUser.DisplayName)" -ForegroundColor Green
    Write-Log "User validated in tenant: $($tenantUser.DisplayName)"
} catch {
    Write-Host "✗ User not found in tenant or error occurred: $_" -ForegroundColor Red
    Write-Log "User validation failed: $_" "ERROR"
    
    Write-Host "`nDo you want to continue anyway? (Y/N): " -NoNewline
    $continue = Read-Host
    if ($continue -ne 'Y' -and $continue -ne 'y') {
        exit 0
    }
}

# Initialize report data
$reportData = @()

# Get sites based on parameters
Write-Host "`nRetrieving sites..." -ForegroundColor Yellow
$sites = @()
$onedriveSites = @()

if ($TestSiteUrls.Count -gt 0) {
    # Test specific sites only
    Write-Host "Testing specific sites only..." -ForegroundColor Yellow
    foreach ($url in $TestSiteUrls) {
        try {
            $site = Get-SPOSite -Identity $url
            $sites += $site
            Write-Host "✓ Added test site: $url" -ForegroundColor Green
        } catch {
            Write-Host "✗ Could not add test site: $url - $_" -ForegroundColor Red
        }
    }
} else {
    # Get all sites
    try {
        $sites = Get-SPOSite -Limit All
        Write-Log "Retrieved $($sites.Count) regular SharePoint sites"
        
        if (-not $SkipOneDrive) {
            $onedriveSites = Get-SPOSite -IncludePersonalSite $true -Limit All | Where-Object { $_.Url -like "*-my.sharepoint.com/personal/*" }
            Write-Log "Retrieved $($onedriveSites.Count) OneDrive personal sites"
        }
    } catch {
        Write-Log "Error retrieving sites: $_" "ERROR"
        exit 1
    }
}

# Combine and limit sites if specified
$allSites = $sites + $onedriveSites
if ($LimitSites -gt 0 -and $allSites.Count -gt $LimitSites) {
    Write-Host "Limiting to first $LimitSites sites for testing..." -ForegroundColor Yellow
    $allSites = $allSites | Select-Object -First $LimitSites
}

$totalSites = $allSites.Count
Write-Host "`nTotal sites to process: $totalSites" -ForegroundColor Cyan

# Generate detailed report if requested
if ($GenerateReport) {
    Write-Host "`nGenerating detailed pre-execution report..." -ForegroundColor Yellow
    
    foreach ($site in $allSites) {
        Write-Progress -Activity "Analyzing sites" -Status "Checking: $($site.Url)" -PercentComplete (($reportData.Count / $totalSites) * 100)
        
        $siteInfo = [PSCustomObject]@{
            SiteUrl = $site.Url
            Title = $site.Title
            Template = $site.Template
            StorageUsageCurrent = $site.StorageUsageCurrent
            UserFound = $false
            UserPermissions = ""
            LastModified = $site.LastContentModifiedDate
            Status = $site.Status
        }
        
        try {
            $siteUsers = Get-SPOUser -Site $site.Url -Limit All | Where-Object { $_.LoginName -eq $UserToRemove }
            if ($siteUsers) {
                $siteInfo.UserFound = $true
                $siteInfo.UserPermissions = ($siteUsers | ForEach-Object { $_.Groups -join ";" }) -join " | "
            }
        } catch {
            $siteInfo.UserPermissions = "Error checking: $_"
        }
        
        $reportData += $siteInfo
    }
    
    # Export report
    $reportData | Export-Csv -Path $reportFile -NoTypeInformation
    Write-Host "✓ Report exported to: $reportFile" -ForegroundColor Green
    
    # Summary from report
    $sitesWithUser = ($reportData | Where-Object { $_.UserFound -eq $true }).Count
    Write-Host "`nREPORT SUMMARY:" -ForegroundColor Cyan
    Write-Host "- Total sites analyzed: $($reportData.Count)" -ForegroundColor White
    Write-Host "- Sites where user exists: $sitesWithUser" -ForegroundColor White
    Write-Host "- Sites where user not found: $($reportData.Count - $sitesWithUser)" -ForegroundColor White
}

# Process removal
Write-Host "`nStarting user removal process..." -ForegroundColor Yellow
$removedCount = 0
$errorCount = 0
$processedCount = 0

foreach ($site in $allSites) {
    $processedCount++
    $percentComplete = [math]::Round(($processedCount / $totalSites) * 100, 2)
    
    Write-Progress -Activity "Processing SharePoint sites" `
                   -Status "Processing: $($site.Url)" `
                   -PercentComplete $percentComplete `
                   -CurrentOperation "$processedCount of $totalSites sites"
    
    try {
        # Check if user exists on the site
        $siteUsers = Get-SPOUser -Site $site.Url -Limit All | Where-Object { $_.LoginName -eq $UserToRemove }
        
        if ($siteUsers) {
            if ($WhatIf) {
                Write-Log "Would remove user from: $($site.Url)" "WHATIF"
                $removedCount++
            } else {
                Remove-SPOUser -Site $site.Url -LoginName $UserToRemove
                $removedCount++
                Write-Log "Removed user from: $($site.Url)" "SUCCESS"
            }
        } else {
            Write-Verbose "User not found on site: $($site.Url)"
        }
    } catch {
        $errorMessage = $_.Exception.Message
        if ($errorMessage -like "*does not exist*" -or $errorMessage -like "*not found*") {
            Write-Verbose "User does not exist on site: $($site.Url)"
        } else {
            $errorCount++
            Write-Log "ERROR processing $($site.Url): $errorMessage" "ERROR"
        }
    }
    
    # Small delay to avoid throttling
    Start-Sleep -Milliseconds 100
}

Write-Progress -Activity "Processing SharePoint sites" -Completed

# Summary
Write-Host "`n========== SUMMARY ==========" -ForegroundColor Cyan
Write-Host "Mode: $(if ($WhatIf) { 'WHATIF (Simulation)' } else { 'LIVE (Changes Made)' })" -ForegroundColor $(if ($WhatIf) { 'Cyan' } else { 'Yellow' })
Write-Host "Total sites processed: $processedCount" -ForegroundColor White
Write-Host "User $(if ($WhatIf) { 'would be' } else { 'was' }) removed from: $removedCount sites" -ForegroundColor Green
Write-Host "Errors encountered: $errorCount" -ForegroundColor $(if ($errorCount -gt 0) { 'Red' } else { 'Green' })
Write-Host "Report file: $reportFile" -ForegroundColor White
Write-Host "Log file: $logFile" -ForegroundColor White
Write-Host "============================" -ForegroundColor Cyan

if ($WhatIf) {
    Write-Host "`nThis was a WHATIF simulation. No changes were made." -ForegroundColor Cyan
    Write-Host "To run for real, use: -Force -WhatIf:`$false" -ForegroundColor Yellow
}