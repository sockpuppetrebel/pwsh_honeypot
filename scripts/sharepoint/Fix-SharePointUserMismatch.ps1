# Fix SharePoint User ID Mismatch Script
# This script removes a user from all SharePoint and OneDrive sites to fix permission mismatches

param(
    [Parameter(Mandatory=$true)]
    [string]$AdminUrl = "https://episerver99-admin.sharepoint.com/",
    
    [Parameter(Mandatory=$true)]
    [string]$UserToRemove = "kaila.trapani@optimizely.com"
)

# Import required module
Import-Module Microsoft.Online.SharePoint.PowerShell -Force

# Connect to SharePoint Online
Write-Host "Connecting to SharePoint Online..." -ForegroundColor Yellow
Connect-SPOService -Url $AdminUrl

Write-Host "`nStarting user removal process for: $UserToRemove" -ForegroundColor Cyan
Write-Host "This may take several minutes depending on the number of sites..." -ForegroundColor Cyan

# Create log file
$logFile = "SharePointUserRemoval_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
$removedCount = 0
$errorCount = 0
$processedCount = 0

# Function to write to log
function Write-Log {
    param($Message, $Type = "INFO")
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "$timestamp [$Type] $Message" | Out-File -FilePath $logFile -Append
    
    switch ($Type) {
        "SUCCESS" { Write-Host $Message -ForegroundColor Green }
        "ERROR" { Write-Host $Message -ForegroundColor Red }
        "WARNING" { Write-Host $Message -ForegroundColor Yellow }
        default { Write-Host $Message }
    }
}

Write-Log "Starting SharePoint user removal for $UserToRemove"

# Get all regular SharePoint sites
Write-Host "`nRetrieving all SharePoint sites..." -ForegroundColor Yellow
$sites = @()
try {
    $sites = Get-SPOSite -Limit All
    Write-Log "Retrieved $($sites.Count) regular SharePoint sites"
} catch {
    Write-Log "Error retrieving SharePoint sites: $_" "ERROR"
}

# Get all OneDrive sites (personal sites)
Write-Host "Retrieving all OneDrive sites..." -ForegroundColor Yellow
$onedriveSites = @()
try {
    # Get all personal sites
    $onedriveSites = Get-SPOSite -IncludePersonalSite $true -Limit All | Where-Object { $_.Url -like "*-my.sharepoint.com/personal/*" }
    Write-Log "Retrieved $($onedriveSites.Count) OneDrive personal sites"
} catch {
    Write-Log "Error retrieving OneDrive sites: $_" "ERROR"
}

# Combine all sites
$allSites = $sites + $onedriveSites
$totalSites = $allSites.Count

Write-Host "`nTotal sites to process: $totalSites" -ForegroundColor Cyan
Write-Log "Total sites to process: $totalSites"

# Process each site
foreach ($site in $allSites) {
    $processedCount++
    $percentComplete = [math]::Round(($processedCount / $totalSites) * 100, 2)
    
    Write-Progress -Activity "Removing user from SharePoint sites" `
                   -Status "Processing: $($site.Url)" `
                   -PercentComplete $percentComplete `
                   -CurrentOperation "$processedCount of $totalSites sites"
    
    try {
        # First check if user exists on the site
        $siteUsers = Get-SPOUser -Site $site.Url -Limit All | Where-Object { $_.LoginName -eq $UserToRemove }
        
        if ($siteUsers) {
            # User exists, remove them
            Remove-SPOUser -Site $site.Url -LoginName $UserToRemove
            $removedCount++
            Write-Log "SUCCESS: Removed user from $($site.Url)" "SUCCESS"
        } else {
            Write-Log "User not found on site: $($site.Url)" "INFO"
        }
    } catch {
        $errorMessage = $_.Exception.Message
        
        # Check if it's a "user does not exist" error (this is ok)
        if ($errorMessage -like "*does not exist*" -or $errorMessage -like "*not found*") {
            Write-Log "User does not exist on site: $($site.Url)" "INFO"
        } else {
            $errorCount++
            Write-Log "ERROR removing user from $($site.Url): $errorMessage" "ERROR"
        }
    }
    
    # Small delay to avoid throttling
    Start-Sleep -Milliseconds 100
}

Write-Progress -Activity "Removing user from SharePoint sites" -Completed

# Summary
Write-Host "`n========== SUMMARY ==========" -ForegroundColor Cyan
Write-Host "Total sites processed: $processedCount" -ForegroundColor White
Write-Host "User removed from: $removedCount sites" -ForegroundColor Green
Write-Host "Errors encountered: $errorCount" -ForegroundColor $(if ($errorCount -gt 0) { "Red" } else { "Green" })
Write-Host "Log file: $logFile" -ForegroundColor White
Write-Host "============================" -ForegroundColor Cyan

Write-Log "Process completed. Sites processed: $processedCount, Removed from: $removedCount, Errors: $errorCount"

# Verification step
Write-Host "`nWould you like to verify the removal by checking a few sites? (Y/N): " -NoNewline
$verify = Read-Host

if ($verify -eq 'Y' -or $verify -eq 'y') {
    Write-Host "`nVerifying removal on first 5 sites..." -ForegroundColor Yellow
    $verifyCount = [Math]::Min(5, $allSites.Count)
    
    for ($i = 0; $i -lt $verifyCount; $i++) {
        $site = $allSites[$i]
        try {
            $users = Get-SPOUser -Site $site.Url -Limit All | Where-Object { $_.LoginName -eq $UserToRemove }
            if ($users) {
                Write-Host "WARNING: User still exists on $($site.Url)" -ForegroundColor Red
                Write-Log "VERIFICATION: User still exists on $($site.Url)" "WARNING"
            } else {
                Write-Host "VERIFIED: User not found on $($site.Url)" -ForegroundColor Green
                Write-Log "VERIFICATION: User successfully removed from $($site.Url)" "SUCCESS"
            }
        } catch {
            Write-Host "Could not verify $($site.Url): $_" -ForegroundColor Yellow
        }
    }
}

Write-Host "`nScript completed. Check the log file for detailed results: $logFile" -ForegroundColor Green
Write-Host "If the user still has access to some sites, you may need to:" -ForegroundColor Yellow
Write-Host "1. Wait for SharePoint to sync changes (can take up to 24 hours)" -ForegroundColor Yellow
Write-Host "2. Clear browser cache and cookies" -ForegroundColor Yellow
Write-Host "3. Run this script again" -ForegroundColor Yellow
Write-Host "4. Check for any site-specific permissions or sharing links" -ForegroundColor Yellow