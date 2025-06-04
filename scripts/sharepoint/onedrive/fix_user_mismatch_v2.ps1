# Fix User Mismatch for Kaila Trapani across all SharePoint/OneDrive sites
# Version 2 - With better error handling and connection management

# Your tenant configuration
$AdminURL = "https://episerver99-admin.sharepoint.com/"
$AdminName = "jason.slater@optimizely.com"
$UserToRemove = "kaila.trapani@optimizely.com"

Write-Host "=== STARTING USER MISMATCH FIX ===" -ForegroundColor Cyan
Write-Host "Tenant: episerver99" -ForegroundColor Yellow
Write-Host "Admin: $AdminName" -ForegroundColor Yellow
Write-Host "User to fix: $UserToRemove" -ForegroundColor Yellow
Write-Host "This will process approximately 7,500 sites" -ForegroundColor Yellow
Write-Host ""

# Test connection first
Write-Host "Testing SharePoint Online connection..." -ForegroundColor Yellow
try {
    # Disconnect any existing connections
    Disconnect-SPOService -ErrorAction SilentlyContinue
    
    # Connect with explicit authentication
    Write-Host "Connecting to SharePoint Online Admin Center..." -ForegroundColor Green
    Write-Host "Please enter your credentials when prompted." -ForegroundColor Yellow
    Connect-SPOService -Url $AdminURL -Credential (Get-Credential -UserName $AdminName -Message "Enter password for SharePoint Admin")
    
    # Test the connection
    $testSite = Get-SPOSite -Limit 1
    Write-Host "Connection successful!" -ForegroundColor Green
    Write-Host ""
} catch {
    Write-Host "ERROR: Failed to connect to SharePoint Online" -ForegroundColor Red
    Write-Host "Error details: $_" -ForegroundColor Red
    Write-Host ""
    Write-Host "Troubleshooting tips:" -ForegroundColor Yellow
    Write-Host "1. Ensure you have SharePoint Admin rights" -ForegroundColor Yellow
    Write-Host "2. Check if MFA is required (use Connect-SPOService -Url $AdminURL without -Credential parameter)" -ForegroundColor Yellow
    Write-Host "3. Verify the admin URL is correct: $AdminURL" -ForegroundColor Yellow
    exit
}

# Step 1: Add yourself as admin to all OneDrive sites
Write-Host "STEP 1: Adding admin access to all OneDrive sites..." -ForegroundColor Green
Write-Host "Fetching all OneDrive sites..." -ForegroundColor Yellow

try {
    $Sites = Get-SPOSite -IncludePersonalSite $true -Limit all -Filter "Url -like '-my.sharepoint.com/personal/'"
    $totalSites = $Sites.Count
    Write-Host "Found $totalSites OneDrive sites" -ForegroundColor Green
    
    $counter = 0
    foreach ($Site in $Sites) {
        $counter++
        Write-Progress -Activity "Adding Admin Access" -Status "Processing site $counter of $totalSites" -PercentComplete (($counter / $totalSites) * 100)
        try {
            Set-SPOUser -site $Site.Url -LoginName $AdminName -IsSiteCollectionAdmin $True -ErrorAction Stop
            Write-Host "[$counter/$totalSites] Added admin to: $($Site.Url)" -ForegroundColor Gray
        } catch {
            Write-Host "[$counter/$totalSites] ERROR adding admin to $($Site.Url): $_" -ForegroundColor Red
        }
    }
    Write-Progress -Activity "Adding Admin Access" -Completed
} catch {
    Write-Host "ERROR fetching OneDrive sites: $_" -ForegroundColor Red
    exit
}

# Step 2: Remove the problematic user from all sites
Write-Host ""
Write-Host "STEP 2: Removing $UserToRemove from all OneDrive sites..." -ForegroundColor Green

$counter = 0
foreach ($Site in $Sites) {
    $counter++
    Write-Progress -Activity "Removing User" -Status "Processing site $counter of $totalSites" -PercentComplete (($counter / $totalSites) * 100)
    try {
        Remove-SPOUser -Site $Site.Url -LoginName $UserToRemove -ErrorAction Stop
        Write-Host "[$counter/$totalSites] Removed user from: $($Site.Url)" -ForegroundColor Gray
    } catch {
        # This error is expected if user doesn't exist on the site
        if ($_.Exception.Message -notlike "*User cannot be found*") {
            Write-Host "[$counter/$totalSites] ERROR removing user from $($Site.Url): $_" -ForegroundColor Red
        }
    }
}
Write-Progress -Activity "Removing User" -Completed

# Step 3: Optional cleanup
Write-Host ""
Write-Host "STEP 3: Would you like to remove your admin access now? (Y/N)" -ForegroundColor Yellow
$response = Read-Host

if ($response -eq 'Y' -or $response -eq 'y') {
    Write-Host "Removing admin access..." -ForegroundColor Green
    $counter = 0
    foreach ($Site in $Sites) {
        $counter++
        Write-Progress -Activity "Removing Admin Access" -Status "Processing site $counter of $totalSites" -PercentComplete (($counter / $totalSites) * 100)
        try {
            Set-SPOUser -site $Site.Url -LoginName $AdminName -IsSiteCollectionAdmin $False -ErrorAction Stop
            Write-Host "[$counter/$totalSites] Removed admin from: $($Site.Url)" -ForegroundColor Gray
        } catch {
            Write-Host "[$counter/$totalSites] ERROR removing admin from $($Site.Url): $_" -ForegroundColor Red
        }
    }
    Write-Progress -Activity "Removing Admin Access" -Completed
} else {
    Write-Host "Admin access retained. Remember to remove it later." -ForegroundColor Yellow
}

Write-Host ""
Write-Host "=== PROCESS COMPLETE ===" -ForegroundColor Cyan
Write-Host "User mismatch fix has been applied across all sites." -ForegroundColor Green

# Disconnect when done
Disconnect-SPOService