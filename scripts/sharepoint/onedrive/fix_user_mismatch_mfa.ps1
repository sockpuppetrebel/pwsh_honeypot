# Fix User Mismatch - MFA Version
# Run this if your account uses Multi-Factor Authentication

$AdminURL = "https://episerver99-admin.sharepoint.com/"
$AdminName = "jason.slater@optimizely.com"
$UserToRemove = "kaila.trapani@optimizely.com"

Write-Host "=== USER MISMATCH FIX (MFA Version) ===" -ForegroundColor Cyan
Write-Host ""

# Connect without credentials parameter for MFA
Write-Host "Connecting to SharePoint Online..." -ForegroundColor Yellow
Write-Host "A browser window will open for authentication." -ForegroundColor Yellow
Connect-SPOService -Url $AdminURL

Write-Host "Testing connection..." -ForegroundColor Yellow
try {
    $test = Get-SPOSite -Limit 1
    Write-Host "Connected successfully!" -ForegroundColor Green
} catch {
    Write-Host "Connection failed. Please try again." -ForegroundColor Red
    exit
}

# Now run the fix
Write-Host ""
Write-Host "Fetching all OneDrive sites..." -ForegroundColor Yellow
$Sites = Get-SPOSite -IncludePersonalSite $true -Limit all -Filter "Url -like '-my.sharepoint.com/personal/'"
Write-Host "Found $($Sites.Count) OneDrive sites" -ForegroundColor Green

# Process sites
$counter = 0
foreach ($Site in $Sites) {
    $counter++
    Write-Progress -Activity "Fixing User Mismatch" -Status "Site $counter of $($Sites.Count)" -PercentComplete (($counter / $Sites.Count) * 100)
    
    # Add admin
    try {
        Set-SPOUser -site $Site.Url -LoginName $AdminName -IsSiteCollectionAdmin $True
    } catch {
        Write-Host "Error adding admin to $($Site.Url)" -ForegroundColor Red
    }
    
    # Remove user
    try {
        Remove-SPOUser -Site $Site.Url -LoginName $UserToRemove
    } catch {
        # Ignore if user not found
    }
    
    # Remove admin
    try {
        Set-SPOUser -site $Site.Url -LoginName $AdminName -IsSiteCollectionAdmin $False
    } catch {
        Write-Host "Error removing admin from $($Site.Url)" -ForegroundColor Red
    }
}

Write-Progress -Activity "Fixing User Mismatch" -Completed
Write-Host ""
Write-Host "=== COMPLETE ===" -ForegroundColor Green
Write-Host "Processed $($Sites.Count) sites" -ForegroundColor Green