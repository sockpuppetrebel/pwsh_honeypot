# Fix User Mismatch for users across all SharePoint/OneDrive sites
# This script will remove the problematic user account from all sites

# Your tenant configuration
$AdminURL = "https://episerver99-admin.sharepoint.com/"
$AdminName = "jason.slater@optimizely.com"
$UserToRemove = "kaila.trapani@optimizely.com"

# Import the functions
. .\manage_onedrive_admins.ps1

Write-Host "=== STARTING USER MISMATCH FIX ===" -ForegroundColor Cyan
Write-Host "Tenant: episerver99" -ForegroundColor Yellow
Write-Host "Admin: $AdminName" -ForegroundColor Yellow
Write-Host "User to fix: $UserToRemove" -ForegroundColor Yellow
Write-Host "This will process approximately 7,500 sites" -ForegroundColor Yellow
Write-Host ""

# Step 1: Add yourself as admin to all OneDrive sites (required for access)
Write-Host "STEP 1: Adding admin access to all OneDrive sites..." -ForegroundColor Green
Add-SPOAdminToAllOneDrive -AdminURL $AdminURL -AdminName $AdminName

# Step 2: Remove the problematic user from all sites
Write-Host ""
Write-Host "STEP 2: Removing $UserToRemove from all OneDrive sites to fix mismatch..." -ForegroundColor Green
Remove-UserFromAllOneDrive -AdminURL $AdminURL -UserToRemove $UserToRemove

# Step 3: Clean up - Remove your admin access (optional but recommended)
Write-Host ""
Write-Host "STEP 3: Would you like to remove your admin access now? (Y/N)" -ForegroundColor Yellow
$response = Read-Host
if ($response -eq 'Y' -or $response -eq 'y') {
    Write-Host "Removing admin access..." -ForegroundColor Green
    Remove-SPOAdminFromAllOneDrive -AdminURL $AdminURL -AdminName $AdminName
} else {
    Write-Host "Admin access retained. Remember to remove it later with:" -ForegroundColor Yellow
    Write-Host "Remove-SPOAdminFromAllOneDrive -AdminURL '$AdminURL' -AdminName '$AdminName'" -ForegroundColor Cyan
}

Write-Host ""
Write-Host "=== PROCESS COMPLETE ===" -ForegroundColor Cyan
Write-Host "User mismatch fix has been applied across all sites." -ForegroundColor Green