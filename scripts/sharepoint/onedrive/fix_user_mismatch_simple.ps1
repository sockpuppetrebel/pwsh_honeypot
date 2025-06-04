# Simplified version - Remove user without adding yourself as site admin
# This uses tenant-level admin commands

$AdminURL = "https://episerver99-admin.sharepoint.com/"
$UserToRemove = "kaila.trapani@optimizely.com"

Write-Host "=== REMOVING USER FROM ALL SITES ===" -ForegroundColor Cyan
Write-Host "User to remove: $UserToRemove" -ForegroundColor Yellow
Write-Host ""

# Connect to SharePoint Online
Write-Host "Connecting to SharePoint Admin..." -ForegroundColor Yellow
Connect-SPOService -Url $AdminURL

# Option 1: Remove user from all site collections they have access to
Write-Host "Removing user from all SharePoint/OneDrive sites..." -ForegroundColor Green

# This removes the user's explicit permissions across all sites
Remove-SPOUser -LoginName $UserToRemove -Site All

Write-Host ""
Write-Host "=== COMPLETE ===" -ForegroundColor Green
Write-Host "User has been removed from all sites where they had explicit permissions." -ForegroundColor Green
Write-Host ""
Write-Host "Note: This removes explicit permissions only. Group memberships need to be handled separately." -ForegroundColor Yellow