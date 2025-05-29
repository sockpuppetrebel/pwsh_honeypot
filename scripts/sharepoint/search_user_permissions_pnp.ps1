# Search-based approach to find user's SharePoint permissions
# This uses SharePoint search which is MUCH faster than checking each site

param(
    [string]$UserEmail = "kaila.trapani@optimizely.com"
)

$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
$certPath = Join-Path $scriptPath "azure_app_cert.pem"
$keyPath = Join-Path $scriptPath "azure_app_key.pem"

# Check if certificate files exist
if (-not (Test-Path $certPath)) {
    Write-Host "Certificate file not found: $certPath" -ForegroundColor Red
    exit 1
}
if (-not (Test-Path $keyPath)) {
    Write-Host "Key file not found: $keyPath" -ForegroundColor Red
    exit 1
}

# Create X509Certificate2 from PEM files
try {
    $certContent = Get-Content $certPath -Raw
    $keyContent = Get-Content $keyPath -Raw
    
    # Create certificate from PEM
    $cert = [System.Security.Cryptography.X509Certificates.X509Certificate2]::CreateFromPem($certContent, $keyContent)
    Write-Host "Certificate loaded successfully. Thumbprint: $($cert.Thumbprint)" -ForegroundColor Green
} catch {
    Write-Host "Failed to load certificate: $_" -ForegroundColor Red
    exit 1
}

# Add certificate to store temporarily
$store = New-Object System.Security.Cryptography.X509Certificates.X509Store("My", "CurrentUser")
$store.Open("ReadWrite")
$store.Add($cert)
$store.Close()

Write-Host "`n=============== SHAREPOINT USER SEARCH ===============" -ForegroundColor Cyan
Write-Host "Searching for: $UserEmail" -ForegroundColor White

# Connect to SharePoint admin center
Write-Host "`nConnecting to SharePoint Admin Center..." -ForegroundColor Yellow
try {
    Connect-PnPOnline -Url "https://episerver99-admin.sharepoint.com" `
                      -ClientId "fe2a9efe-3000-4b02-96ea-344a2583dd52" `
                      -Tenant "3ec00d79-021a-42d4-aac8-dcb35973dff2" `
                      -Thumbprint $cert.Thumbprint
    
    Write-Host "✓ Connected to admin center" -ForegroundColor Green
} catch {
    Write-Host "✗ Failed to connect: $_" -ForegroundColor Red
    exit 1
}

# Method 1: Get user profile to find personal site and followed sites
Write-Host "`nMethod 1: Checking user profile..." -ForegroundColor Yellow
try {
    $userProfile = Get-PnPUserProfileProperty -Account $UserEmail
    if ($userProfile.PersonalUrl) {
        Write-Host "✓ Personal site (OneDrive): $($userProfile.PersonalUrl)" -ForegroundColor Green
    }
    if ($userProfile.UserName) {
        Write-Host "✓ User found in directory: $($userProfile.DisplayName)" -ForegroundColor Green
    }
} catch {
    Write-Host "Could not get user profile" -ForegroundColor Gray
}

# Method 2: Search for sites where user might have explicit permissions
Write-Host "`nMethod 2: Using SharePoint search..." -ForegroundColor Yellow
Write-Host "Searching for sites, documents, and items related to user..." -ForegroundColor Gray

# Search for sites
try {
    # Search for sites where the user is mentioned or has modified content
    $searchQuery = "Author:$UserEmail OR Editor:$UserEmail OR ModifiedBy:$UserEmail"
    $searchResults = Submit-PnPSearchQuery -Query $searchQuery -SelectProperties "Title,Path,SiteName,SiteTitle" -RowLimit 100
    
    if ($searchResults.ResultRows.Count -gt 0) {
        Write-Host "✓ Found $($searchResults.ResultRows.Count) items modified by user" -ForegroundColor Green
        
        # Extract unique sites
        $sites = $searchResults.ResultRows | 
            Where-Object { $_.Path -like "*.sharepoint.com/sites/*" } |
            ForEach-Object {
                $url = $_.Path
                if ($url -match "(https://[^/]+/sites/[^/]+)") {
                    $matches[1]
                }
            } | Select-Object -Unique
        
        if ($sites) {
            Write-Host "`nSites where user has been active:" -ForegroundColor Yellow
            $sites | ForEach-Object {
                Write-Host "  - $_" -ForegroundColor White
            }
        }
    }
} catch {
    Write-Host "Search failed: $_" -ForegroundColor Red
}

# Method 3: Get sites where user is explicitly listed (slower but thorough)
Write-Host "`nMethod 3: Checking site collection admins..." -ForegroundColor Yellow
$adminSites = @()
$siteCount = 0

Get-PnPTenantSite | Where-Object { $_.Template -notlike "REDIRECT*" } | ForEach-Object {
    $siteCount++
    if ($siteCount % 100 -eq 0) {
        Write-Host "`rChecked $siteCount sites..." -NoNewline
    }
    
    try {
        $admins = Get-PnPSiteCollectionAdmin -Url $_.Url
        if ($admins | Where-Object { $_.Email -eq $UserEmail }) {
            $adminSites += $_.Url
            Write-Host "`n✓ Site Collection Admin: $($_.Url)" -ForegroundColor Green
        }
    } catch {
        # Skip sites we can't check
    }
}

Write-Host "`n`nFound user as Site Collection Admin on $($adminSites.Count) sites" -ForegroundColor Cyan

# Method 4: Export all findings
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$reportPath = Join-Path $scriptPath "user_permissions_search_${timestamp}.txt"

$report = @"
SharePoint Permissions Search Report
User: $UserEmail
Generated: $(Get-Date)

KNOWN PERMISSIONS:
- APACOppDusk: Site Member Group
- NET5: Site Member Group

SITES TO CHECK (based on activity):
$($sites -join "`n")

SITE COLLECTION ADMIN:
$($adminSites -join "`n")

RECOMMENDATION:
1. For immediate removal from known sites, run:
   .\quick_permissions_check_pnp.ps1 -OnlyKnownSites

2. For comprehensive audit (will take hours), run overnight:
   .\remove_sharepoint_permissions_pnp.ps1 -UserEmail "$UserEmail"

3. Alternative: Use SharePoint Admin Center UI to search and remove user
"@

$report | Out-File -FilePath $reportPath
Write-Host "`nReport saved to: $reportPath" -ForegroundColor Green

# Disconnect
Disconnect-PnPOnline

# Clean up certificate
try {
    $store = New-Object System.Security.Cryptography.X509Certificates.X509Store("My", "CurrentUser")
    $store.Open("ReadWrite")
    $certToRemove = $store.Certificates | Where-Object { $_.Thumbprint -eq $cert.Thumbprint }
    if ($certToRemove) {
        $store.Remove($certToRemove)
    }
    $store.Close()
} catch { }

Write-Host "`n=============== SUMMARY ===============" -ForegroundColor Cyan
Write-Host "The user has SharePoint-only permissions (not Azure AD groups)" -ForegroundColor Yellow
Write-Host "These must be removed site-by-site using PnP PowerShell" -ForegroundColor Yellow
Write-Host "`nFastest approach:" -ForegroundColor Cyan
Write-Host "1. Remove from known sites immediately" -ForegroundColor White
Write-Host "2. Run comprehensive scan overnight" -ForegroundColor White
Write-Host "3. Or use SharePoint Admin Center UI for manual removal" -ForegroundColor White