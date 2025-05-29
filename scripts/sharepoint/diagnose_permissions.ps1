# Diagnostic script to understand permission issues

# Disconnect existing session
try {
    Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
} catch {}

# Load cert and connect
$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
$certPath = Join-Path $scriptPath "azure_app_cert.pem"
$keyPath = Join-Path $scriptPath "azure_app_key.pem"

$certContent = Get-Content $certPath -Raw
$keyContent = Get-Content $keyPath -Raw
$cert = [System.Security.Cryptography.X509Certificates.X509Certificate2]::CreateFromPem($certContent, $keyContent)

Connect-MgGraph -TenantId "3ec00d79-021a-42d4-aac8-dcb35973dff2" `
                -ClientId "fe2a9efe-3000-4b02-96ea-344a2583dd52" `
                -Certificate $cert `
                -NoWelcome

Write-Host "Connected!" -ForegroundColor Green

# Test with a known site
$testSiteUrl = "https://episerver99.sharepoint.com/sites/APACOppDusk"
$upn = 'kaila.trapani@optimizely.com'

Write-Host "`nTest 1: Finding the site..." -ForegroundColor Cyan
$site = Get-MgSite -Search "APACOppDusk" | Where-Object { $_.WebUrl -eq $testSiteUrl } | Select-Object -First 1

if ($site) {
    Write-Host "✓ Found site: $($site.DisplayName)" -ForegroundColor Green
    Write-Host "  ID: $($site.Id)" -ForegroundColor Gray
    
    Write-Host "`nTest 2: Getting site permissions via Graph API..." -ForegroundColor Cyan
    try {
        $permissions = Get-MgSitePermission -SiteId $site.Id -All
        Write-Host "✓ Retrieved $($permissions.Count) permissions" -ForegroundColor Green
        
        if ($permissions.Count -eq 0) {
            Write-Host "  ⚠ No permissions returned - this might be the issue!" -ForegroundColor Yellow
        }
    } catch {
        Write-Host "✗ Error getting permissions: $_" -ForegroundColor Red
    }
    
    Write-Host "`nTest 3: Trying raw Graph API call..." -ForegroundColor Cyan
    try {
        $rawPerms = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/v1.0/sites/$($site.Id)/permissions" -Method GET
        Write-Host "✓ Raw API returned $($rawPerms.value.Count) permissions" -ForegroundColor Green
        
        if ($rawPerms.value.Count -gt 0) {
            Write-Host "`nFirst permission details:" -ForegroundColor Yellow
            $rawPerms.value[0] | ConvertTo-Json -Depth 3
        }
    } catch {
        Write-Host "✗ Raw API error: $_" -ForegroundColor Red
    }
    
    Write-Host "`nTest 4: Checking site drives (document libraries)..." -ForegroundColor Cyan
    try {
        $drives = Get-MgSiteDrive -SiteId $site.Id -All
        Write-Host "✓ Found $($drives.Count) drives" -ForegroundColor Green
        
        if ($drives.Count -gt 0) {
            Write-Host "`nTest 5: Checking first drive's permissions..." -ForegroundColor Cyan
            $firstDrive = $drives[0]
            try {
                $drivePerms = Get-MgDriveRootPermission -DriveId $firstDrive.Id -All
                Write-Host "✓ Drive has $($drivePerms.Count) permissions" -ForegroundColor Green
                
                # Check for user
                $userFound = $false
                foreach ($perm in $drivePerms) {
                    if ($perm.GrantedTo.User.Email -eq $upn -or $perm.GrantedTo.User.DisplayName -eq $upn) {
                        $userFound = $true
                        Write-Host "✓ FOUND USER PERMISSION!" -ForegroundColor Green
                        Write-Host "  Permission ID: $($perm.Id)" -ForegroundColor Yellow
                        Write-Host "  Roles: $($perm.Roles -join ', ')" -ForegroundColor Yellow
                    }
                }
                
                if (-not $userFound) {
                    Write-Host "  User not found in drive permissions" -ForegroundColor Yellow
                }
            } catch {
                Write-Host "✗ Error checking drive permissions: $_" -ForegroundColor Red
            }
        }
    } catch {
        Write-Host "✗ Error getting drives: $_" -ForegroundColor Red
    }
    
    Write-Host "`nTest 6: Checking if this is a SharePoint permissions vs Graph API issue..." -ForegroundColor Cyan
    Write-Host "SharePoint permissions might be stored differently than what Graph API exposes." -ForegroundColor Yellow
    Write-Host "The Graph API site permissions might only show 'sharing' permissions, not SharePoint groups." -ForegroundColor Yellow
    
    Write-Host "`nTest 7: Trying beta endpoint..." -ForegroundColor Cyan
    try {
        $betaPerms = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/beta/sites/$($site.Id)/permissions" -Method GET
        Write-Host "✓ Beta API returned $($betaPerms.value.Count) permissions" -ForegroundColor Green
    } catch {
        Write-Host "✗ Beta API error: $_" -ForegroundColor Red
    }
    
} else {
    Write-Host "✗ Could not find site" -ForegroundColor Red
}

Write-Host "`nDIAGNOSIS:" -ForegroundColor Cyan
Write-Host "1. If permissions count is 0, Graph API might not expose SharePoint permissions" -ForegroundColor Yellow
Write-Host "2. SharePoint group memberships might not be visible via Graph API" -ForegroundColor Yellow
Write-Host "3. You might need SharePoint Application permissions (not just Delegated)" -ForegroundColor Yellow
Write-Host "4. Consider using SharePoint REST API or CSOM instead of Graph API" -ForegroundColor Yellow

Disconnect-MgGraph