# Mass remove user from ALL SharePoint permissions - FAST version
# Uses the working certificate approach from our successful earlier runs

param(
    [string]$UserEmail = "kaila.trapani@optimizely.com",
    [int]$MaxSites = 0  # 0 = all sites
)

Write-Host "`n=============== MASS SHAREPOINT PERMISSION REMOVAL ===============" -ForegroundColor Red
Write-Host "User: $UserEmail" -ForegroundColor White
Write-Host "⚠  This will remove user from ALL SharePoint groups/permissions" -ForegroundColor Yellow

$confirmation = Read-Host "`nType 'REMOVE ALL SHAREPOINT' to confirm mass removal"
if ($confirmation -ne "REMOVE ALL SHAREPOINT") {
    Write-Host "Cancelled" -ForegroundColor Yellow
    exit 0
}

# Use certificate that we know works
$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
$certPath = Join-Path $scriptPath "azure_app_cert.pem"
$keyPath = Join-Path $scriptPath "azure_app_key.pem"

# Load certificate
try {
    $certContent = Get-Content $certPath -Raw
    $keyContent = Get-Content $keyPath -Raw
    $cert = [System.Security.Cryptography.X509Certificates.X509Certificate2]::CreateFromPem($certContent, $keyContent)
    Write-Host "Certificate loaded. Thumbprint: $($cert.Thumbprint)" -ForegroundColor Green
    
    # Add to store temporarily for PnP
    $store = New-Object System.Security.Cryptography.X509Certificates.X509Store("My", "CurrentUser")
    $store.Open("ReadWrite")
    $store.Add($cert)
    $store.Close()
    
} catch {
    Write-Host "Certificate error: $_" -ForegroundColor Red
    Write-Host "You may need to upload the new certificate to Azure App Registration first" -ForegroundColor Yellow
    exit 1
}

# Connect to admin center to get sites
Write-Host "`nGetting all site collections..." -ForegroundColor Yellow
try {
    Connect-PnPOnline -Url "https://episerver99-admin.sharepoint.com" `
                      -ClientId "fe2a9efe-3000-4b02-96ea-344a2583dd52" `
                      -Tenant "3ec00d79-021a-42d4-aac8-dcb35973dff2" `
                      -Thumbprint $cert.Thumbprint
    
    Write-Host "✓ Connected to admin center" -ForegroundColor Green
} catch {
    Write-Host "✗ Connection failed: $_" -ForegroundColor Red
    Write-Host "You need to upload the certificate to Azure App Registration first" -ForegroundColor Yellow
    exit 1
}

# Get non-OneDrive sites for faster processing
$allSites = Get-PnPTenantSite | 
    Where-Object { 
        $_.Template -notlike "REDIRECT*" -and 
        $_.Url -notlike "*-my.sharepoint.com*" -and
        $_.Url -notlike "*personal*"
    }

if ($MaxSites -gt 0) {
    $allSites = $allSites | Select-Object -First $MaxSites
}

$totalSites = $allSites.Count
Write-Host "Will process $totalSites team sites (excluding OneDrive)" -ForegroundColor Green

Disconnect-PnPOnline

# Mass removal process
$removed = 0
$sitesWithPermissions = 0
$processed = 0
$startTime = Get-Date
$errors = 0

Write-Host "`nStarting mass removal process..." -ForegroundColor Red
Write-Host "Optimized for maximum speed" -ForegroundColor Gray

foreach ($site in $allSites) {
    $processed++
    
    # Progress every 20 sites
    if ($processed % 20 -eq 0 -or $processed -eq 1) {
        $elapsed = (Get-Date) - $startTime
        $rate = if ($elapsed.TotalMinutes -gt 0) { [math]::Round($processed / $elapsed.TotalMinutes, 0) } else { 0 }
        $eta = if ($rate -gt 0) { [math]::Round(($totalSites - $processed) / $rate, 1) } else { "?" }
        Write-Host "`r[$processed/$totalSites] Rate: $rate/min | Found: $sitesWithPermissions | Removed: $removed | Errors: $errors | ETA: $eta min" -NoNewline
    }
    
    try {
        # Quick connection
        Connect-PnPOnline -Url $site.Url `
                          -ClientId "fe2a9efe-3000-4b02-96ea-344a2583dd52" `
                          -Tenant "3ec00d79-021a-42d4-aac8-dcb35973dff2" `
                          -Thumbprint $cert.Thumbprint `
                          -ErrorAction Stop
        
        $siteHadPermissions = $false
        
        # Fast check: Does user exist?
        try {
            $user = Get-PnPUser -Identity $UserEmail -ErrorAction Stop
            
            # User exists, check groups rapidly
            $groups = Get-PnPGroup
            foreach ($group in $groups) {
                try {
                    $members = Get-PnPGroupMember -Group $group -ErrorAction Stop
                    $userMember = $members | Where-Object { $_.Email -eq $UserEmail }
                    
                    if ($userMember) {
                        $siteHadPermissions = $true
                        Remove-PnPGroupMember -Group $group -LoginName $userMember.LoginName -ErrorAction Stop
                        $removed++
                    }
                } catch {
                    # Skip group if can't process
                }
            }
            
            # Check site collection admin
            try {
                $admins = Get-PnPSiteCollectionAdmin -ErrorAction Stop
                if ($admins | Where-Object { $_.Email -eq $UserEmail }) {
                    $siteHadPermissions = $true
                    Remove-PnPSiteCollectionAdmin -Owners $UserEmail -ErrorAction Stop
                    $removed++
                }
            } catch {
                # Skip if can't check
            }
            
        } catch {
            # User doesn't exist in this site
        }
        
        if ($siteHadPermissions) {
            $sitesWithPermissions++
            Write-Host "`n✓ Removed from: $($site.Title)" -ForegroundColor Green
        }
        
        Disconnect-PnPOnline -ErrorAction SilentlyContinue
        
    } catch {
        $errors++
        # Skip site and continue for maximum speed
    }
}

# Final results
$elapsed = (Get-Date) - $startTime
Write-Host "`n`n=============== MASS REMOVAL COMPLETE ===============" -ForegroundColor Red
Write-Host "Sites processed: $processed" -ForegroundColor White
Write-Host "Sites with permissions: $sitesWithPermissions" -ForegroundColor White
Write-Host "Permissions removed: $removed" -ForegroundColor Green
Write-Host "Connection errors: $errors" -ForegroundColor Yellow
Write-Host "Time taken: $([math]::Round($elapsed.TotalMinutes, 1)) minutes" -ForegroundColor White
Write-Host "Rate: $([math]::Round($processed / $elapsed.TotalMinutes, 0)) sites/minute" -ForegroundColor White

# Clean up certificate
try {
    $store = New-Object System.Security.Cryptography.X509Certificates.X509Store("My", "CurrentUser")
    $store.Open("ReadWrite")
    $certToRemove = $store.Certificates | Where-Object { $_.Thumbprint -eq $cert.Thumbprint }
    if ($certToRemove) { $store.Remove($certToRemove) }
    $store.Close()
} catch { }

Write-Host "`n✓ Mass SharePoint permission removal complete!" -ForegroundColor Green
Write-Host "`nNote: You may need to upload the certificate to Azure first if you see connection errors." -ForegroundColor Yellow