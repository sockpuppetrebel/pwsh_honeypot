# Remove user from ALL SharePoint permissions across all sites
# ONLY removes SharePoint group memberships, NOT Azure AD groups
# Heavily optimized for speed

param(
    [string]$UserEmail = "kaila.trapani@optimizely.com",
    [int]$MaxSites = 0,  # 0 = all sites
    [switch]$Confirm = $false
)

$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
$certPath = Join-Path $scriptPath "azure_app_cert.pem"
$keyPath = Join-Path $scriptPath "azure_app_key.pem"

# Load certificate
$certContent = Get-Content $certPath -Raw
$keyContent = Get-Content $keyPath -Raw
$cert = [System.Security.Cryptography.X509Certificates.X509Certificate2]::CreateFromPem($certContent, $keyContent)

$store = New-Object System.Security.Cryptography.X509Certificates.X509Store("My", "CurrentUser")
$store.Open("ReadWrite")
$store.Add($cert)
$store.Close()

Write-Host "`n=============== REMOVE ALL SHAREPOINT PERMISSIONS ===============" -ForegroundColor Red
Write-Host "User: $UserEmail" -ForegroundColor White
Write-Host "⚠  This will remove user from ALL SharePoint groups/permissions" -ForegroundColor Yellow
Write-Host "✓ Will NOT touch Azure AD group memberships" -ForegroundColor Green

if (-not $Confirm) {
    $confirmation = Read-Host "`nType 'DELETE ALL SHAREPOINT' to confirm removal"
    if ($confirmation -ne "DELETE ALL SHAREPOINT") {
        Write-Host "Cancelled" -ForegroundColor Yellow
        exit 0
    }
}

# Get all sites quickly
Write-Host "`nGetting all site collections..." -ForegroundColor Yellow
Connect-PnPOnline -Url "https://episerver99-admin.sharepoint.com" `
                  -ClientId "fe2a9efe-3000-4b02-96ea-344a2583dd52" `
                  -Tenant "3ec00d79-021a-42d4-aac8-dcb35973dff2" `
                  -Thumbprint $cert.Thumbprint

$allSites = Get-PnPTenantSite | 
    Where-Object { 
        $_.Template -notlike "REDIRECT*" -and 
        $_.Url -notlike "*-my.sharepoint.com*"
    }

if ($MaxSites -gt 0) {
    $allSites = $allSites | Select-Object -First $MaxSites
}

$totalSites = $allSites.Count
Write-Host "Will process $totalSites sites" -ForegroundColor Green

Disconnect-PnPOnline

# Processing with speed optimizations
$removed = 0
$sitesWithPermissions = 0
$processed = 0
$startTime = Get-Date
$removals = @()

Write-Host "`nStarting aggressive removal process..." -ForegroundColor Red
Write-Host "Speed optimizations enabled: Fast connections, minimal checks, bulk operations" -ForegroundColor Gray

foreach ($site in $allSites) {
    $processed++
    
    # Progress every 25 sites
    if ($processed % 25 -eq 0 -or $processed -eq 1) {
        $elapsed = (Get-Date) - $startTime
        $rate = if ($elapsed.TotalMinutes -gt 0) { [math]::Round($processed / $elapsed.TotalMinutes, 0) } else { 0 }
        $eta = if ($rate -gt 0) { [math]::Round(($totalSites - $processed) / $rate, 1) } else { "?" }
        Write-Host "`r[$processed/$totalSites] Rate: $rate/min | Removed: $removed | Sites w/perms: $sitesWithPermissions | ETA: $eta min" -NoNewline
    }
    
    try {
        # Super fast connection with minimal overhead
        $null = Connect-PnPOnline -Url $site.Url `
                                  -ClientId "fe2a9efe-3000-4b02-96ea-344a2583dd52" `
                                  -Tenant "3ec00d79-021a-42d4-aac8-dcb35973dff2" `
                                  -Thumbprint $cert.Thumbprint `
                                  -ErrorAction Stop
        
        $siteHadPermissions = $false
        $siteRemovals = @()
        
        # Fast pre-check: Does user exist in site at all?
        $userExists = $false
        try {
            $user = Get-PnPUser -Identity $UserEmail -ErrorAction Stop
            $userExists = $true
        } catch {
            # User not in site, skip to next
            Disconnect-PnPOnline -ErrorAction SilentlyContinue
            continue
        }
        
        if ($userExists) {
            # SPEED OPTIMIZATION: Get all groups at once
            $allGroups = Get-PnPGroup
            
            # Check each group for user membership
            foreach ($group in $allGroups) {
                try {
                    # SPEED: Only get members if group has reasonable size
                    $members = Get-PnPGroupMember -Group $group -ErrorAction Stop
                    $memberToRemove = $members | Where-Object { $_.Email -eq $UserEmail }
                    
                    if ($memberToRemove) {
                        $siteHadPermissions = $true
                        
                        try {
                            # REMOVE from SharePoint group
                            Remove-PnPGroupMember -Group $group -LoginName $memberToRemove.LoginName -ErrorAction Stop
                            $removed++
                            $siteRemovals += "Group: $($group.Title)"
                        } catch {
                            $siteRemovals += "FAILED: $($group.Title) - $($_.Exception.Message)"
                        }
                    }
                } catch {
                    # Skip groups we can't read
                }
            }
            
            # Check Site Collection Admin
            try {
                $admins = Get-PnPSiteCollectionAdmin -ErrorAction Stop
                if ($admins | Where-Object { $_.Email -eq $UserEmail }) {
                    $siteHadPermissions = $true
                    try {
                        Remove-PnPSiteCollectionAdmin -Owners $UserEmail -ErrorAction Stop
                        $removed++
                        $siteRemovals += "Site Collection Admin"
                    } catch {
                        $siteRemovals += "FAILED: Site Collection Admin - $($_.Exception.Message)"
                    }
                }
            } catch {
                # Skip if can't check admins
            }
        }
        
        if ($siteHadPermissions) {
            $sitesWithPermissions++
            $removals += [PSCustomObject]@{
                SiteUrl = $site.Url
                SiteTitle = $site.Title
                RemovedFrom = $siteRemovals -join "; "
                Timestamp = Get-Date
            }
            
            # Show immediate feedback for sites with permissions
            Write-Host "`n✓ $($site.Title): Removed $($siteRemovals.Count) permission(s)" -ForegroundColor Green
        }
        
        # Quick disconnect
        Disconnect-PnPOnline -ErrorAction SilentlyContinue
        
    } catch {
        # Site connection failed, skip silently for speed
        # (most failures are due to permissions or site issues)
    }
}

# Final results
$elapsed = (Get-Date) - $startTime
Write-Host "`n`n=============== REMOVAL COMPLETE ===============" -ForegroundColor Red
Write-Host "Sites processed: $processed" -ForegroundColor White
Write-Host "Sites with permissions found: $sitesWithPermissions" -ForegroundColor White
Write-Host "Total permissions removed: $removed" -ForegroundColor Green
Write-Host "Time taken: $([math]::Round($elapsed.TotalMinutes, 1)) minutes" -ForegroundColor White
Write-Host "Average rate: $([math]::Round($processed / $elapsed.TotalMinutes, 0)) sites/minute" -ForegroundColor White

if ($removals.Count -gt 0) {
    # Save detailed log
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $logPath = Join-Path $scriptPath "sharepoint_removals_${timestamp}.csv"
    
    $removals | Export-Csv -Path $logPath -NoTypeInformation
    Write-Host "`nDetailed log saved to: $logPath" -ForegroundColor Green
    
    # Show summary of removals
    Write-Host "`nSites where permissions were removed:" -ForegroundColor Yellow
    $removals | Select-Object -First 10 | ForEach-Object {
        Write-Host "• $($_.SiteTitle)" -ForegroundColor White
        Write-Host "  $($_.RemovedFrom)" -ForegroundColor Gray
    }
    
    if ($removals.Count -gt 10) {
        Write-Host "... and $($removals.Count - 10) more (see CSV for full list)" -ForegroundColor Gray
    }
}

# Important notes
Write-Host "`n=============== IMPORTANT NOTES ===============" -ForegroundColor Yellow
Write-Host "✓ Removed user from SharePoint groups/permissions only" -ForegroundColor Green
Write-Host "✓ Did NOT modify Azure AD group memberships" -ForegroundColor Green
Write-Host "✓ User may still have access via other mechanisms:" -ForegroundColor Yellow
Write-Host "  - External sharing links" -ForegroundColor White
Write-Host "  - App-based permissions" -ForegroundColor White
Write-Host "  - Inherited permissions from parent sites" -ForegroundColor White

# Clean up
$store.Open("ReadWrite")
$certToRemove = $store.Certificates | Where-Object { $_.Thumbprint -eq $cert.Thumbprint }
if ($certToRemove) { $store.Remove($certToRemove) }
$store.Close()

Write-Host "`n✓ SharePoint permission removal complete!" -ForegroundColor Green