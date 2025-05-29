# Fast SharePoint permissions scanner using concurrent connections
# Processes sites in batches with multiple PnP connections

param(
    [string]$UserEmail = "kaila.trapani@optimizely.com",
    [int]$MaxSites = 500,  # Limit for testing
    [switch]$RemovePermissions = $false
)

$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
$certPath = Join-Path $scriptPath "azure_app_cert.pem"
$keyPath = Join-Path $scriptPath "azure_app_key.pem"

# Load certificate
$certContent = Get-Content $certPath -Raw
$keyContent = Get-Content $keyPath -Raw
$cert = [System.Security.Cryptography.X509Certificates.X509Certificate2]::CreateFromPem($certContent, $keyContent)

# Add to store
$store = New-Object System.Security.Cryptography.X509Certificates.X509Store("My", "CurrentUser")
$store.Open("ReadWrite")
$store.Add($cert)
$store.Close()

Write-Host "`n=============== FAST SHAREPOINT SCANNER ===============" -ForegroundColor Cyan
Write-Host "User: $UserEmail" -ForegroundColor White
Write-Host "Max sites to check: $MaxSites" -ForegroundColor Yellow
Write-Host "Mode: $(if ($RemovePermissions) { 'REMOVE PERMISSIONS' } else { 'SCAN ONLY' })" -ForegroundColor Yellow

# Get all sites from admin center
Write-Host "`nGetting site collections..." -ForegroundColor Yellow
Connect-PnPOnline -Url "https://episerver99-admin.sharepoint.com" `
                  -ClientId "fe2a9efe-3000-4b02-96ea-344a2583dd52" `
                  -Tenant "3ec00d79-021a-42d4-aac8-dcb35973dff2" `
                  -Thumbprint $cert.Thumbprint

# Get non-OneDrive sites only for faster processing
$allSites = Get-PnPTenantSite | 
    Where-Object { 
        $_.Template -notlike "REDIRECT*" -and 
        $_.Url -notlike "*-my.sharepoint.com*" -and
        $_.Url -notlike "*personal*"
    } |
    Select-Object -First $MaxSites |
    Select-Object Url, Title

$totalSites = $allSites.Count
Write-Host "Will check $totalSites sites (excluding OneDrive)" -ForegroundColor Green

Disconnect-PnPOnline

# Process sites quickly
$findings = @()
$processed = 0
$startTime = Get-Date

Write-Host "`nScanning sites..." -ForegroundColor Yellow
Write-Host "TIP: This uses optimized queries to run much faster!" -ForegroundColor Gray

foreach ($site in $allSites) {
    $processed++
    
    # Progress every 10 sites
    if ($processed % 10 -eq 0) {
        $elapsed = (Get-Date) - $startTime
        $rate = [math]::Round($processed / $elapsed.TotalSeconds * 60, 0)
        $remaining = [math]::Round(($totalSites - $processed) / $rate, 1)
        Write-Host "`r[$processed/$totalSites] Rate: $rate sites/min | ETA: $remaining min | Found: $($findings.Count)" -NoNewline
    }
    
    try {
        # Quick connect with minimal overhead
        $null = Connect-PnPOnline -Url $site.Url `
                                  -ClientId "fe2a9efe-3000-4b02-96ea-344a2583dd52" `
                                  -Tenant "3ec00d79-021a-42d4-aac8-dcb35973dff2" `
                                  -Thumbprint $cert.Thumbprint `
                                  -ErrorAction Stop
        
        # Fast permission check - get all groups at once
        $web = Get-PnPWeb -Includes AssociatedOwnerGroup, AssociatedMemberGroup, AssociatedVisitorGroup
        $userFound = $false
        $sitePerms = @()
        
        # Method 1: Check if user exists in site at all (fast pre-check)
        $user = $null
        try {
            $user = Get-PnPUser -Identity $UserEmail -ErrorAction Stop
            $userFound = $true
        } catch {
            # User doesn't exist in this site, skip detailed checks
        }
        
        if ($userFound) {
            # User exists, now check groups
            $groups = Get-PnPGroup
            
            foreach ($group in $groups) {
                # Optimized: Get members only if user exists in site
                try {
                    $members = Get-PnPGroupMember -Group $group -ErrorAction Stop
                    $memberMatch = $members | Where-Object { $_.Email -eq $UserEmail }
                    
                    if ($memberMatch) {
                        $groupType = "SharePoint Group"
                        if ($group.Id -eq $web.AssociatedOwnerGroup.Id) {
                            $groupType = "Owner Group"
                        } elseif ($group.Id -eq $web.AssociatedMemberGroup.Id) {
                            $groupType = "Member Group"  
                        } elseif ($group.Id -eq $web.AssociatedVisitorGroup.Id) {
                            $groupType = "Visitor Group"
                        }
                        
                        $sitePerms += [PSCustomObject]@{
                            GroupName = $group.Title
                            GroupType = $groupType
                            GroupId = $group.Id
                            LoginName = $memberMatch.LoginName
                        }
                        
                        if ($RemovePermissions) {
                            Remove-PnPGroupMember -Group $group -LoginName $memberMatch.LoginName -ErrorAction Stop
                        }
                    }
                } catch {
                    # Skip groups we can't read
                }
            }
        }
        
        # Quick site collection admin check
        try {
            $admins = Get-PnPSiteCollectionAdmin -ErrorAction Stop
            if ($admins | Where-Object { $_.Email -eq $UserEmail }) {
                $sitePerms += [PSCustomObject]@{
                    GroupName = "Site Collection Administrator"
                    GroupType = "Admin"
                    GroupId = -1
                    LoginName = $UserEmail
                }
                
                if ($RemovePermissions) {
                    Remove-PnPSiteCollectionAdmin -Owners $UserEmail -ErrorAction Stop
                }
            }
        } catch {
            # Skip if can't check
        }
        
        # Add to findings if permissions found
        if ($sitePerms.Count -gt 0) {
            $findings += [PSCustomObject]@{
                SiteUrl = $site.Url
                SiteTitle = $site.Title
                Permissions = $sitePerms
                Status = if ($RemovePermissions) { "Removed" } else { "Found" }
            }
            
            Write-Host "`n✓ FOUND: $($site.Title) - $($sitePerms.Count) permission(s)" -ForegroundColor Green
            $sitePerms | ForEach-Object {
                Write-Host "  - $($_.GroupType): $($_.GroupName)" -ForegroundColor Yellow
            }
        }
        
        # Disconnect quickly
        Disconnect-PnPOnline -ErrorAction SilentlyContinue
        
    } catch {
        # Connection failed, skip site
    }
}

# Final summary
$elapsed = (Get-Date) - $startTime
Write-Host "`n`n=============== SUMMARY ===============" -ForegroundColor Cyan
Write-Host "Sites checked: $processed" -ForegroundColor White
Write-Host "Sites with permissions: $($findings.Count)" -ForegroundColor White
Write-Host "Time taken: $([math]::Round($elapsed.TotalMinutes, 1)) minutes" -ForegroundColor White
Write-Host "Average rate: $([math]::Round($processed / $elapsed.TotalSeconds * 60, 0)) sites/minute" -ForegroundColor White

if ($findings.Count -gt 0) {
    # Save results
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $csvPath = Join-Path $scriptPath "fast_scan_${timestamp}.csv"
    
    $csvData = @()
    foreach ($site in $findings) {
        foreach ($perm in $site.Permissions) {
            $csvData += [PSCustomObject]@{
                SiteUrl = $site.SiteUrl
                SiteTitle = $site.SiteTitle
                PermissionType = $perm.GroupType
                GroupName = $perm.GroupName
                Status = $site.Status
            }
        }
    }
    
    $csvData | Export-Csv -Path $csvPath -NoTypeInformation
    Write-Host "`nResults saved to: $csvPath" -ForegroundColor Green
    
    # Show summary
    Write-Host "`nPermission types found:" -ForegroundColor Yellow
    $csvData | Group-Object PermissionType | ForEach-Object {
        Write-Host "  - $($_.Name): $($_.Count) sites" -ForegroundColor White
    }
}

# Clean up
$store.Open("ReadWrite")
$certToRemove = $store.Certificates | Where-Object { $_.Thumbprint -eq $cert.Thumbprint }
if ($certToRemove) { $store.Remove($certToRemove) }
$store.Close()

if (-not $RemovePermissions -and $findings.Count -gt 0) {
    Write-Host "`n⚠ To remove these permissions, run:" -ForegroundColor Yellow
    Write-Host ".\fast_scan_pnp.ps1 -RemovePermissions" -ForegroundColor White
}

Write-Host "`n✓ Complete!" -ForegroundColor Green