# Ultra-fast parallel SharePoint permissions checker/remover
# Uses PowerShell 7 parallel processing to check multiple sites simultaneously

param(
    [string]$UserEmail = "kaila.trapani@optimizely.com",
    [int]$ThrottleLimit = 10,  # Number of parallel connections
    [int]$BatchSize = 100,      # Sites per batch
    [switch]$RemovePermissions = $false
)

$ErrorActionPreference = 'SilentlyContinue'
$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
$certPath = Join-Path $scriptPath "azure_app_cert.pem"
$keyPath = Join-Path $scriptPath "azure_app_key.pem"

# Load certificate
$certContent = Get-Content $certPath -Raw
$keyContent = Get-Content $keyPath -Raw
$cert = [System.Security.Cryptography.X509Certificates.X509Certificate2]::CreateFromPem($certContent, $keyContent)

# Add to store (needed for each parallel thread)
$store = New-Object System.Security.Cryptography.X509Certificates.X509Store("My", "CurrentUser")
$store.Open("ReadWrite")
$store.Add($cert)
$store.Close()

Write-Host "`n=============== PARALLEL SHAREPOINT SCANNER ===============" -ForegroundColor Cyan
Write-Host "User: $UserEmail" -ForegroundColor White
Write-Host "Parallel threads: $ThrottleLimit" -ForegroundColor Yellow
Write-Host "Mode: $(if ($RemovePermissions) { 'REMOVE PERMISSIONS' } else { 'SCAN ONLY' })" -ForegroundColor Yellow

# Connect to admin center to get all sites
Write-Host "`nGetting all site collections..." -ForegroundColor Yellow
Connect-PnPOnline -Url "https://episerver99-admin.sharepoint.com" `
                  -ClientId "fe2a9efe-3000-4b02-96ea-344a2583dd52" `
                  -Tenant "3ec00d79-021a-42d4-aac8-dcb35973dff2" `
                  -Thumbprint $cert.Thumbprint

$allSites = Get-PnPTenantSite -IncludeOneDriveSites:$false | 
    Where-Object { $_.Template -notlike "REDIRECT*" -and $_.Url -notlike "*-my.sharepoint.com*" } |
    Select-Object -ExpandProperty Url

$totalSites = $allSites.Count
Write-Host "Found $totalSites sites to check" -ForegroundColor Green

Disconnect-PnPOnline

# Create thread-safe collections
$findings = [System.Collections.Concurrent.ConcurrentBag[PSObject]]::new()
$processed = [System.Collections.Concurrent.ConcurrentBag[int]]::new()
$startTime = Get-Date

# Progress tracking
$progressTimer = New-Object System.Timers.Timer
$progressTimer.Interval = 5000  # Update every 5 seconds
$progressTimer.AutoReset = $true
Register-ObjectEvent -InputObject $progressTimer -EventName Elapsed -Action {
    $count = $processed.Count
    $elapsed = (Get-Date) - $startTime
    $rate = [math]::Round($count / $elapsed.TotalMinutes, 1)
    $eta = if ($rate -gt 0) { [math]::Round(($totalSites - $count) / $rate, 1) } else { "?" }
    Write-Host "`r[$count/$totalSites] Sites checked | Found: $($findings.Count) | Rate: $rate/min | ETA: $eta min" -NoNewline
} | Out-Null
$progressTimer.Start()

# Parallel processing function
$scriptBlock = {
    param($siteUrl, $userEmail, $thumbprint, $removePerms)
    
    try {
        # Each thread needs its own connection
        Connect-PnPOnline -Url $siteUrl `
                          -ClientId "fe2a9efe-3000-4b02-96ea-344a2583dd52" `
                          -Tenant "3ec00d79-021a-42d4-aac8-dcb35973dff2" `
                          -Thumbprint $thumbprint `
                          -ErrorAction Stop
        
        $web = Get-PnPWeb
        $siteTitle = if ($web.Title) { $web.Title } else { Split-Path $siteUrl -Leaf }
        $permissions = @()
        
        # Quick check for user in any group
        $allGroups = Get-PnPGroup
        foreach ($group in $allGroups) {
            try {
                $members = Get-PnPGroupMember -Group $group -ErrorAction Stop
                $userInGroup = $members | Where-Object { $_.Email -eq $userEmail }
                
                if ($userInGroup) {
                    $groupType = "SharePoint Group"
                    if ($group.Id -eq $web.AssociatedOwnerGroup.Id) {
                        $groupType = "Owner Group"
                    } elseif ($group.Id -eq $web.AssociatedMemberGroup.Id) {
                        $groupType = "Member Group"
                    } elseif ($group.Id -eq $web.AssociatedVisitorGroup.Id) {
                        $groupType = "Visitor Group"
                    }
                    
                    $permissions += [PSCustomObject]@{
                        GroupName = $group.Title
                        GroupType = $groupType
                        LoginName = $userInGroup.LoginName
                        GroupId = $group.Id
                    }
                    
                    if ($removePerms) {
                        Remove-PnPGroupMember -Group $group -LoginName $userInGroup.LoginName -ErrorAction Stop
                    }
                }
            } catch {
                # Skip groups we can't read
            }
        }
        
        # Check site collection admin
        try {
            $admins = Get-PnPSiteCollectionAdmin -ErrorAction Stop
            if ($admins | Where-Object { $_.Email -eq $userEmail }) {
                $permissions += [PSCustomObject]@{
                    GroupName = "Site Collection Administrator"
                    GroupType = "Admin"
                    LoginName = $userEmail
                    GroupId = -1
                }
                
                if ($removePerms) {
                    Remove-PnPSiteCollectionAdmin -Owners $userEmail -ErrorAction Stop
                }
            }
        } catch {
            # Skip if can't check admins
        }
        
        Disconnect-PnPOnline -ErrorAction SilentlyContinue
        
        if ($permissions.Count -gt 0) {
            return [PSCustomObject]@{
                SiteUrl = $siteUrl
                SiteTitle = $siteTitle
                Permissions = $permissions
                Removed = $removePerms
            }
        }
    } catch {
        # Site connection failed, skip it
    }
    
    return $null
}

# Process sites in parallel batches
Write-Host "`nProcessing $totalSites sites with $ThrottleLimit parallel threads..." -ForegroundColor Yellow
Write-Host "This will be MUCH faster! Estimated time: $([math]::Round($totalSites / 50 / $ThrottleLimit, 1)) minutes" -ForegroundColor Green

for ($i = 0; $i -lt $allSites.Count; $i += $BatchSize) {
    $batch = $allSites[$i..[math]::Min($i + $BatchSize - 1, $allSites.Count - 1)]
    
    $batch | ForEach-Object -Parallel {
        $result = & $using:scriptBlock -siteUrl $_ -userEmail $using:UserEmail -thumbprint $using:cert.Thumbprint -removePerms $using:RemovePermissions
        
        if ($result) {
            $using:findings.Add($result)
            Write-Host "`n✓ FOUND: $($result.SiteTitle) - $($result.Permissions.Count) permissions" -ForegroundColor Green
        }
        
        $using:processed.Add(1)
    } -ThrottleLimit $ThrottleLimit
}

$progressTimer.Stop()

# Final results
$elapsed = (Get-Date) - $startTime
Write-Host "`n`n=============== RESULTS ===============" -ForegroundColor Cyan
Write-Host "Total sites checked: $($processed.Count)" -ForegroundColor White
Write-Host "Sites with permissions: $($findings.Count)" -ForegroundColor White
Write-Host "Time taken: $([math]::Round($elapsed.TotalMinutes, 1)) minutes" -ForegroundColor White
Write-Host "Average rate: $([math]::Round($processed.Count / $elapsed.TotalMinutes, 1)) sites/minute" -ForegroundColor White

if ($findings.Count -gt 0) {
    Write-Host "`nPermissions found:" -ForegroundColor Yellow
    $findings | ForEach-Object {
        Write-Host "`n• $($_.SiteTitle)" -ForegroundColor White
        Write-Host "  URL: $($_.SiteUrl)" -ForegroundColor Gray
        $_.Permissions | ForEach-Object {
            $status = if ($RemovePermissions) { " [REMOVED]" } else { "" }
            Write-Host "  - $($_.GroupType): $($_.GroupName)$status" -ForegroundColor $(if ($RemovePermissions) { "Green" } else { "Yellow" })
        }
    }
    
    # Save results
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $csvPath = Join-Path $scriptPath "parallel_scan_${timestamp}.csv"
    
    $csvData = @()
    foreach ($site in $findings) {
        foreach ($perm in $site.Permissions) {
            $csvData += [PSCustomObject]@{
                SiteUrl = $site.SiteUrl
                SiteTitle = $site.SiteTitle
                PermissionType = $perm.GroupType
                GroupName = $perm.GroupName
                Status = if ($site.Removed) { "Removed" } else { "Found" }
                Timestamp = Get-Date
            }
        }
    }
    
    $csvData | Export-Csv -Path $csvPath -NoTypeInformation
    Write-Host "`nResults saved to: $csvPath" -ForegroundColor Green
}

# Clean up cert
$store.Open("ReadWrite")
$certToRemove = $store.Certificates | Where-Object { $_.Thumbprint -eq $cert.Thumbprint }
if ($certToRemove) { $store.Remove($certToRemove) }
$store.Close()

Write-Host "`n✓ Complete!" -ForegroundColor Green

# Recommendations
if (-not $RemovePermissions -and $findings.Count -gt 0) {
    Write-Host "`n⚠ To remove these permissions, run:" -ForegroundColor Yellow
    Write-Host ".\remove_permissions_parallel_pnp.ps1 -RemovePermissions" -ForegroundColor White
}