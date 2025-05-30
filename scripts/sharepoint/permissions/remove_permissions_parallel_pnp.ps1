# Ultra-fast parallel SharePoint permissions checker/remover
# Uses PowerShell 7 parallel processing to check multiple sites simultaneously
#
# Features:
# - All SharePoint site collections (~4k sites)
# - All OneDrive personal sites (~3k sites)
# - All root sites and subsites (optional with -IncludeSubsites)
#
# Workflow:
# 1. Scans ALL sites in tenant
# 2. Continues until specified number of sites with permissions found
# 3. Stops and prompts for removal options when permissions found
# 4. Includes OneDrive sites
#
# Usage examples:
# .\remove_permissions_parallel_pnp.ps1 -IncludeSubsites
# .\remove_permissions_parallel_pnp.ps1 -UserEmail "user@domain.com"

param(
    [string]$UserEmail = "kaila.trapani@optimizely.com",
    [int]$ThrottleLimit = 10,  # Number of parallel connections
    [int]$BatchSize = 100,      # Sites per batch
    [switch]$RemovePermissions = $false,
    [int]$MaxSitesToFind = 10,  # Stop after finding this many sites
    [switch]$IncludeSubsites = $false  # Include subsites (much slower)
)

$ErrorActionPreference = 'SilentlyContinue'
$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
$projectRoot = Split-Path -Parent (Split-Path -Parent $scriptPath)
$certPath = Join-Path $projectRoot "certificates/azure/azure_app_cert.pem"
$keyPath = Join-Path $projectRoot "certificates/azure/azure_app_key.pem"

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
try {
    Connect-PnPOnline -Url "https://episerver99-admin.sharepoint.com" `
                      -ClientId "fe2a9efe-3000-4b02-96ea-344a2583dd52" `
                      -Tenant "3ec00d79-021a-42d4-aac8-dcb35973dff2" `
                      -Thumbprint $cert.Thumbprint `
                      -ErrorAction Stop
    
    Write-Host "✓ Connected to SharePoint Admin Center" -ForegroundColor Green
    
    Write-Host "Getting site collections..." -ForegroundColor Gray
    $siteCollections = Get-PnPTenantSite -IncludeOneDriveSites:$true -ErrorAction Stop | 
        Where-Object { $_.Template -notlike "REDIRECT*" } |
        Select-Object -ExpandProperty Url
    
    Write-Host "Found $($siteCollections.Count) site collections (including OneDrive)" -ForegroundColor Gray
    
    Write-Host "Getting all subsites..." -ForegroundColor Gray
    $allSites = [System.Collections.Generic.List[string]]::new()
    
    # Add all site collections
    foreach ($site in $siteCollections) {
        $allSites.Add($site)
    }
    
    # Get subsites for each site collection (excluding OneDrive sites for subsites)
    if ($IncludeSubsites) {
        $regularSites = $siteCollections | Where-Object { $_ -notlike "*-my.sharepoint.com*" }
        Write-Host "⚠️  Including subsites - this will take much longer! Scanning $($regularSites.Count) site collections..." -ForegroundColor Yellow
        
        $subsiteCount = 0
        for ($i = 0; $i -lt $regularSites.Count; $i++) {
            $siteUrl = $regularSites[$i]
            Write-Host "`r[$($i+1)/$($regularSites.Count)] Checking subsites..." -NoNewline
            
            try {
                Connect-PnPOnline -Url $siteUrl `
                                  -ClientId "fe2a9efe-3000-4b02-96ea-344a2583dd52" `
                                  -Tenant "3ec00d79-021a-42d4-aac8-dcb35973dff2" `
                                  -Thumbprint $cert.Thumbprint `
                                  -ErrorAction Stop
                
                $subsites = Get-PnPSubWeb -Recurse -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Url
                if ($subsites) {
                    foreach ($subsite in $subsites) {
                        $allSites.Add($subsite)
                        $subsiteCount++
                    }
                }
                
                Disconnect-PnPOnline -ErrorAction SilentlyContinue
            } catch {
                # Skip sites we can't access
            }
        }
        Write-Host "`nFound $subsiteCount additional subsites" -ForegroundColor Green
    } else {
        Write-Host "Skipping subsites for speed. Use -IncludeSubsites to include them." -ForegroundColor Gray
    }
    
    $allSites = $allSites | Sort-Object -Unique

    $totalSites = $allSites.Count
    Write-Host "Found $totalSites sites to check" -ForegroundColor Green
    
    if ($totalSites -eq 0) {
        Write-Host "⚠️  No sites found! This usually means:" -ForegroundColor Yellow
        Write-Host "   - App registration lacks SharePoint administrator permissions" -ForegroundColor Yellow
        Write-Host "   - Certificate/authentication issue" -ForegroundColor Yellow
        Write-Host "   - All sites are filtered out" -ForegroundColor Yellow
    }

} catch {
    Write-Host "❌ Failed to connect or get sites: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Make sure your app registration has SharePoint administrator permissions" -ForegroundColor Yellow
} finally {
    Disconnect-PnPOnline -ErrorAction SilentlyContinue
}

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
                    
                    # Don't remove in parallel - we'll do confirmation later
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
                
                # Don't remove in parallel - we'll do confirmation later
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
Write-Host "Searching for permissions until 10 sites found..." -ForegroundColor Green
Write-Host "This will scan ALL sites if needed. Press Ctrl+C to stop early." -ForegroundColor Yellow

for ($i = 0; $i -lt $allSites.Count; $i += $BatchSize) {
    $batch = $allSites[$i..[math]::Min($i + $BatchSize - 1, $allSites.Count - 1)]
    
    $batchResults = $batch | ForEach-Object -Parallel {
        $result = & $using:scriptBlock -siteUrl $_ -userEmail $using:UserEmail -thumbprint $using:cert.Thumbprint -removePerms $false
        
        if ($result) {
            Write-Host "`n✓ FOUND: $($result.SiteTitle) - $($result.Permissions.Count) permissions" -ForegroundColor Green
            return $result
        }
        
        return $null
    } -ThrottleLimit $ThrottleLimit
    
    # Add batch results to findings
    $batchResults | Where-Object { $_ -ne $null } | ForEach-Object { $findings.Add($_) }
    
    # Add processed count for each site in batch
    foreach ($site in $batch) {
        $processed.Add(1)
    }
    
    # Show progress after each batch
    $currentCount = $processed.Count
    $foundCount = $findings.Count
    Write-Host "`r[$currentCount/$totalSites] Sites scanned | Found: $foundCount sites with permissions" -NoNewline
    
    # Stop if we've found enough sites WITH PERMISSIONS
    if ($findings.Count -ge $MaxSitesToFind) {
        Write-Host "`n🎯 Found $MaxSitesToFind sites with permissions - stopping scan!" -ForegroundColor Green
        break
    }
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
            Write-Host "  - $($_.GroupType): $($_.GroupName)" -ForegroundColor Yellow
        }
    }
    
    # Always prompt for removal options when sites are found
    Write-Host "`n⚠️  REMOVAL OPTIONS ⚠️" -ForegroundColor Red
    Write-Host "Found $($findings.Count) sites with permissions for $UserEmail" -ForegroundColor Yellow
    Write-Host "`nWhat would you like to do?" -ForegroundColor Cyan
    Write-Host "1. Remove permissions from ALL $($findings.Count) sites" -ForegroundColor White
    Write-Host "2. Select specific sites to remove permissions from" -ForegroundColor White
    Write-Host "3. Exit without removing anything" -ForegroundColor White
    
    $choice = Read-Host "`nEnter your choice (1/2/3)"
    
    $sitesToProcess = @()
    
    if ($choice -eq '1') {
        $sitesToProcess = $findings
        Write-Host "`n🔥 Will remove permissions from ALL $($findings.Count) sites..." -ForegroundColor Red
    }
    elseif ($choice -eq '2') {
        Write-Host "`n📋 Select sites to remove permissions from:" -ForegroundColor Cyan
        for ($i = 0; $i -lt $findings.Count; $i++) {
            $site = $findings[$i]
            Write-Host "[$($i+1)] $($site.SiteTitle)" -ForegroundColor White
            Write-Host "     URL: $($site.SiteUrl)" -ForegroundColor Gray
            Write-Host "     Permissions: $($site.Permissions.Count)" -ForegroundColor Yellow
        }
        
        Write-Host "`nEnter site numbers to remove from (e.g., 1,3,5 or 1-3,7):" -ForegroundColor Cyan
        $selection = Read-Host "Site numbers"
        
        # Parse selection (handles ranges like 1-3 and lists like 1,3,5)
        $selectedIndices = @()
        $parts = $selection -split ','
        foreach ($part in $parts) {
            $part = $part.Trim()
            if ($part -match '(\d+)-(\d+)') {
                $start = [int]$matches[1] - 1
                $end = [int]$matches[2] - 1
                $selectedIndices += $start..$end
            } elseif ($part -match '^\d+$') {
                $selectedIndices += [int]$part - 1
            }
        }
        
        $sitesToProcess = $findings | Where-Object { $selectedIndices -contains $findings.IndexOf($_) }
        Write-Host "`n🎯 Will remove permissions from $($sitesToProcess.Count) selected sites..." -ForegroundColor Yellow
    }
    else {
        Write-Host "`n❌ Exiting without removing permissions." -ForegroundColor Yellow
        $sitesToProcess = @()
    }
    
    # Perform removal if sites were selected
    if ($sitesToProcess.Count -gt 0) {
        Write-Host "`n🔥 Starting permission removal from $($sitesToProcess.Count) sites..." -ForegroundColor Red
        $removedCount = 0
        
        foreach ($site in $sitesToProcess) {
            Write-Host "`nProcessing: $($site.SiteTitle)" -ForegroundColor Cyan
            
            try {
                # Connect to site
                Connect-PnPOnline -Url $site.SiteUrl `
                                  -ClientId "fe2a9efe-3000-4b02-96ea-344a2583dd52" `
                                  -Tenant "3ec00d79-021a-42d4-aac8-dcb35973dff2" `
                                  -Thumbprint $cert.Thumbprint `
                                  -ErrorAction Stop
                
                foreach ($permission in $site.Permissions) {
                    try {
                        if ($permission.GroupType -eq "Admin") {
                            Remove-PnPSiteCollectionAdmin -Owners $UserEmail -ErrorAction Stop
                            Write-Host "  ✓ Removed from Site Collection Admin" -ForegroundColor Green
                        } else {
                            $group = Get-PnPGroup -Identity $permission.GroupId -ErrorAction Stop
                            Remove-PnPGroupMember -Group $group -LoginName $permission.LoginName -ErrorAction Stop
                            Write-Host "  ✓ Removed from $($permission.GroupName)" -ForegroundColor Green
                        }
                        $removedCount++
                    } catch {
                        Write-Host "  ❌ Failed to remove from $($permission.GroupName): $($_.Exception.Message)" -ForegroundColor Red
                    }
                }
                
                Disconnect-PnPOnline -ErrorAction SilentlyContinue
            } catch {
                Write-Host "  ❌ Failed to connect to site: $($_.Exception.Message)" -ForegroundColor Red
            }
        }
        
        Write-Host "`n🎉 Removal complete! Removed $removedCount permissions from $($sitesToProcess.Count) sites." -ForegroundColor Green
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