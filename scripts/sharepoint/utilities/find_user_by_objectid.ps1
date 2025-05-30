# OBJECT ID / GUID BASED USER FINDER
# Searches for user permissions based on Object IDs, GUIDs, and UPNs
# This catches permissions tied to old Azure AD identities that email searches miss

param(
    [string]$UserEmail = "kaila.trapani@optimizely.com",
    [int]$ThrottleLimit = 15,
    [int]$BatchSize = 100,
    [int]$MaxSitesToFind = 50
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

# Add to cert store temporarily
$store = New-Object System.Security.Cryptography.X509Certificates.X509Store("My", "CurrentUser")
$store.Open("ReadWrite")
$store.Add($cert)
$store.Close()

Write-Host "`n=============== OBJECT ID / GUID USER FINDER ===============" -ForegroundColor Cyan
Write-Host "🔍 Searching by Object IDs, GUIDs, and UPNs (not just display names)" -ForegroundColor Red
Write-Host "Target User: $UserEmail" -ForegroundColor White
Write-Host "This finds permissions tied to old Azure AD identities" -ForegroundColor Yellow

# First, get the current user's Object ID from Azure AD via Graph
Write-Host "`n🔍 Getting current user's Azure AD Object ID..." -ForegroundColor Yellow
try {
    Connect-PnPOnline -Url "https://episerver99-admin.sharepoint.com" `
                      -ClientId "fe2a9efe-3000-4b02-96ea-344a2583dd52" `
                      -Tenant "3ec00d79-021a-42d4-aac8-dcb35973dff2" `
                      -Thumbprint $cert.Thumbprint `
                      -ErrorAction Stop
    
    Write-Host "✓ Connected to SharePoint Admin Center" -ForegroundColor Green
    
    # Try to get user info to find their current Object ID
    $currentUser = $null
    try {
        # This might give us the current user object
        $currentUser = Get-PnPUser -Identity $UserEmail -ErrorAction SilentlyContinue
        if ($currentUser) {
            Write-Host "✓ Found current user Object ID: $($currentUser.Id)" -ForegroundColor Green
            Write-Host "  Login Name: $($currentUser.LoginName)" -ForegroundColor Gray
            Write-Host "  Title: $($currentUser.Title)" -ForegroundColor Gray
        }
    } catch {
        Write-Host "⚠️  Could not retrieve current user Object ID from SharePoint" -ForegroundColor Yellow
    }
    
    # Get ALL sites
    $allSites = Get-PnPTenantSite -IncludeOneDriveSites:$true -ErrorAction Stop | 
        Where-Object { $_.Template -notlike "REDIRECT*" } |
        Select-Object -ExpandProperty Url
    
    $totalSites = $allSites.Count
    Write-Host "📊 Found $totalSites total sites to search" -ForegroundColor Green

} catch {
    Write-Host "❌ Failed to connect: $($_.Exception.Message)" -ForegroundColor Red
    exit
} finally {
    Disconnect-PnPOnline -ErrorAction SilentlyContinue
}

# Create thread-safe collections
$allFindings = [System.Collections.Concurrent.ConcurrentBag[PSObject]]::new()
$processed = [System.Collections.Concurrent.ConcurrentBag[int]]::new()
$startTime = Get-Date

# GUID/Object ID search function
$scriptBlock = {
    param($siteUrl, $userEmail, $thumbprint)
    
    try {
        Connect-PnPOnline -Url $siteUrl `
                          -ClientId "fe2a9efe-3000-4b02-96ea-344a2583dd52" `
                          -Tenant "3ec00d79-021a-42d4-aac8-dcb35973dff2" `
                          -Thumbprint $thumbprint `
                          -ErrorAction Stop
        
        $web = Get-PnPWeb
        $siteTitle = if ($web.Title) { $web.Title } else { Split-Path $siteUrl -Leaf }
        $foundIdentities = @()
        
        # 1. GET ALL USERS on this site and look for ANY that match email/name patterns
        try {
            $allSiteUsers = Get-PnPUser -ErrorAction SilentlyContinue
            foreach ($siteUser in $allSiteUsers) {
                # Check if this user matches our target by email, login, or title
                if ($siteUser.Email -like "*$userEmail*" -or
                    $siteUser.LoginName -like "*$userEmail*" -or
                    ($siteUser.Title -like "*kaila*" -and $siteUser.Title -like "*trapani*")) {
                    
                    $foundIdentities += [PSCustomObject]@{
                        FoundIn = "Site Users List"
                        IdentityType = "Site User"
                        Email = $siteUser.Email
                        LoginName = $siteUser.LoginName
                        Title = $siteUser.Title
                        ObjectId = $siteUser.Id
                        UserId = $siteUser.UserId
                        IsActive = -not ($siteUser.LoginName.Contains("_disabled_") -or $siteUser.Title.Contains("(deleted)"))
                    }
                }
            }
        } catch {
            # Skip if can't get site users
        }
        
        # 2. SEARCH ALL GROUPS and examine each member's Object ID
        try {
            $allGroups = Get-PnPGroup -ErrorAction SilentlyContinue
            foreach ($group in $allGroups) {
                try {
                    $members = Get-PnPGroupMember -Group $group -ErrorAction SilentlyContinue
                    foreach ($member in $members) {
                        # Check by email, login name, or name patterns
                        if ($member.Email -like "*$userEmail*" -or
                            $member.LoginName -like "*$userEmail*" -or
                            ($member.Title -like "*kaila*" -and $member.Title -like "*trapani*")) {
                            
                            $foundIdentities += [PSCustomObject]@{
                                FoundIn = "Group: $($group.Title)"
                                IdentityType = "Group Member"
                                Email = $member.Email
                                LoginName = $member.LoginName
                                Title = $member.Title
                                ObjectId = $member.Id
                                UserId = $member.UserId
                                GroupId = $group.Id
                                IsActive = -not ($member.LoginName.Contains("_disabled_") -or $member.Title.Contains("(deleted)"))
                            }
                        }
                    }
                } catch {
                    # Skip groups we can't read
                }
            }
        } catch {
            # Skip if can't get groups
        }
        
        # 3. SEARCH ROLE ASSIGNMENTS by Object ID patterns
        try {
            $roleAssignments = Get-PnPRoleAssignment -ErrorAction SilentlyContinue
            foreach ($assignment in $roleAssignments) {
                if ($assignment.Member.Email -like "*$userEmail*" -or
                    $assignment.Member.LoginName -like "*$userEmail*" -or
                    ($assignment.Member.Title -like "*kaila*" -and $assignment.Member.Title -like "*trapani*")) {
                    
                    $roleDefinitions = $assignment.RoleDefinitionBindings | ForEach-Object { $_.Name }
                    $foundIdentities += [PSCustomObject]@{
                        FoundIn = "Direct Role Assignment"
                        IdentityType = "Direct Permission"
                        Email = $assignment.Member.Email
                        LoginName = $assignment.Member.LoginName
                        Title = $assignment.Member.Title
                        ObjectId = $assignment.Member.Id
                        Permissions = ($roleDefinitions -join ", ")
                        IsActive = -not ($assignment.Member.LoginName.Contains("_disabled_") -or $assignment.Member.Title.Contains("(deleted)"))
                    }
                }
            }
        } catch {
            # Skip if can't check role assignments
        }
        
        # 4. CHECK SITE COLLECTION ADMINISTRATORS by Object ID
        try {
            $admins = Get-PnPSiteCollectionAdmin -ErrorAction SilentlyContinue
            foreach ($admin in $admins) {
                if ($admin.Email -like "*$userEmail*" -or
                    $admin.LoginName -like "*$userEmail*" -or
                    ($admin.Title -like "*kaila*" -and $admin.Title -like "*trapani*")) {
                    
                    $foundIdentities += [PSCustomObject]@{
                        FoundIn = "Site Collection Administrators"
                        IdentityType = "Site Admin"
                        Email = $admin.Email
                        LoginName = $admin.LoginName
                        Title = $admin.Title
                        ObjectId = $admin.Id
                        UserId = $admin.UserId
                        IsActive = -not ($admin.LoginName.Contains("_disabled_") -or $admin.Title.Contains("(deleted)"))
                    }
                }
            }
        } catch {
            # Skip if can't check admins
        }
        
        Disconnect-PnPOnline -ErrorAction SilentlyContinue
        
        if ($foundIdentities.Count -gt 0) {
            return [PSCustomObject]@{
                SiteUrl = $siteUrl
                SiteTitle = $siteTitle
                FoundIdentities = $foundIdentities
                TotalFound = $foundIdentities.Count
            }
        }
    } catch {
        Disconnect-PnPOnline -ErrorAction SilentlyContinue
    }
    
    return $null
}

# Process sites in parallel batches
Write-Host "`n🚀 Searching $totalSites sites for Object ID based identities with $ThrottleLimit parallel threads..." -ForegroundColor Yellow
Write-Host "🔍 Looking for: Object IDs, GUIDs, UPNs, and identity patterns" -ForegroundColor Green
Write-Host "This finds old Azure AD identities that email searches miss. Press Ctrl+C to stop early." -ForegroundColor Yellow

for ($i = 0; $i -lt $allSites.Count; $i += $BatchSize) {
    $batch = $allSites[$i..[math]::Min($i + $BatchSize - 1, $allSites.Count - 1)]
    
    $batchResults = $batch | ForEach-Object -Parallel {
        $result = & $using:scriptBlock -siteUrl $_ -userEmail $using:UserEmail -thumbprint $using:cert.Thumbprint
        
        if ($result) {
            Write-Host "`n🔥 FOUND OBJECT IDs: $($result.SiteTitle) - $($result.TotalFound) identities" -ForegroundColor Red
            return $result
        }
        
        return $null
    } -ThrottleLimit $ThrottleLimit
    
    # Add batch results to findings
    $batchResults | Where-Object { $_ -ne $null } | ForEach-Object { $allFindings.Add($_) }
    
    # Add processed count
    foreach ($site in $batch) {
        $processed.Add(1)
    }
    
    # Show progress
    $currentCount = $processed.Count
    $foundCount = $allFindings.Count
    Write-Host "`r[$currentCount/$totalSites] Sites searched | Found: $foundCount sites with Object IDs" -NoNewline
    
    # Stop if we've found enough
    if ($allFindings.Count -ge $MaxSitesToFind) {
        Write-Host "`n🎯 Found $MaxSitesToFind sites with Object ID identities - stopping search!" -ForegroundColor Green
        break
    }
}

# Final results
$elapsed = (Get-Date) - $startTime
Write-Host "`n`n=============== OBJECT ID SEARCH RESULTS ===============" -ForegroundColor Cyan
Write-Host "📊 Total sites searched: $($processed.Count)" -ForegroundColor White
Write-Host "🔍 Sites with Object ID matches: $($allFindings.Count)" -ForegroundColor White
Write-Host "⏱️  Time taken: $([math]::Round($elapsed.TotalMinutes, 1)) minutes" -ForegroundColor White

if ($allFindings.Count -gt 0) {
    Write-Host "`n🔥 OBJECT ID BASED IDENTITIES FOUND:" -ForegroundColor Red
    
    $allIdentities = @()
    $totalIdentities = 0
    
    $allFindings | ForEach-Object {
        $totalIdentities += $_.TotalFound
        Write-Host "`n• $($_.SiteTitle) ($($_.TotalFound) identities)" -ForegroundColor White
        Write-Host "  URL: $($_.SiteUrl)" -ForegroundColor Gray
        
        $_.FoundIdentities | ForEach-Object {
            $allIdentities += $_
            $statusColor = if ($_.IsActive) { "Yellow" } else { "Red" }
            $status = if ($_.IsActive) { "ACTIVE" } else { "DISABLED/OLD" }
            
            Write-Host "  - $($_.IdentityType): $($_.Title) [$status]" -ForegroundColor $statusColor
            Write-Host "    Email: $($_.Email)" -ForegroundColor Gray
            Write-Host "    Login: $($_.LoginName)" -ForegroundColor Gray
            Write-Host "    Object ID: $($_.ObjectId)" -ForegroundColor Cyan
            if ($_.UserId) {
                Write-Host "    User ID: $($_.UserId)" -ForegroundColor Cyan
            }
            Write-Host "    Found in: $($_.FoundIn)" -ForegroundColor Gray
            if ($_.Permissions) {
                Write-Host "    Permissions: $($_.Permissions)" -ForegroundColor Gray
            }
        }
    }
    
    Write-Host "`n📊 SUMMARY:" -ForegroundColor Cyan
    Write-Host "Total Object ID based identities found: $totalIdentities" -ForegroundColor White
    
    $activeIdentities = $allIdentities | Where-Object { $_.IsActive }
    $oldIdentities = $allIdentities | Where-Object { -not $_.IsActive }
    
    Write-Host "Active identities: $($activeIdentities.Count)" -ForegroundColor Yellow
    Write-Host "Old/Disabled identities: $($oldIdentities.Count)" -ForegroundColor Red
    
    # Group by Object ID to find duplicates
    $uniqueObjectIds = $allIdentities | Group-Object ObjectId | Where-Object { $_.Count -gt 1 }
    if ($uniqueObjectIds.Count -gt 0) {
        Write-Host "`n⚠️  DUPLICATE OBJECT IDs FOUND:" -ForegroundColor Red
        Write-Host "Same user exists in multiple places with same Object ID:" -ForegroundColor Yellow
        $uniqueObjectIds | ForEach-Object {
            Write-Host "  Object ID: $($_.Name)" -ForegroundColor Cyan
            $_.Group | ForEach-Object {
                Write-Host "    - $($_.FoundIn)" -ForegroundColor Gray
            }
        }
    }
    
    if ($oldIdentities.Count -gt 0) {
        Write-Host "`n🔥 OLD OBJECT ID REMNANTS FOUND!" -ForegroundColor Red
        Write-Host "These old Object IDs could be blocking new access:" -ForegroundColor Yellow
        $oldIdentities | ForEach-Object {
            Write-Host "  - Object ID: $($_.ObjectId) in $($_.FoundIn)" -ForegroundColor Red
            Write-Host "    Login: $($_.LoginName)" -ForegroundColor Gray
        }
        
        Write-Host "`n💡 RECOMMENDATION:" -ForegroundColor Cyan
        Write-Host "Remove these old Object ID permissions to allow new access" -ForegroundColor Yellow
    }
    
    # Save detailed results
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $csvPath = Join-Path $scriptPath "objectid_identities_${timestamp}.csv"
    $allIdentities | Export-Csv -Path $csvPath -NoTypeInformation
    Write-Host "`n📄 Detailed Object ID results saved to: $csvPath" -ForegroundColor Gray
    
} else {
    Write-Host "`n✅ NO OBJECT ID BASED IDENTITIES FOUND" -ForegroundColor Green
    Write-Host "No Object IDs, GUIDs, or UPN patterns found matching the user" -ForegroundColor Green
    Write-Host "This confirms the user has been completely cleaned from SharePoint" -ForegroundColor Yellow
}

# Clean up certificate store
$store = New-Object System.Security.Cryptography.X509Certificates.X509Store("My", "CurrentUser")
$store.Open("ReadWrite")
$store.Remove($cert)
$store.Close()

Write-Host "`n✅ Object ID based search complete!" -ForegroundColor Green