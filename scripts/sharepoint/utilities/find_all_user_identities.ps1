# COMPREHENSIVE USER IDENTITY FINDER
# Searches for ALL possible variations of a user's identity across SharePoint
# Finds old accounts, orphaned permissions, cached profiles that could block new access

param(
    [string]$UserEmail = "kaila.trapani@optimizely.com",
    [string]$FirstName = "Kaila",
    [string]$LastName = "Trapani",
    [int]$ThrottleLimit = 15,
    [int]$BatchSize = 100,
    [int]$MaxSitesToFind = 50  # Find more variations
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

Write-Host "`n=============== COMPREHENSIVE USER IDENTITY FINDER ===============" -ForegroundColor Cyan
Write-Host "🔍 Searching for ALL identity variations that could be blocking access" -ForegroundColor Red
Write-Host "Target User: $UserEmail ($FirstName $LastName)" -ForegroundColor White
Write-Host "This will find old accounts, orphaned permissions, cached profiles" -ForegroundColor Yellow

# Generate all possible identity variations to search for
$searchPatterns = @()
$userPart = $UserEmail.Split('@')[0]
$domain = $UserEmail.Split('@')[1]

# Email variations
$searchPatterns += $UserEmail
$searchPatterns += $userPart  # just the username part
$searchPatterns += "$FirstName.$LastName"
$searchPatterns += "$($FirstName.ToLower()).$($LastName.ToLower())"
$searchPatterns += "$($FirstName.Substring(0,1)).$LastName"
$searchPatterns += "$($FirstName.Substring(0,1).ToLower()).$($LastName.ToLower())"

# Name variations
$searchPatterns += "$FirstName $LastName"
$searchPatterns += "$LastName, $FirstName"
$searchPatterns += $FirstName
$searchPatterns += $LastName

# Login name pattern variations (common SharePoint formats)
$searchPatterns += "i:0#.f|membership|$UserEmail"
$searchPatterns += "i:0#.w|$domain\$userPart"
$searchPatterns += "c:0(.s|true"  # Claims patterns
$searchPatterns += "i:0).f|membership|$UserEmail"

Write-Host "`n🔍 Generated $($searchPatterns.Count) identity patterns to search for:" -ForegroundColor Yellow
$searchPatterns | ForEach-Object { Write-Host "  - $_" -ForegroundColor Gray }

# Connect to admin center
Write-Host "`nConnecting to SharePoint Admin Center..." -ForegroundColor Yellow
try {
    Connect-PnPOnline -Url "https://episerver99-admin.sharepoint.com" `
                      -ClientId "fe2a9efe-3000-4b02-96ea-344a2583dd52" `
                      -Tenant "3ec00d79-021a-42d4-aac8-dcb35973dff2" `
                      -Thumbprint $cert.Thumbprint `
                      -ErrorAction Stop
    
    Write-Host "✓ Connected to SharePoint Admin Center" -ForegroundColor Green
    
    # Get ALL sites (SharePoint + OneDrive)
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

# COMPREHENSIVE identity search function
$scriptBlock = {
    param($siteUrl, $searchPatterns, $thumbprint)
    
    try {
        Connect-PnPOnline -Url $siteUrl `
                          -ClientId "fe2a9efe-3000-4b02-96ea-344a2583dd52" `
                          -Tenant "3ec00d79-021a-42d4-aac8-dcb35973dff2" `
                          -Thumbprint $thumbprint `
                          -ErrorAction Stop
        
        $web = Get-PnPWeb
        $siteTitle = if ($web.Title) { $web.Title } else { Split-Path $siteUrl -Leaf }
        $foundIdentities = @()
        
        # 1. SEARCH ALL SHAREPOINT GROUPS for any identity variations
        try {
            $allGroups = Get-PnPGroup -ErrorAction SilentlyContinue
            foreach ($group in $allGroups) {
                try {
                    $members = Get-PnPGroupMember -Group $group -ErrorAction SilentlyContinue
                    foreach ($member in $members) {
                        foreach ($pattern in $searchPatterns) {
                            if ($member.Email -like "*$pattern*" -or 
                                $member.LoginName -like "*$pattern*" -or 
                                $member.Title -like "*$pattern*") {
                                
                                $foundIdentities += [PSCustomObject]@{
                                    FoundIn = "Group: $($group.Title)"
                                    IdentityType = "Group Member"
                                    Email = $member.Email
                                    LoginName = $member.LoginName
                                    Title = $member.Title
                                    MatchedPattern = $pattern
                                    IsActive = -not $member.LoginName.Contains("_disabled_")
                                }
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
        
        # 2. SEARCH SITE COLLECTION ADMINISTRATORS
        try {
            $admins = Get-PnPSiteCollectionAdmin -ErrorAction SilentlyContinue
            foreach ($admin in $admins) {
                foreach ($pattern in $searchPatterns) {
                    if ($admin.Email -like "*$pattern*" -or 
                        $admin.LoginName -like "*$pattern*" -or 
                        $admin.Title -like "*$pattern*") {
                        
                        $foundIdentities += [PSCustomObject]@{
                            FoundIn = "Site Collection Administrators"
                            IdentityType = "Site Admin"
                            Email = $admin.Email
                            LoginName = $admin.LoginName
                            Title = $admin.Title
                            MatchedPattern = $pattern
                            IsActive = -not $admin.LoginName.Contains("_disabled_")
                        }
                    }
                }
            }
        } catch {
            # Skip if can't check admins
        }
        
        # 3. SEARCH ALL ROLE ASSIGNMENTS (direct permissions)
        try {
            $roleAssignments = Get-PnPRoleAssignment -ErrorAction SilentlyContinue
            foreach ($assignment in $roleAssignments) {
                foreach ($pattern in $searchPatterns) {
                    if ($assignment.Member.Email -like "*$pattern*" -or 
                        $assignment.Member.LoginName -like "*$pattern*" -or 
                        $assignment.Member.Title -like "*$pattern*") {
                        
                        $roleDefinitions = $assignment.RoleDefinitionBindings | ForEach-Object { $_.Name }
                        $foundIdentities += [PSCustomObject]@{
                            FoundIn = "Direct Role Assignment"
                            IdentityType = "Direct Permission"
                            Email = $assignment.Member.Email
                            LoginName = $assignment.Member.LoginName
                            Title = $assignment.Member.Title
                            MatchedPattern = $pattern
                            IsActive = -not $assignment.Member.LoginName.Contains("_disabled_")
                            Permissions = ($roleDefinitions -join ", ")
                        }
                    }
                }
            }
        } catch {
            # Skip if can't check role assignments
        }
        
        # 4. SEARCH USER INFORMATION LIST (cached user profiles)
        try {
            $userInfoList = Get-PnPList -Identity "User Information List" -ErrorAction SilentlyContinue
            if ($userInfoList) {
                $userItems = Get-PnPListItem -List $userInfoList -ErrorAction SilentlyContinue
                foreach ($userItem in $userItems) {
                    foreach ($pattern in $searchPatterns) {
                        $name = $userItem["Name"]
                        $email = $userItem["EMail"]
                        $loginName = $userItem["UserName"]
                        
                        if ($name -like "*$pattern*" -or 
                            $email -like "*$pattern*" -or 
                            $loginName -like "*$pattern*") {
                            
                            $foundIdentities += [PSCustomObject]@{
                                FoundIn = "User Information List (Cached)"
                                IdentityType = "Cached User Profile"
                                Email = $email
                                LoginName = $loginName
                                Title = $name
                                MatchedPattern = $pattern
                                IsActive = -not $loginName.Contains("_disabled_")
                            }
                        }
                    }
                }
            }
        } catch {
            # Skip if can't check user info list
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
Write-Host "`n🚀 Searching $totalSites sites for identity variations with $ThrottleLimit parallel threads..." -ForegroundColor Yellow
Write-Host "🔍 Looking for: Old accounts, orphaned permissions, cached profiles" -ForegroundColor Green
Write-Host "This will find ANY identity variation that exists. Press Ctrl+C to stop early." -ForegroundColor Yellow

for ($i = 0; $i -lt $allSites.Count; $i += $BatchSize) {
    $batch = $allSites[$i..[math]::Min($i + $BatchSize - 1, $allSites.Count - 1)]
    
    $batchResults = $batch | ForEach-Object -Parallel {
        $result = & $using:scriptBlock -siteUrl $_ -searchPatterns $using:searchPatterns -thumbprint $using:cert.Thumbprint
        
        if ($result) {
            Write-Host "`n🔥 FOUND IDENTITIES: $($result.SiteTitle) - $($result.TotalFound) variations" -ForegroundColor Red
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
    Write-Host "`r[$currentCount/$totalSites] Sites searched | Found: $foundCount sites with identities" -NoNewline
    
    # Stop if we've found enough
    if ($allFindings.Count -ge $MaxSitesToFind) {
        Write-Host "`n🎯 Found $MaxSitesToFind sites with identity variations - stopping search!" -ForegroundColor Green
        break
    }
}

# Final results
$elapsed = (Get-Date) - $startTime
Write-Host "`n`n=============== IDENTITY SEARCH RESULTS ===============" -ForegroundColor Cyan
Write-Host "📊 Total sites searched: $($processed.Count)" -ForegroundColor White
Write-Host "🔍 Sites with identity variations: $($allFindings.Count)" -ForegroundColor White
Write-Host "⏱️  Time taken: $([math]::Round($elapsed.TotalMinutes, 1)) minutes" -ForegroundColor White

if ($allFindings.Count -gt 0) {
    Write-Host "`n🔥 IDENTITY VARIATIONS FOUND:" -ForegroundColor Red
    
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
            Write-Host "    Found in: $($_.FoundIn)" -ForegroundColor Gray
            if ($_.Permissions) {
                Write-Host "    Permissions: $($_.Permissions)" -ForegroundColor Gray
            }
        }
    }
    
    Write-Host "`n📊 SUMMARY:" -ForegroundColor Cyan
    Write-Host "Total identity variations found: $totalIdentities" -ForegroundColor White
    
    $activeIdentities = $allIdentities | Where-Object { $_.IsActive }
    $oldIdentities = $allIdentities | Where-Object { -not $_.IsActive }
    
    Write-Host "Active identities: $($activeIdentities.Count)" -ForegroundColor Yellow
    Write-Host "Old/Disabled identities: $($oldIdentities.Count)" -ForegroundColor Red
    
    if ($oldIdentities.Count -gt 0) {
        Write-Host "`n⚠️  OLD ACCOUNT REMNANTS FOUND!" -ForegroundColor Red
        Write-Host "These could be blocking new access:" -ForegroundColor Yellow
        $oldIdentities | ForEach-Object {
            Write-Host "  - $($_.LoginName) in $($_.FoundIn)" -ForegroundColor Red
        }
    }
    
    # Save detailed results
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $csvPath = Join-Path $scriptPath "identity_variations_${timestamp}.csv"
    $allIdentities | Export-Csv -Path $csvPath -NoTypeInformation
    Write-Host "`n📄 Detailed results saved to: $csvPath" -ForegroundColor Gray
    
} else {
    Write-Host "`n✅ NO IDENTITY VARIATIONS FOUND" -ForegroundColor Green
    Write-Host "No old accounts, orphaned permissions, or cached profiles found" -ForegroundColor Green
    Write-Host "This suggests the blocking issue is not from old account remnants" -ForegroundColor Yellow
}

# Clean up certificate store
$store = New-Object System.Security.Cryptography.X509Certificates.X509Store("My", "CurrentUser")
$store.Open("ReadWrite")
$store.Remove($cert)
$store.Close()

Write-Host "`n✅ Comprehensive identity search complete!" -ForegroundColor Green