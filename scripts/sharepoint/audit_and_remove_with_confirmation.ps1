# Script to audit and remove user permissions with confirmation
# Checks both direct permissions and group memberships

# Disconnect existing session
try {
    Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
} catch {
    # Ignore disconnection errors
}

# Load cert and connect to Microsoft Graph
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

# Connect to Microsoft Graph
try {
    Connect-MgGraph -TenantId "3ec00d79-021a-42d4-aac8-dcb35973dff2" `
                    -ClientId "fe2a9efe-3000-4b02-96ea-344a2583dd52" `
                    -Certificate $cert `
                    -NoWelcome
} catch {
    Write-Host "Failed to connect to Microsoft Graph: $_" -ForegroundColor Red
    Write-Host "Error details: $($_.Exception.ToString())" -ForegroundColor Red
    exit 1
}

# Verify connection
$context = Get-MgContext
if (-not $context -or ($context.AuthType -ne "AppOnly" -and -not $context.Account)) {
    Write-Host "Authentication failed. Please verify your certificate and app registration." -ForegroundColor Red
    return
}

Write-Host "Connected via: $($context.AuthType)" -ForegroundColor Green
Write-Host "App ID: $($context.ClientId)" -ForegroundColor Green

$upn = 'kaila.trapani@optimizely.com'
$pageSize = 100
$sitesChecked = 0

# Initialize collections for findings
$directPermissions = @()
$groupMemberships = @()
$errors = @()

# Progress tracking
$startTime = Get-Date

Write-Host "`nPhase 1: Scanning all sites for permissions..." -ForegroundColor Cyan
Write-Host "Looking for both direct permissions and group memberships for: $upn" -ForegroundColor Cyan
Write-Host "This may take some time for 7000+ sites..." -ForegroundColor Yellow
Write-Host "" -NoNewline

# First, try to get user information
$userId = $null
try {
    $user = Get-MgUser -Filter "userPrincipalName eq '$upn'" -ErrorAction Stop
    $userId = $user.Id
    Write-Host "Found user: $($user.DisplayName) (ID: $userId)" -ForegroundColor Green
} catch {
    Write-Host "Warning: Could not look up user ID. Will check by UPN only." -ForegroundColor Yellow
}

# Get all sites using pagination
$nextLink = $null
$pageCount = 0

do {
    $pageCount++
    
    if ($nextLink) {
        $response = Invoke-MgGraphRequest -Uri $nextLink -Method GET
    } else {
        $response = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/v1.0/sites?`$top=$pageSize" -Method GET
    }
    
    $sites = $response.value
    $nextLink = $response.'@odata.nextLink'
    
    if (-not $sites -or $sites.Count -eq 0) {
        break
    }

    foreach ($site in $sites) {
        $sitesChecked++
        
        # Show progress every 50 sites
        if ($sitesChecked % 50 -eq 0) {
            $elapsed = (Get-Date) - $startTime
            $rate = $sitesChecked / $elapsed.TotalMinutes
            Write-Host "`rChecked: $sitesChecked | Direct perms: $($directPermissions.Count) | Via groups: $($groupMemberships.Count) | Rate: $([math]::Round($rate, 1)) sites/min" -NoNewline
        }

        try {
            # Get all permissions for the site
            $allPermissions = Get-MgSitePermission -SiteId $site.id -All -ErrorAction Stop
            
            foreach ($perm in $allPermissions) {
                # Check for direct user permissions
                $directUserFound = $false
                if ($perm.GrantedToIdentities) {
                    foreach ($identity in $perm.GrantedToIdentities) {
                        if ($identity.User -and $identity.User.UserPrincipalName -eq $upn) {
                            $directUserFound = $true
                            break
                        }
                    }
                }
                if ($perm.GrantedToIdentitiesV2 -and -not $directUserFound) {
                    foreach ($identity in $perm.GrantedToIdentitiesV2) {
                        if ($identity.User -and $identity.User.UserPrincipalName -eq $upn) {
                            $directUserFound = $true
                            break
                        }
                    }
                }
                
                if ($directUserFound) {
                    $directPermissions += [PSCustomObject]@{
                        SiteId = $site.id
                        SiteName = $site.displayName
                        SiteUrl = $site.webUrl
                        PermissionId = $perm.Id
                        Roles = $perm.Roles -join ", "
                        Type = "Direct"
                    }
                }
                
                # Check for group permissions
                if ($perm.GrantedToIdentities) {
                    foreach ($identity in $perm.GrantedToIdentities) {
                        if ($identity.Group) {
                            # Store group info for later checking
                            $groupMemberships += [PSCustomObject]@{
                                SiteId = $site.id
                                SiteName = $site.displayName
                                SiteUrl = $site.webUrl
                                PermissionId = $perm.Id
                                GroupId = $identity.Group.Id
                                GroupName = $identity.Group.DisplayName
                                Roles = $perm.Roles -join ", "
                                Type = "Group"
                                UserIsMember = "ToBeChecked"  # We'll check this later
                            }
                        }
                    }
                }
                if ($perm.GrantedToIdentitiesV2) {
                    foreach ($identity in $perm.GrantedToIdentitiesV2) {
                        if ($identity.Group) {
                            $groupMemberships += [PSCustomObject]@{
                                SiteId = $site.id
                                SiteName = $site.displayName
                                SiteUrl = $site.webUrl
                                PermissionId = $perm.Id
                                GroupId = $identity.Group.Id
                                GroupName = $identity.Group.DisplayName
                                Roles = $perm.Roles -join ", "
                                Type = "Group"
                                UserIsMember = "ToBeChecked"
                            }
                        }
                    }
                }
            }
        } catch {
            $errors += [PSCustomObject]@{
                SiteName = $site.displayName
                SiteUrl = $site.webUrl
                Error = $_.Exception.Message
            }
        }
    }
} while ($nextLink)

Write-Host "`n`nPhase 1 Complete!" -ForegroundColor Green
Write-Host "Sites checked: $sitesChecked" -ForegroundColor Cyan
Write-Host "Direct permissions found: $($directPermissions.Count)" -ForegroundColor Cyan
Write-Host "Potential group permissions: $($groupMemberships.Count)" -ForegroundColor Cyan
Write-Host "Errors encountered: $($errors.Count)" -ForegroundColor Cyan

# Phase 2: Check group memberships if we have a user ID
$confirmedGroupPerms = @()
if ($userId -and $groupMemberships.Count -gt 0) {
    Write-Host "`nPhase 2: Checking group memberships..." -ForegroundColor Cyan
    
    # Get unique groups
    $uniqueGroups = $groupMemberships | Select-Object -Property GroupId, GroupName -Unique
    Write-Host "Checking $($uniqueGroups.Count) unique groups..." -ForegroundColor Yellow
    
    foreach ($group in $uniqueGroups) {
        try {
            # Check if user is member of this group
            $isMember = $false
            try {
                $members = Get-MgGroupMember -GroupId $group.GroupId -All -ErrorAction Stop
                $isMember = $members | Where-Object { $_.Id -eq $userId }
            } catch {
                # Try transitive membership
                try {
                    $response = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/v1.0/users/$userId/memberOf" -Method GET
                    $isMember = $response.value | Where-Object { $_.id -eq $group.GroupId }
                } catch {
                    # Skip if we can't check
                }
            }
            
            if ($isMember) {
                Write-Host "✓ User is member of: $($group.GroupName)" -ForegroundColor Green
                # Update all permissions for this group
                $groupMemberships | Where-Object { $_.GroupId -eq $group.GroupId } | ForEach-Object {
                    $_.UserIsMember = "Confirmed"
                    $confirmedGroupPerms += $_
                }
            }
        } catch {
            # Skip groups we can't check
        }
    }
} else {
    Write-Host "`nSkipping group membership check (no user ID available)" -ForegroundColor Yellow
}

# Generate summary report
Write-Host "`n" -NoNewline
Write-Host "=============== FINDINGS SUMMARY ===============" -ForegroundColor Cyan

if ($directPermissions.Count -gt 0) {
    Write-Host "`nDIRECT PERMISSIONS ($($directPermissions.Count) found):" -ForegroundColor Yellow
    $directPermissions | Group-Object -Property SiteName | ForEach-Object {
        Write-Host "  • $($_.Name)" -ForegroundColor White
        Write-Host "    URL: $(($_.Group | Select-Object -First 1).SiteUrl)" -ForegroundColor Gray
        Write-Host "    Roles: $(($_.Group | Select-Object -First 1).Roles)" -ForegroundColor Gray
    }
}

if ($confirmedGroupPerms.Count -gt 0) {
    Write-Host "`nGROUP-BASED PERMISSIONS ($($confirmedGroupPerms.Count) found):" -ForegroundColor Yellow
    $confirmedGroupPerms | Group-Object -Property GroupName | ForEach-Object {
        $groupName = $_.Name
        $sitesCount = ($_.Group | Select-Object -Property SiteName -Unique).Count
        Write-Host "  • Group: $groupName (affects $sitesCount sites)" -ForegroundColor White
        $_.Group | Select-Object -Property SiteName, SiteUrl -Unique | ForEach-Object {
            Write-Host "    - $($_.SiteName)" -ForegroundColor Gray
        }
    }
}

if ($directPermissions.Count -eq 0 -and $confirmedGroupPerms.Count -eq 0) {
    Write-Host "`nNo permissions found for $upn" -ForegroundColor Yellow
}

Write-Host "===============================================" -ForegroundColor Cyan

# Save detailed findings to CSV
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$csvPath = Join-Path $scriptPath "permissions_audit_${timestamp}.csv"

$allFindings = @()
$allFindings += $directPermissions | Select-Object *, @{Name="Status";Expression={"Found"}}
$allFindings += $confirmedGroupPerms | Select-Object *, @{Name="Status";Expression={"Found"}}

if ($allFindings.Count -gt 0) {
    $allFindings | Export-Csv -Path $csvPath -NoTypeInformation
    Write-Host "`nDetailed findings saved to: $csvPath" -ForegroundColor Green
}

# Prompt for action
if ($directPermissions.Count -gt 0 -or $confirmedGroupPerms.Count -gt 0) {
    Write-Host "`n" -NoNewline
    Write-Host "=============== ACTION REQUIRED ===============" -ForegroundColor Yellow
    Write-Host "Found $($directPermissions.Count) direct permissions and $($confirmedGroupPerms.Count) group-based permissions" -ForegroundColor Cyan
    
    $choice = Read-Host "`nDo you want to remove these permissions? (Y/N)"
    
    if ($choice -eq 'Y' -or $choice -eq 'y') {
        Write-Host "`nSelect removal option:" -ForegroundColor Cyan
        Write-Host "1. Remove ALL permissions (direct and group-based)" -ForegroundColor White
        Write-Host "2. Remove only DIRECT permissions" -ForegroundColor White
        Write-Host "3. Remove only GROUP memberships" -ForegroundColor White
        Write-Host "4. Cancel - don't remove anything" -ForegroundColor White
        
        $removalChoice = Read-Host "`nEnter your choice (1-4)"
        
        $removedCount = 0
        
        switch ($removalChoice) {
            '1' {
                # Remove all
                Write-Host "`nRemoving all permissions..." -ForegroundColor Yellow
                
                # Remove direct permissions
                foreach ($perm in $directPermissions) {
                    try {
                        Remove-MgSitePermission -SiteId $perm.SiteId -PermissionId $perm.PermissionId -ErrorAction Stop
                        Write-Host "✓ Removed direct permission from: $($perm.SiteName)" -ForegroundColor Green
                        $removedCount++
                    } catch {
                        Write-Host "✗ Failed to remove from $($perm.SiteName): $_" -ForegroundColor Red
                    }
                }
                
                # Remove group memberships
                if ($userId) {
                    $uniqueGroups = $confirmedGroupPerms | Select-Object -Property GroupId, GroupName -Unique
                    foreach ($group in $uniqueGroups) {
                        try {
                            Remove-MgGroupMemberByRef -GroupId $group.GroupId -DirectoryObjectId $userId -ErrorAction Stop
                            Write-Host "✓ Removed from group: $($group.GroupName)" -ForegroundColor Green
                            $removedCount++
                        } catch {
                            Write-Host "✗ Failed to remove from group $($group.GroupName): $_" -ForegroundColor Red
                        }
                    }
                }
            }
            '2' {
                # Remove only direct
                Write-Host "`nRemoving direct permissions only..." -ForegroundColor Yellow
                foreach ($perm in $directPermissions) {
                    try {
                        Remove-MgSitePermission -SiteId $perm.SiteId -PermissionId $perm.PermissionId -ErrorAction Stop
                        Write-Host "✓ Removed direct permission from: $($perm.SiteName)" -ForegroundColor Green
                        $removedCount++
                    } catch {
                        Write-Host "✗ Failed to remove from $($perm.SiteName): $_" -ForegroundColor Red
                    }
                }
            }
            '3' {
                # Remove only groups
                if ($userId) {
                    Write-Host "`nRemoving group memberships only..." -ForegroundColor Yellow
                    $uniqueGroups = $confirmedGroupPerms | Select-Object -Property GroupId, GroupName -Unique
                    foreach ($group in $uniqueGroups) {
                        try {
                            Remove-MgGroupMemberByRef -GroupId $group.GroupId -DirectoryObjectId $userId -ErrorAction Stop
                            Write-Host "✓ Removed from group: $($group.GroupName)" -ForegroundColor Green
                            $removedCount++
                        } catch {
                            Write-Host "✗ Failed to remove from group $($group.GroupName): $_" -ForegroundColor Red
                        }
                    }
                } else {
                    Write-Host "Cannot remove group memberships without user ID" -ForegroundColor Red
                }
            }
            default {
                Write-Host "`nCancelled - no permissions removed" -ForegroundColor Yellow
            }
        }
        
        if ($removedCount -gt 0) {
            Write-Host "`n✓ Successfully removed $removedCount permission(s)" -ForegroundColor Green
            
            # Update CSV with removal status
            $allFindings | ForEach-Object { $_.Status = "Removed" }
            $allFindings | Export-Csv -Path $csvPath -NoTypeInformation -Force
            Write-Host "Updated audit log: $csvPath" -ForegroundColor Green
        }
    } else {
        Write-Host "`nNo action taken - permissions remain unchanged" -ForegroundColor Yellow
    }
}

# Final summary
$elapsed = (Get-Date) - $startTime
Write-Host "`n" -NoNewline
Write-Host "=============== FINAL SUMMARY ===============" -ForegroundColor Cyan
Write-Host "Total execution time: $([math]::Round($elapsed.TotalMinutes, 1)) minutes" -ForegroundColor Cyan
Write-Host "Sites scanned: $sitesChecked" -ForegroundColor Cyan
Write-Host "Average rate: $([math]::Round($sitesChecked / $elapsed.TotalMinutes, 1)) sites/minute" -ForegroundColor Cyan
Write-Host "=============================================" -ForegroundColor Cyan

# Disconnect
Disconnect-MgGraph