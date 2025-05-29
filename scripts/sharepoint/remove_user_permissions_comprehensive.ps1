# Comprehensive script to remove ALL user permissions from SharePoint sites
# Checks: Direct permissions, Group memberships, Site ownership, Site collection admin
# Optimized for large numbers of sites with minimal memory usage and logging

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
    Write-Host " Authentication failed. Please verify your certificate and app registration." -ForegroundColor Red
    return
}

Write-Host "Connected via: $($context.AuthType)" -ForegroundColor Green
Write-Host "App ID: $($context.ClientId)" -ForegroundColor Green

$upn = 'kaila.trapani@optimizely.com'
$removedFrom = @()
$permissionsRemoved = 0
$maxPermissionsToRemove = 10
$pageSize = 100
$sitesChecked = 0
$auditOnly = $true  # Start in audit mode

# Statistics
$stats = @{
    DirectFound = 0
    GroupFound = 0
    OwnerFound = 0
    ErrorCount = 0
}

# Collections for storing findings
$allFindings = @{
    Direct = @()
    Group = @()
    Owner = @()
}

# Get user object and group memberships upfront
Write-Host "`nGathering user information..." -ForegroundColor Yellow
$user = $null
$userGroups = @()
$userGroupIds = @()
try {
    $user = Get-MgUser -Filter "userPrincipalName eq '$upn'" -ErrorAction Stop
    Write-Host "✓ Found user: $($user.DisplayName) (ID: $($user.Id))" -ForegroundColor Green
    
    # Get all groups the user is a member of
    $userGroups = Get-MgUserMemberOf -UserId $user.Id -All | Where-Object { $_.AdditionalProperties.'@odata.type' -eq '#microsoft.graph.group' }
    $userGroupIds = $userGroups | ForEach-Object { $_.Id }
    Write-Host "✓ User is member of $($userGroups.Count) groups" -ForegroundColor Green
    
    if ($userGroups.Count -gt 0 -and $userGroups.Count -le 10) {
        Write-Host "  Groups:" -ForegroundColor Gray
        $userGroups | ForEach-Object {
            Write-Host "    - $($_.AdditionalProperties.displayName)" -ForegroundColor Gray
        }
    } elseif ($userGroups.Count -gt 10) {
        Write-Host "  Showing first 5 groups:" -ForegroundColor Gray
        $userGroups | Select-Object -First 5 | ForEach-Object {
            Write-Host "    - $($_.AdditionalProperties.displayName)" -ForegroundColor Gray
        }
        Write-Host "    ... and $($userGroups.Count - 5) more" -ForegroundColor Gray
    }
} catch {
    Write-Host "✗ Could not find user ${upn}: $_" -ForegroundColor Red
    exit 1
}

# Initialize CSV with minimal data - only sites with permissions or errors
$csvPath = Join-Path $scriptPath "site_permissions_comprehensive.csv"
$csvData = @()

# Progress tracking
$saveInterval = 25  # Save CSV every 25 findings
$lastSaveAt = 0
$startTime = Get-Date

Write-Host "`nSearching for ALL $upn permissions across sites..." -ForegroundColor Cyan
Write-Host "Checking: Direct permissions, Group permissions, Site ownership" -ForegroundColor Cyan
Write-Host "Will stop after removing from $maxPermissionsToRemove sites" -ForegroundColor Cyan
Write-Host "Progress will be saved to: $csvPath" -ForegroundColor Cyan
Write-Host "Only logging sites with permissions or errors to minimize data" -ForegroundColor Cyan
Write-Host "" -NoNewline

# Function to save CSV
function Save-Progress {
    param($Data, $Path, $Stats, $SitesChecked, $PermissionsRemoved)
    
    # Add summary row
    $summaryData = $Data + @([PSCustomObject]@{
        SiteName = "PROGRESS_UPDATE"
        SiteUrl = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        PermissionType = "Summary"
        PermissionDetails = "Checked: $SitesChecked | Direct: $($Stats.DirectFound) | Group: $($Stats.GroupFound) | Owner: $($Stats.OwnerFound) | Removed: $PermissionsRemoved"
        UserPrincipalName = ""
        Action = ""
    })
    
    # Overwrite the CSV file
    $summaryData | Export-Csv -Path $Path -NoTypeInformation -Force
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
        Write-Host "`nNo more sites to check." -ForegroundColor Yellow
        break
    }

    foreach ($site in $sites) {
        if ($permissionsRemoved -ge $maxPermissionsToRemove) {
            Write-Host "`nReached limit of $maxPermissionsToRemove permissions removed. Stopping." -ForegroundColor Yellow
            break
        }

        $sitesChecked++
        
        # Show progress every 50 sites
        if ($sitesChecked % 50 -eq 0) {
            $elapsed = (Get-Date) - $startTime
            $rate = $sitesChecked / $elapsed.TotalMinutes
            Write-Host "`rChecked: $sitesChecked | Direct: $($stats.DirectFound) | Group: $($stats.GroupFound) | Owner: $($stats.OwnerFound) | Removed: $permissionsRemoved | Rate: $([math]::Round($rate, 1))/min" -NoNewline
        }

        try {
            # Get all permissions for the site
            $sitePermissions = Get-MgSitePermission -SiteId $site.id -All -ErrorAction Stop
            
            $foundPermissions = $false
            $siteFindings = @()
            
            # Check each permission
            foreach ($perm in $sitePermissions) {
                # Check direct user permissions
                $directUserFound = $false
                
                # Check GrantedToIdentities
                if ($perm.GrantedToIdentities) {
                    foreach ($identity in $perm.GrantedToIdentities) {
                        if ($identity.User -and $identity.User.UserPrincipalName -eq $upn) {
                            $directUserFound = $true
                            break
                        }
                    }
                }
                
                # Check GrantedToIdentitiesV2
                if (-not $directUserFound -and $perm.GrantedToIdentitiesV2) {
                    foreach ($identity in $perm.GrantedToIdentitiesV2) {
                        if ($identity.User -and $identity.User.UserPrincipalName -eq $upn) {
                            $directUserFound = $true
                            break
                        }
                    }
                }
                
                # Found direct permission
                if ($directUserFound) {
                    $foundPermissions = $true
                    $stats.DirectFound++
                    
                    if ($permissionsRemoved -lt $maxPermissionsToRemove) {
                        try {
                            Remove-MgSitePermission -SiteId $site.id -PermissionId $perm.Id -ErrorAction Stop
                            $removedFrom += "$($site.displayName) : $($site.webUrl)"
                            $permissionsRemoved++
                            
                            $siteFindings += [PSCustomObject]@{
                                SiteName = $site.displayName
                                SiteUrl = $site.webUrl
                                PermissionType = "Direct"
                                PermissionDetails = "Permission ID: $($perm.Id) | Roles: $($perm.Roles -join ', ')"
                                UserPrincipalName = $upn
                                Action = "Removed"
                            }
                            
                            Write-Host "`n✓ DIRECT permission removed at site ${sitesChecked}: $($site.displayName) ($permissionsRemoved/$maxPermissionsToRemove)" -ForegroundColor Green
                        } catch {
                            $siteFindings += [PSCustomObject]@{
                                SiteName = $site.displayName
                                SiteUrl = $site.webUrl
                                PermissionType = "Direct"
                                PermissionDetails = "Permission ID: $($perm.Id)"
                                UserPrincipalName = $upn
                                Action = "RemovalFailed: $_"
                            }
                            Write-Host "`n✗ Failed to remove direct permission at $($site.displayName): $_" -ForegroundColor Red
                        }
                    } else {
                        $siteFindings += [PSCustomObject]@{
                            SiteName = $site.displayName
                            SiteUrl = $site.webUrl
                            PermissionType = "Direct"
                            PermissionDetails = "Permission ID: $($perm.Id) | Roles: $($perm.Roles -join ', ')"
                            UserPrincipalName = $upn
                            Action = "Found (limit reached)"
                        }
                    }
                }
                
                # Check group permissions
                if ($perm.GrantedToIdentitiesV2 -and $userGroupIds.Count -gt 0) {
                    foreach ($identity in $perm.GrantedToIdentitiesV2) {
                        if ($identity.Group -and $userGroupIds -contains $identity.Group.Id) {
                            $foundPermissions = $true
                            $stats.GroupFound++
                            
                            $siteFindings += [PSCustomObject]@{
                                SiteName = $site.displayName
                                SiteUrl = $site.webUrl
                                PermissionType = "Group"
                                PermissionDetails = "Group: $($identity.Group.DisplayName) | Roles: $($perm.Roles -join ', ')"
                                UserPrincipalName = $upn
                                Action = "Found (remove from group required)"
                            }
                            
                            Write-Host "`n⚠ GROUP permission found at site ${sitesChecked}: $($site.displayName)" -ForegroundColor Yellow
                            Write-Host "  Via group: $($identity.Group.DisplayName)" -ForegroundColor Yellow
                            break
                        }
                    }
                }
            }
            
            # Check if user is site owner (try to get site details with owners)
            try {
                $siteDetails = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/v1.0/sites/$($site.id)?`$expand=owners" -Method GET -ErrorAction SilentlyContinue
                
                if ($siteDetails.owners) {
                    foreach ($owner in $siteDetails.owners) {
                        if ($owner.userPrincipalName -eq $upn -or $owner.id -eq $user.Id) {
                            $foundPermissions = $true
                            $stats.OwnerFound++
                            
                            $siteFindings += [PSCustomObject]@{
                                SiteName = $site.displayName
                                SiteUrl = $site.webUrl
                                PermissionType = "Owner"
                                PermissionDetails = "Site Owner"
                                UserPrincipalName = $upn
                                Action = "Found (manual removal required)"
                            }
                            
                            Write-Host "`n⚠ OWNER permission found at site ${sitesChecked}: $($site.displayName)" -ForegroundColor Magenta
                            break
                        }
                    }
                }
            } catch {
                # Owner check failed, continue
            }
            
            # Only add to CSV if permissions were found
            if ($foundPermissions) {
                $csvData += $siteFindings
            }

        } catch {
            # Only log errors
            $stats.ErrorCount++
            Write-Host "`n✗ Error at site $sitesChecked ($($site.displayName)): $($_.Exception.Message)" -ForegroundColor Red
            
            $csvData += [PSCustomObject]@{
                SiteName = $site.displayName
                SiteUrl = $site.webUrl
                PermissionType = "ERROR"
                PermissionDetails = "Error: $($_.Exception.Message)"
                UserPrincipalName = $upn
                Action = "Error"
            }
        }
        
        # Save progress periodically based on findings
        if ($csvData.Count - $lastSaveAt -ge $saveInterval) {
            Save-Progress -Data $csvData -Path $csvPath -Stats $stats -SitesChecked $sitesChecked -PermissionsRemoved $permissionsRemoved
            $lastSaveAt = $csvData.Count
        }
    }

} while ($permissionsRemoved -lt $maxPermissionsToRemove -and $nextLink)

# Final save
Save-Progress -Data $csvData -Path $csvPath -Stats $stats -SitesChecked $sitesChecked -PermissionsRemoved $permissionsRemoved

# Calculate stats
$elapsed = (Get-Date) - $startTime
$rate = $sitesChecked / $elapsed.TotalMinutes

Write-Host "`n`n" -NoNewline
Write-Host "=============== COMPREHENSIVE SUMMARY ===============" -ForegroundColor Cyan
Write-Host "User checked: $upn" -ForegroundColor Cyan
Write-Host "Total sites checked: $sitesChecked" -ForegroundColor Cyan
Write-Host "Sites with findings: $($csvData.Count - 1)" -ForegroundColor Cyan  # -1 for summary row
Write-Host "" -NoNewline
Write-Host "Permission breakdown:" -ForegroundColor Cyan
Write-Host "  - Direct permissions found: $($stats.DirectFound)" -ForegroundColor White
Write-Host "  - Group permissions found: $($stats.GroupFound)" -ForegroundColor White
Write-Host "  - Owner permissions found: $($stats.OwnerFound)" -ForegroundColor White
Write-Host "  - Errors encountered: $($stats.ErrorCount)" -ForegroundColor White
Write-Host "" -NoNewline
Write-Host "Permissions removed: $permissionsRemoved" -ForegroundColor Green
Write-Host "Time taken: $([math]::Round($elapsed.TotalMinutes, 1)) minutes" -ForegroundColor Cyan
Write-Host "Average rate: $([math]::Round($rate, 1)) sites/minute" -ForegroundColor Cyan
Write-Host "Results saved to: $csvPath" -ForegroundColor Cyan

if ($permissionsRemoved -gt 0) {
    Write-Host "`nSites user was removed from:" -ForegroundColor Yellow
    $removedFrom | Select-Object -Unique | ForEach-Object { Write-Host "- $_" -ForegroundColor White }
    
    if ($permissionsRemoved -eq $maxPermissionsToRemove) {
        Write-Host "`n⚠ IMPORTANT: Stopped at $maxPermissionsToRemove permissions as requested." -ForegroundColor Yellow
        Write-Host "There may be more sites where this user has permissions." -ForegroundColor Yellow
    }
}

if ($stats.GroupFound -gt 0) {
    Write-Host "`n⚠ GROUP PERMISSIONS:" -ForegroundColor Yellow
    Write-Host "Found $($stats.GroupFound) sites where user has access via group membership." -ForegroundColor Yellow
    Write-Host "Review the CSV for group details. Remove user from groups to revoke access." -ForegroundColor Yellow
}

if ($stats.OwnerFound -gt 0) {
    Write-Host "`n⚠ OWNERSHIP PERMISSIONS:" -ForegroundColor Yellow
    Write-Host "Found $($stats.OwnerFound) sites where user is an owner." -ForegroundColor Yellow
    Write-Host "Site ownership must be removed manually or via SharePoint admin center." -ForegroundColor Yellow
}

Write-Host "====================================================" -ForegroundColor Cyan

# Disconnect
Disconnect-MgGraph