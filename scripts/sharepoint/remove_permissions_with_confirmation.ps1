# Comprehensive script to audit and remove ALL user permissions from SharePoint sites
# Checks: Direct permissions, Group memberships, Site ownership
# Features: Audit-first approach with confirmation before any removals
# Optimized for large numbers of sites with minimal memory usage

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

# Prompt for user to check
Write-Host "`n=============== USER SELECTION ===============" -ForegroundColor Cyan
$upn = Read-Host "Enter the user's email address to check permissions for"

# Validate email format
if ($upn -notmatch '^[\w\.-]+@[\w\.-]+\.\w+$') {
    Write-Host "Invalid email format. Please run the script again with a valid email address." -ForegroundColor Red
    Disconnect-MgGraph
    exit 1
}

Write-Host "Will check permissions for: $upn" -ForegroundColor Green
$pageSize = 100
$sitesChecked = 0

# Statistics
$stats = @{
    DirectFound = 0
    GroupFound = 0
    OwnerFound = 0
    ErrorCount = 0
}

# Collections for storing findings
$findings = @{
    Direct = @()
    Group = @()
    Owner = @()
    Errors = @()
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

# Progress tracking
$startTime = Get-Date

Write-Host "`n=============== PHASE 1: DISCOVERY ===============" -ForegroundColor Cyan
Write-Host "Scanning all sites for permissions..." -ForegroundColor Yellow
Write-Host "This audit will check:" -ForegroundColor Yellow
Write-Host "  - Direct user permissions" -ForegroundColor White
Write-Host "  - Group-based permissions (via $($userGroups.Count) groups)" -ForegroundColor White
Write-Host "  - Site ownership" -ForegroundColor White
Write-Host "" -NoNewline

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
        $sitesChecked++
        
        # Show progress every 50 sites
        if ($sitesChecked % 50 -eq 0) {
            $elapsed = (Get-Date) - $startTime
            $rate = $sitesChecked / $elapsed.TotalMinutes
            Write-Host "`rChecked: $sitesChecked | Direct: $($stats.DirectFound) | Group: $($stats.GroupFound) | Owner: $($stats.OwnerFound) | Rate: $([math]::Round($rate, 1))/min" -NoNewline
        }

        try {
            # Get all permissions for the site
            $sitePermissions = Get-MgSitePermission -SiteId $site.id -All -ErrorAction Stop
            
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
                    $stats.DirectFound++
                    
                    $findings.Direct += [PSCustomObject]@{
                        SiteId = $site.id
                        SiteName = $site.displayName
                        SiteUrl = $site.webUrl
                        PermissionId = $perm.Id
                        Roles = $perm.Roles -join ', '
                    }
                    
                    Write-Host "`n✓ DIRECT permission found at site ${sitesChecked}: $($site.displayName)" -ForegroundColor Green
                }
                
                # Check group permissions
                if ($perm.GrantedToIdentitiesV2 -and $userGroupIds.Count -gt 0) {
                    foreach ($identity in $perm.GrantedToIdentitiesV2) {
                        if ($identity.Group -and $userGroupIds -contains $identity.Group.Id) {
                            $stats.GroupFound++
                            
                            $findings.Group += [PSCustomObject]@{
                                SiteId = $site.id
                                SiteName = $site.displayName
                                SiteUrl = $site.webUrl
                                GroupId = $identity.Group.Id
                                GroupName = $identity.Group.DisplayName
                                Roles = $perm.Roles -join ', '
                            }
                            
                            Write-Host "`n⚠ GROUP permission found at site ${sitesChecked}: $($site.displayName)" -ForegroundColor Yellow
                            Write-Host "  Via group: $($identity.Group.DisplayName)" -ForegroundColor Yellow
                            break
                        }
                    }
                }
            }
            
            # Check if user is site owner
            try {
                $siteDetails = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/v1.0/sites/$($site.id)?`$expand=owners" -Method GET -ErrorAction SilentlyContinue
                
                if ($siteDetails.owners) {
                    foreach ($owner in $siteDetails.owners) {
                        if ($owner.userPrincipalName -eq $upn -or $owner.id -eq $user.Id) {
                            $stats.OwnerFound++
                            
                            $findings.Owner += [PSCustomObject]@{
                                SiteId = $site.id
                                SiteName = $site.displayName
                                SiteUrl = $site.webUrl
                            }
                            
                            Write-Host "`n⚠ OWNER permission found at site ${sitesChecked}: $($site.displayName)" -ForegroundColor Magenta
                            break
                        }
                    }
                }
            } catch {
                # Owner check failed, continue
            }

        } catch {
            $stats.ErrorCount++
            $findings.Errors += [PSCustomObject]@{
                SiteName = $site.displayName
                SiteUrl = $site.webUrl
                Error = $_.Exception.Message
            }
        }
    }

} while ($nextLink)

# Calculate stats
$elapsed = (Get-Date) - $startTime
$rate = $sitesChecked / $elapsed.TotalMinutes

Write-Host "`n`n=============== DISCOVERY COMPLETE ===============" -ForegroundColor Cyan
Write-Host "Total sites scanned: $sitesChecked" -ForegroundColor Cyan
Write-Host "Time taken: $([math]::Round($elapsed.TotalMinutes, 1)) minutes" -ForegroundColor Cyan
Write-Host "Average rate: $([math]::Round($rate, 1)) sites/minute" -ForegroundColor Cyan

# Save all findings to CSV
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$csvPath = Join-Path $scriptPath "permissions_audit_${timestamp}.csv"

$allFindings = @()
$findings.Direct | ForEach-Object { 
    $allFindings += $_ | Select-Object *, @{Name="Type";Expression={"Direct"}}, @{Name="Status";Expression={"Found"}}
}
$findings.Group | ForEach-Object { 
    $allFindings += $_ | Select-Object *, @{Name="Type";Expression={"Group"}}, @{Name="Status";Expression={"Found"}}
}
$findings.Owner | ForEach-Object { 
    $allFindings += $_ | Select-Object *, @{Name="Type";Expression={"Owner"}}, @{Name="Status";Expression={"Found"}}
}

if ($allFindings.Count -gt 0) {
    $allFindings | Export-Csv -Path $csvPath -NoTypeInformation
    Write-Host "`nDetailed findings saved to: $csvPath" -ForegroundColor Green
}

# Display findings summary
Write-Host "`n=============== FINDINGS SUMMARY ===============" -ForegroundColor Cyan

if ($findings.Direct.Count -gt 0) {
    Write-Host "`nDIRECT PERMISSIONS ($($findings.Direct.Count) found):" -ForegroundColor Yellow
    $findings.Direct | Select-Object -First 10 | ForEach-Object {
        Write-Host "  • $($_.SiteName)" -ForegroundColor White
        Write-Host "    URL: $($_.SiteUrl)" -ForegroundColor Gray
        Write-Host "    Roles: $($_.Roles)" -ForegroundColor Gray
    }
    if ($findings.Direct.Count -gt 10) {
        Write-Host "  ... and $($findings.Direct.Count - 10) more" -ForegroundColor Gray
    }
}

if ($findings.Group.Count -gt 0) {
    Write-Host "`nGROUP-BASED PERMISSIONS ($($findings.Group.Count) found):" -ForegroundColor Yellow
    $groupSummary = $findings.Group | Group-Object -Property GroupName
    $groupSummary | Select-Object -First 5 | ForEach-Object {
        Write-Host "  • Group: $($_.Name) (affects $($_.Count) sites)" -ForegroundColor White
    }
    if ($groupSummary.Count -gt 5) {
        Write-Host "  ... and $($groupSummary.Count - 5) more groups" -ForegroundColor Gray
    }
}

if ($findings.Owner.Count -gt 0) {
    Write-Host "`nOWNERSHIP PERMISSIONS ($($findings.Owner.Count) found):" -ForegroundColor Yellow
    $findings.Owner | Select-Object -First 10 | ForEach-Object {
        Write-Host "  • $($_.SiteName)" -ForegroundColor White
        Write-Host "    URL: $($_.SiteUrl)" -ForegroundColor Gray
    }
    if ($findings.Owner.Count -gt 10) {
        Write-Host "  ... and $($findings.Owner.Count - 10) more" -ForegroundColor Gray
    }
}

if ($findings.Direct.Count -eq 0 -and $findings.Group.Count -eq 0 -and $findings.Owner.Count -eq 0) {
    Write-Host "`nNo permissions found for $upn" -ForegroundColor Yellow
    Write-Host "====================================================" -ForegroundColor Cyan
    Disconnect-MgGraph
    exit 0
}

Write-Host "====================================================" -ForegroundColor Cyan

# Prompt for action
Write-Host "`n=============== ACTION REQUIRED ===============" -ForegroundColor Yellow
Write-Host "Found:" -ForegroundColor Cyan
Write-Host "  - $($findings.Direct.Count) direct permissions" -ForegroundColor White
Write-Host "  - $($findings.Group.Count) group-based permissions" -ForegroundColor White
Write-Host "  - $($findings.Owner.Count) ownership permissions" -ForegroundColor White

$choice = Read-Host "`nDo you want to remove these permissions? (Y/N)"

if ($choice -ne 'Y' -and $choice -ne 'y') {
    Write-Host "`nNo action taken - permissions remain unchanged" -ForegroundColor Yellow
    Write-Host "Audit results saved to: $csvPath" -ForegroundColor Green
    Disconnect-MgGraph
    exit 0
}

# Removal options
Write-Host "`nSelect removal option:" -ForegroundColor Cyan
Write-Host "1. Remove ALL direct permissions ($($findings.Direct.Count))" -ForegroundColor White
Write-Host "2. Select specific direct permissions to remove" -ForegroundColor White
Write-Host "3. Export findings only (no removal)" -ForegroundColor White
Write-Host "4. Cancel - don't remove anything" -ForegroundColor White

$removalChoice = Read-Host "`nEnter your choice (1-4)"

$removedCount = 0
$selectedPermissions = @()

switch ($removalChoice) {
    '1' {
        # Remove all direct permissions
        $selectedPermissions = $findings.Direct
    }
    '2' {
        # Select specific permissions
        Write-Host "`nAvailable direct permissions to remove:" -ForegroundColor Cyan
        for ($i = 0; $i -lt $findings.Direct.Count; $i++) {
            $perm = $findings.Direct[$i]
            Write-Host "$($i+1). $($perm.SiteName) ($($perm.Roles))" -ForegroundColor White
        }
        
        $selection = Read-Host "`nEnter the numbers to remove (comma-separated, e.g., 1,3,5) or 'all'"
        
        if ($selection -eq 'all') {
            $selectedPermissions = $findings.Direct
        } else {
            $indices = $selection -split ',' | ForEach-Object { [int]$_.Trim() - 1 }
            $selectedPermissions = $indices | ForEach-Object { $findings.Direct[$_] }
        }
    }
    '3' {
        Write-Host "`nFindings exported to: $csvPath" -ForegroundColor Green
        if ($findings.Group.Count -gt 0) {
            Write-Host "`n⚠ NOTE: Group permissions require removing user from groups" -ForegroundColor Yellow
        }
        if ($findings.Owner.Count -gt 0) {
            Write-Host "⚠ NOTE: Ownership permissions require SharePoint admin center access" -ForegroundColor Yellow
        }
        Disconnect-MgGraph
        exit 0
    }
    default {
        Write-Host "`nCancelled - no permissions removed" -ForegroundColor Yellow
        Disconnect-MgGraph
        exit 0
    }
}

# Confirm removal
if ($selectedPermissions.Count -gt 0) {
    Write-Host "`nConfirm removal of $($selectedPermissions.Count) permission(s):" -ForegroundColor Yellow
    $selectedPermissions | Select-Object -First 10 | ForEach-Object {
        Write-Host "  - $($_.SiteName)" -ForegroundColor White
    }
    if ($selectedPermissions.Count -gt 10) {
        Write-Host "  ... and $($selectedPermissions.Count - 10) more" -ForegroundColor Gray
    }
    
    $confirm = Read-Host "`nAre you sure? (Y/N)"
    
    if ($confirm -eq 'Y' -or $confirm -eq 'y') {
        Write-Host "`nRemoving permissions..." -ForegroundColor Yellow
        
        foreach ($perm in $selectedPermissions) {
            try {
                Remove-MgSitePermission -SiteId $perm.SiteId -PermissionId $perm.PermissionId -ErrorAction Stop
                Write-Host "✓ Removed permission from: $($perm.SiteName)" -ForegroundColor Green
                $removedCount++
                
                # Update status in findings
                $allFindings | Where-Object { $_.SiteId -eq $perm.SiteId -and $_.PermissionId -eq $perm.PermissionId } | ForEach-Object {
                    $_.Status = "Removed"
                }
            } catch {
                Write-Host "✗ Failed to remove from $($perm.SiteName): $_" -ForegroundColor Red
            }
        }
        
        # Update CSV with removal status
        if ($removedCount -gt 0) {
            $allFindings | Export-Csv -Path $csvPath -NoTypeInformation -Force
            Write-Host "`n✓ Successfully removed $removedCount permission(s)" -ForegroundColor Green
            Write-Host "Updated audit log: $csvPath" -ForegroundColor Green
        }
    } else {
        Write-Host "`nCancelled - no permissions removed" -ForegroundColor Yellow
    }
}

# Final notes
if ($findings.Group.Count -gt 0) {
    Write-Host "`n⚠ GROUP PERMISSIONS:" -ForegroundColor Yellow
    Write-Host "Found $($findings.Group.Count) sites where user has access via group membership." -ForegroundColor Yellow
    Write-Host "To remove these, remove user from the relevant groups." -ForegroundColor Yellow
}

if ($findings.Owner.Count -gt 0) {
    Write-Host "`n⚠ OWNERSHIP PERMISSIONS:" -ForegroundColor Yellow
    Write-Host "Found $($findings.Owner.Count) sites where user is an owner." -ForegroundColor Yellow
    Write-Host "Site ownership must be removed manually via SharePoint admin center." -ForegroundColor Yellow
}

Write-Host "`n=============== OPERATION COMPLETE ===============" -ForegroundColor Cyan

# Disconnect
Disconnect-MgGraph