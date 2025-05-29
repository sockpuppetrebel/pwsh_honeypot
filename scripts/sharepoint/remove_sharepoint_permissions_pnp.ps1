# Comprehensive SharePoint permissions removal script using PnP PowerShell
# Handles SharePoint Groups, Site Collection Admin, and all permission types
#5/25 Confirmation prompts added prior to any removal 

param(
    [string]$UserEmail,
    [switch]$WhatIf = $false
)

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

# Add certificate to store temporarily
$store = New-Object System.Security.Cryptography.X509Certificates.X509Store("My", "CurrentUser")
$store.Open("ReadWrite")
$store.Add($cert)
$store.Close()

# Prompt for user email if not provided
if (-not $UserEmail) {
    Write-Host "`n=============== USER SELECTION ===============" -ForegroundColor Cyan
    $UserEmail = Read-Host "Enter the user's email address to check permissions for"
}

# Validate email format
if ($UserEmail -notmatch '^[\w\.-]+@[\w\.-]+\.\w+$') {
    Write-Host "Invalid email format." -ForegroundColor Red
    exit 1
}

Write-Host "`n=============== SHAREPOINT PERMISSIONS AUDIT ===============" -ForegroundColor Cyan
Write-Host "User: $UserEmail" -ForegroundColor White
Write-Host "Mode: $(if ($WhatIf) { 'AUDIT ONLY (WhatIf mode)' } else { 'LIVE - Will prompt before removal' })" -ForegroundColor Yellow

# Initialize collections
$findings = @{
    SiteCollectionAdmin = @()
    SiteOwnerGroup = @()
    SiteMemberGroup = @()
    SiteVisitorGroup = @()
    OtherSharePointGroups = @()
    DirectPermissions = @()
    Errors = @()
}

$stats = @{
    SitesChecked = 0
    TotalFindings = 0
    StartTime = Get-Date
}

# Connect to admin center first to get all site collections
Write-Host "`nConnecting to SharePoint Admin Center..." -ForegroundColor Yellow
try {
    Connect-PnPOnline -Url "https://episerver99-admin.sharepoint.com" `
                      -ClientId "fe2a9efe-3000-4b02-96ea-344a2583dd52" `
                      -Tenant "3ec00d79-021a-42d4-aac8-dcb35973dff2" `
                      -Thumbprint $cert.Thumbprint
    
    Write-Host "✓ Connected to admin center" -ForegroundColor Green
} catch {
    Write-Host "✗ Failed to connect to admin center: $_" -ForegroundColor Red
    exit 1
}

# Get all site collections
Write-Host "`nGetting all site collections..." -ForegroundColor Yellow
$allSites = Get-PnPTenantSite -IncludeOneDriveSites | Where-Object { $_.Template -notlike "REDIRECT*" }
$totalSites = $allSites.Count
Write-Host "Found $totalSites site collections to check" -ForegroundColor Green

# Disconnect from admin center
Disconnect-PnPOnline

# Process each site
foreach ($site in $allSites) {
    $stats.SitesChecked++
    
    # Progress indicator
    if ($stats.SitesChecked % 10 -eq 0 -or $stats.SitesChecked -eq 1) {
        $elapsed = (Get-Date) - $stats.StartTime
        $rate = [math]::Round($stats.SitesChecked / $elapsed.TotalMinutes, 1)
        Write-Host "`r[$($stats.SitesChecked)/$totalSites] Checking sites... Found $($stats.TotalFindings) permissions | Rate: $rate sites/min" -NoNewline
    }
    
    try {
        # Connect to the site
        Connect-PnPOnline -Url $site.Url `
                          -ClientId "fe2a9efe-3000-4b02-96ea-344a2583dd52" `
                          -Tenant "3ec00d79-021a-42d4-aac8-dcb35973dff2" `
                          -Thumbprint $cert.Thumbprint
        
        $web = Get-PnPWeb
        $siteTitle = if ($web.Title) { $web.Title } else { Split-Path $site.Url -Leaf }
        
        # 1. Check Site Collection Admin
        try {
            $siteAdmins = Get-PnPSiteCollectionAdmin
            $isAdmin = $siteAdmins | Where-Object { $_.Email -eq $UserEmail }
            if ($isAdmin) {
                $findings.SiteCollectionAdmin += [PSCustomObject]@{
                    SiteUrl = $site.Url
                    SiteTitle = $siteTitle
                    PermissionType = "Site Collection Administrator"
                    Details = "Full control of site collection"
                }
                $stats.TotalFindings++
                Write-Host "`n✓ SITE COLLECTION ADMIN: $siteTitle" -ForegroundColor Green -BackgroundColor DarkGreen
            }
        } catch {
            # Skip if can't check admins
        }
        
        # 2. Check SharePoint Groups
        $allGroups = Get-PnPGroup
        
        foreach ($group in $allGroups) {
            try {
                $groupMembers = Get-PnPGroupMember -Group $group
                $userInGroup = $groupMembers | Where-Object { $_.Email -eq $UserEmail }
                
                if ($userInGroup) {
                    $groupType = "OtherSharePointGroups"
                    $highlight = $false
                    
                    # Categorize the group
                    if ($group.Id -eq $web.AssociatedOwnerGroup.Id) {
                        $groupType = "SiteOwnerGroup"
                        $highlight = $true
                    } elseif ($group.Id -eq $web.AssociatedMemberGroup.Id) {
                        $groupType = "SiteMemberGroup"
                        $highlight = $true
                    } elseif ($group.Id -eq $web.AssociatedVisitorGroup.Id) {
                        $groupType = "SiteVisitorGroup"
                    }
                    
                    $findings.$groupType += [PSCustomObject]@{
                        SiteUrl = $site.Url
                        SiteTitle = $siteTitle
                        GroupId = $group.Id
                        GroupName = $group.Title
                        GroupType = $groupType
                        LoginName = $userInGroup.LoginName
                    }
                    $stats.TotalFindings++
                    
                    if ($highlight) {
                        Write-Host "`n✓ $($groupType.ToUpper()): $siteTitle - $($group.Title)" -ForegroundColor Green
                    }
                }
            } catch {
                # Skip groups we can't read
            }
        }
        
        # 3. Check direct permissions
        $user = Get-PnPUser -Identity $UserEmail -ErrorAction SilentlyContinue
        if ($user) {
            $roleAssignments = Get-PnPWeb -Includes RoleAssignments
            foreach ($ra in $roleAssignments.RoleAssignments) {
                if ($ra.PrincipalId -eq $user.Id) {
                    $findings.DirectPermissions += [PSCustomObject]@{
                        SiteUrl = $site.Url
                        SiteTitle = $siteTitle
                        UserId = $user.Id
                        PermissionType = "Direct Web Permission"
                    }
                    $stats.TotalFindings++
                    Write-Host "`n✓ DIRECT PERMISSION: $siteTitle" -ForegroundColor Green
                }
            }
        }
        
        # Disconnect from this site
        Disconnect-PnPOnline
        
    } catch {
        $findings.Errors += [PSCustomObject]@{
            SiteUrl = $site.Url
            Error = $_.Exception.Message
        }
    }
}

# Final progress update
$elapsed = (Get-Date) - $stats.StartTime
Write-Host "`n`n=============== AUDIT COMPLETE ===============" -ForegroundColor Cyan
Write-Host "Sites checked: $($stats.SitesChecked)" -ForegroundColor White
Write-Host "Total findings: $($stats.TotalFindings)" -ForegroundColor White
Write-Host "Time taken: $([math]::Round($elapsed.TotalMinutes, 1)) minutes" -ForegroundColor White

# Display findings summary
Write-Host "`n=============== FINDINGS SUMMARY ===============" -ForegroundColor Cyan

if ($findings.SiteCollectionAdmin.Count -gt 0) {
    Write-Host "`nSITE COLLECTION ADMIN ($($findings.SiteCollectionAdmin.Count) sites):" -ForegroundColor Yellow -BackgroundColor DarkRed
    $findings.SiteCollectionAdmin | ForEach-Object {
        Write-Host "  • $($_.SiteTitle)" -ForegroundColor White
        Write-Host "    $($_.SiteUrl)" -ForegroundColor Gray
    }
}

if ($findings.SiteOwnerGroup.Count -gt 0) {
    Write-Host "`nSITE OWNER GROUPS ($($findings.SiteOwnerGroup.Count) sites):" -ForegroundColor Yellow
    $findings.SiteOwnerGroup | ForEach-Object {
        Write-Host "  • $($_.SiteTitle) - $($_.GroupName)" -ForegroundColor White
        Write-Host "    $($_.SiteUrl)" -ForegroundColor Gray
    }
}

if ($findings.SiteMemberGroup.Count -gt 0) {
    Write-Host "`nSITE MEMBER GROUPS ($($findings.SiteMemberGroup.Count) sites):" -ForegroundColor Green
    $findings.SiteMemberGroup | Select-Object -First 10 | ForEach-Object {
        Write-Host "  • $($_.SiteTitle) - $($_.GroupName)" -ForegroundColor White
    }
    if ($findings.SiteMemberGroup.Count -gt 10) {
        Write-Host "  ... and $($findings.SiteMemberGroup.Count - 10) more" -ForegroundColor Gray
    }
}

if ($findings.SiteVisitorGroup.Count -gt 0) {
    Write-Host "`nSITE VISITOR GROUPS ($($findings.SiteVisitorGroup.Count) sites):" -ForegroundColor Cyan
    $findings.SiteVisitorGroup | Select-Object -First 5 | ForEach-Object {
        Write-Host "  • $($_.SiteTitle) - $($_.GroupName)" -ForegroundColor White
    }
    if ($findings.SiteVisitorGroup.Count -gt 5) {
        Write-Host "  ... and $($findings.SiteVisitorGroup.Count - 5) more" -ForegroundColor Gray
    }
}

if ($findings.OtherSharePointGroups.Count -gt 0) {
    Write-Host "`nOTHER SHAREPOINT GROUPS ($($findings.OtherSharePointGroups.Count)):" -ForegroundColor Magenta
    $findings.OtherSharePointGroups | Select-Object -First 5 | ForEach-Object {
        Write-Host "  • $($_.SiteTitle) - $($_.GroupName)" -ForegroundColor White
    }
    if ($findings.OtherSharePointGroups.Count -gt 5) {
        Write-Host "  ... and $($findings.OtherSharePointGroups.Count - 5) more" -ForegroundColor Gray
    }
}

if ($findings.DirectPermissions.Count -gt 0) {
    Write-Host "`nDIRECT PERMISSIONS ($($findings.DirectPermissions.Count) sites):" -ForegroundColor Yellow
    $findings.DirectPermissions | ForEach-Object {
        Write-Host "  • $($_.SiteTitle)" -ForegroundColor White
    }
}

# Save findings to CSV
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$csvPath = Join-Path $scriptPath "sharepoint_permissions_pnp_${timestamp}.csv"

$allFindings = @()
foreach ($type in $findings.Keys) {
    if ($type -ne "Errors") {
        $findings.$type | ForEach-Object {
            $allFindings += $_ | Select-Object *, @{Name="FindingType";Expression={$type}}
        }
    }
}

if ($allFindings.Count -gt 0) {
    $allFindings | Export-Csv -Path $csvPath -NoTypeInformation
    Write-Host "`nFindings saved to: $csvPath" -ForegroundColor Green
}

# If WhatIf mode, we're done
if ($WhatIf) {
    Write-Host "`n✓ Audit complete (WhatIf mode - no changes made)" -ForegroundColor Green
} elseif ($stats.TotalFindings -eq 0) {
    Write-Host "`nNo permissions found for $UserEmail" -ForegroundColor Yellow
} else {
    # Prompt for removal
    Write-Host "`n=============== REMOVAL OPTIONS ===============" -ForegroundColor Yellow
    Write-Host "Found $($stats.TotalFindings) total permissions for $UserEmail" -ForegroundColor Cyan
    
    $choice = Read-Host "`nDo you want to remove these permissions? (Y/N)"
    
    if ($choice -eq 'Y' -or $choice -eq 'y') {
        Write-Host "`nSelect what to remove:" -ForegroundColor Cyan
        Write-Host "1. Remove ALL permissions" -ForegroundColor White
        Write-Host "2. Remove only SharePoint group memberships" -ForegroundColor White
        Write-Host "3. Remove only Site Collection Admin rights" -ForegroundColor White
        Write-Host "4. Select specific sites" -ForegroundColor White
        Write-Host "5. Cancel" -ForegroundColor White
        
        $removalChoice = Read-Host "`nEnter your choice (1-5)"
        
        $removed = 0
        
        switch ($removalChoice) {
            '1' {
                # Remove all permissions
                Write-Host "`nRemoving all permissions..." -ForegroundColor Yellow
                
                # Process each site with permissions
                $sitesToProcess = @()
                $sitesToProcess += $findings.SiteCollectionAdmin | Select-Object -ExpandProperty SiteUrl -Unique
                $sitesToProcess += $findings.SiteOwnerGroup | Select-Object -ExpandProperty SiteUrl -Unique
                $sitesToProcess += $findings.SiteMemberGroup | Select-Object -ExpandProperty SiteUrl -Unique
                $sitesToProcess += $findings.SiteVisitorGroup | Select-Object -ExpandProperty SiteUrl -Unique
                $sitesToProcess += $findings.OtherSharePointGroups | Select-Object -ExpandProperty SiteUrl -Unique
                $sitesToProcess += $findings.DirectPermissions | Select-Object -ExpandProperty SiteUrl -Unique
                $sitesToProcess = $sitesToProcess | Select-Object -Unique
                
                foreach ($siteUrl in $sitesToProcess) {
                    Write-Host "`nProcessing: $siteUrl" -ForegroundColor Cyan
                    
                    try {
                        Connect-PnPOnline -Url $siteUrl `
                                          -ClientId "fe2a9efe-3000-4b02-96ea-344a2583dd52" `
                                          -Tenant "3ec00d79-021a-42d4-aac8-dcb35973dff2" `
                                          -Thumbprint $cert.Thumbprint
                        
                        # Remove from Site Collection Admin
                        $siteAdminRecord = $findings.SiteCollectionAdmin | Where-Object { $_.SiteUrl -eq $siteUrl }
                        if ($siteAdminRecord) {
                            try {
                                Remove-PnPSiteCollectionAdmin -Owners $UserEmail
                                Write-Host "  ✓ Removed Site Collection Admin" -ForegroundColor Green
                                $removed++
                            } catch {
                                Write-Host "  ✗ Failed to remove Site Collection Admin: $_" -ForegroundColor Red
                            }
                        }
                        
                        # Remove from SharePoint groups
                        $groupsToRemove = @()
                        $groupsToRemove += $findings.SiteOwnerGroup | Where-Object { $_.SiteUrl -eq $siteUrl }
                        $groupsToRemove += $findings.SiteMemberGroup | Where-Object { $_.SiteUrl -eq $siteUrl }
                        $groupsToRemove += $findings.SiteVisitorGroup | Where-Object { $_.SiteUrl -eq $siteUrl }
                        $groupsToRemove += $findings.OtherSharePointGroups | Where-Object { $_.SiteUrl -eq $siteUrl }
                        
                        foreach ($groupRecord in $groupsToRemove) {
                            try {
                                Remove-PnPGroupMember -Group $groupRecord.GroupId -LoginName $groupRecord.LoginName
                                Write-Host "  ✓ Removed from group: $($groupRecord.GroupName)" -ForegroundColor Green
                                $removed++
                            } catch {
                                Write-Host "  ✗ Failed to remove from $($groupRecord.GroupName): $_" -ForegroundColor Red
                            }
                        }
                        
                        Disconnect-PnPOnline
                        
                    } catch {
                        Write-Host "  ✗ Error processing site: $_" -ForegroundColor Red
                    }
                }
            }
            '2' {
                # Remove only group memberships
                Write-Host "`nRemoving SharePoint group memberships only..." -ForegroundColor Yellow
                
                $groupFindings = @()
                $groupFindings += $findings.SiteOwnerGroup
                $groupFindings += $findings.SiteMemberGroup
                $groupFindings += $findings.SiteVisitorGroup
                $groupFindings += $findings.OtherSharePointGroups
                
                $siteGroups = $groupFindings | Group-Object -Property SiteUrl
                
                foreach ($site in $siteGroups) {
                    Write-Host "`nProcessing: $($site.Name)" -ForegroundColor Cyan
                    
                    try {
                        Connect-PnPOnline -Url $site.Name `
                                          -ClientId "fe2a9efe-3000-4b02-96ea-344a2583dd52" `
                                          -Tenant "3ec00d79-021a-42d4-aac8-dcb35973dff2" `
                                          -Thumbprint $cert.Thumbprint
                        
                        foreach ($groupRecord in $site.Group) {
                            try {
                                Remove-PnPGroupMember -Group $groupRecord.GroupId -LoginName $groupRecord.LoginName
                                Write-Host "  ✓ Removed from: $($groupRecord.GroupName)" -ForegroundColor Green
                                $removed++
                            } catch {
                                Write-Host "  ✗ Failed to remove from $($groupRecord.GroupName): $_" -ForegroundColor Red
                            }
                        }
                        
                        Disconnect-PnPOnline
                        
                    } catch {
                        Write-Host "  ✗ Error connecting to site: $_" -ForegroundColor Red
                    }
                }
            }
            default {
                Write-Host "`nCancelled - no permissions removed" -ForegroundColor Yellow
            }
        }
        
        if ($removed -gt 0) {
            Write-Host "`n✓ Successfully removed $removed permission(s)" -ForegroundColor Green
            
            # Update CSV with removal status
            $allFindings | ForEach-Object { 
                Add-Member -InputObject $_ -MemberType NoteProperty -Name "Status" -Value "Removed" -Force
            }
            $allFindings | Export-Csv -Path $csvPath -NoTypeInformation -Force
        }
    } else {
        Write-Host "`nNo action taken - permissions remain unchanged" -ForegroundColor Yellow
    }
}

# Clean up certificate
Write-Host "`nCleaning up..." -ForegroundColor Yellow
try {
    $store = New-Object System.Security.Cryptography.X509Certificates.X509Store("My", "CurrentUser")
    $store.Open("ReadWrite")
    $certToRemove = $store.Certificates | Where-Object { $_.Thumbprint -eq $cert.Thumbprint }
    if ($certToRemove) {
        $store.Remove($certToRemove)
    }
    $store.Close()
} catch {
    # Ignore cleanup errors
}

Write-Host "`n✓ Complete!" -ForegroundColor Green