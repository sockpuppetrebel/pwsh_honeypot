# Fast approach: Find user's Azure AD groups, then check SharePoint sites connected to those groups
# This is MUCH faster than checking every site

param(
    [string]$UserEmail = "kaila.trapani@optimizely.com"
)

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
    exit 1
}

Write-Host "Connected via: $($context.AuthType)" -ForegroundColor Green

Write-Host "`n=============== FAST GROUP-BASED REMOVAL ===============" -ForegroundColor Cyan
Write-Host "User: $UserEmail" -ForegroundColor White

# Step 1: Get user and their groups
Write-Host "`nStep 1: Getting user's Azure AD groups..." -ForegroundColor Yellow
try {
    $user = Get-MgUser -Filter "userPrincipalName eq '$UserEmail'" -ErrorAction Stop
    Write-Host "✓ Found user: $($user.DisplayName) (ID: $($user.Id))" -ForegroundColor Green
    
    # Get all groups the user is a member of
    $userGroups = Get-MgUserMemberOf -UserId $user.Id -All | Where-Object { $_.AdditionalProperties.'@odata.type' -eq '#microsoft.graph.group' }
    Write-Host "✓ User is member of $($userGroups.Count) Azure AD groups" -ForegroundColor Green
    
} catch {
    Write-Host "✗ Could not find user: $_" -ForegroundColor Red
    exit 1
}

# Step 2: For Microsoft 365 groups, check if they have SharePoint sites
Write-Host "`nStep 2: Checking which groups have SharePoint sites..." -ForegroundColor Yellow
$groupsWithSites = @()
$m365Groups = $userGroups | Where-Object { $_.AdditionalProperties.groupTypes -contains "Unified" }

Write-Host "Found $($m365Groups.Count) Microsoft 365 groups (these can have SharePoint sites)" -ForegroundColor White

foreach ($group in $m365Groups) {
    try {
        # Try to get the group's SharePoint site
        $site = Get-MgGroupSite -GroupId $group.Id -ErrorAction SilentlyContinue
        if ($site) {
            $groupsWithSites += [PSCustomObject]@{
                GroupId = $group.Id
                GroupName = $group.AdditionalProperties.displayName
                SiteUrl = $site.WebUrl
                SiteId = $site.Id
            }
            Write-Host "✓ Found SharePoint site for group: $($group.AdditionalProperties.displayName)" -ForegroundColor Green
        }
    } catch {
        # No site for this group
    }
}

Write-Host "`nFound $($groupsWithSites.Count) groups with SharePoint sites" -ForegroundColor Cyan

# Step 3: Show all groups and ask what to do
Write-Host "`n=============== USER'S GROUP MEMBERSHIPS ===============" -ForegroundColor Cyan
Write-Host "`nMicrosoft 365 Groups with SharePoint sites ($($groupsWithSites.Count)):" -ForegroundColor Yellow
$groupsWithSites | ForEach-Object {
    Write-Host "  • $($_.GroupName)" -ForegroundColor White
    Write-Host "    Site: $($_.SiteUrl)" -ForegroundColor Gray
}

Write-Host "`nOther Azure AD Groups ($($userGroups.Count - $m365Groups.Count)):" -ForegroundColor Yellow
$otherGroups = $userGroups | Where-Object { $_.AdditionalProperties.groupTypes -notcontains "Unified" }
$otherGroups | Select-Object -First 10 | ForEach-Object {
    Write-Host "  • $($_.AdditionalProperties.displayName)" -ForegroundColor White
}
if ($otherGroups.Count -gt 10) {
    Write-Host "  ... and $($otherGroups.Count - 10) more" -ForegroundColor Gray
}

# Step 4: Removal options
Write-Host "`n=============== REMOVAL OPTIONS ===============" -ForegroundColor Yellow
Write-Host "What would you like to do?" -ForegroundColor Cyan
Write-Host "1. Remove user from ALL groups ($($userGroups.Count) groups)" -ForegroundColor White
Write-Host "2. Remove user from Microsoft 365 groups with SharePoint sites only ($($groupsWithSites.Count) groups)" -ForegroundColor White
Write-Host "3. Remove user from specific groups (select from list)" -ForegroundColor White
Write-Host "4. Export group list only (no removal)" -ForegroundColor White
Write-Host "5. Cancel" -ForegroundColor White

$choice = Read-Host "`nEnter your choice (1-5)"

$groupsToRemove = @()

switch ($choice) {
    '1' {
        $groupsToRemove = $userGroups
        Write-Host "`nWill remove user from ALL $($groupsToRemove.Count) groups" -ForegroundColor Yellow
    }
    '2' {
        $groupsToRemove = $m365Groups | Where-Object { $_.Id -in $groupsWithSites.GroupId }
        Write-Host "`nWill remove user from $($groupsToRemove.Count) Microsoft 365 groups with SharePoint sites" -ForegroundColor Yellow
    }
    '3' {
        Write-Host "`nSelect groups to remove user from:" -ForegroundColor Cyan
        $allGroups = $userGroups | Sort-Object { $_.AdditionalProperties.displayName }
        
        for ($i = 0; $i -lt $allGroups.Count; $i++) {
            $group = $allGroups[$i]
            $hassite = if ($group.Id -in $groupsWithSites.GroupId) { " [HAS SHAREPOINT SITE]" } else { "" }
            Write-Host "$($i+1). $($group.AdditionalProperties.displayName)$hassite" -ForegroundColor White
        }
        
        $selection = Read-Host "`nEnter group numbers to remove (comma-separated, e.g., 1,3,5)"
        $indices = $selection -split ',' | ForEach-Object { [int]$_.Trim() - 1 }
        $groupsToRemove = $indices | ForEach-Object { $allGroups[$_] }
        
        Write-Host "`nWill remove user from $($groupsToRemove.Count) selected groups" -ForegroundColor Yellow
    }
    '4' {
        # Export to CSV
        $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
        $csvPath = Join-Path $scriptPath "user_groups_${timestamp}.csv"
        
        $exportData = $userGroups | ForEach-Object {
            [PSCustomObject]@{
                GroupName = $_.AdditionalProperties.displayName
                GroupId = $_.Id
                GroupType = if ($_.AdditionalProperties.groupTypes -contains "Unified") { "Microsoft 365" } else { "Security/Other" }
                HasSharePointSite = if ($_.Id -in $groupsWithSites.GroupId) { "Yes" } else { "No" }
                SiteUrl = ($groupsWithSites | Where-Object { $_.GroupId -eq $_.Id }).SiteUrl
            }
        }
        
        $exportData | Export-Csv -Path $csvPath -NoTypeInformation
        Write-Host "`n✓ Group list exported to: $csvPath" -ForegroundColor Green
        Disconnect-MgGraph
        exit 0
    }
    default {
        Write-Host "`nCancelled - no changes made" -ForegroundColor Yellow
        Disconnect-MgGraph
        exit 0
    }
}

# Confirm removal
if ($groupsToRemove.Count -gt 0) {
    Write-Host "`n⚠ CONFIRMATION REQUIRED" -ForegroundColor Yellow
    Write-Host "Remove $UserEmail from $($groupsToRemove.Count) groups?" -ForegroundColor Cyan
    $groupsToRemove | Select-Object -First 5 | ForEach-Object {
        Write-Host "  - $($_.AdditionalProperties.displayName)" -ForegroundColor White
    }
    if ($groupsToRemove.Count -gt 5) {
        Write-Host "  ... and $($groupsToRemove.Count - 5) more" -ForegroundColor Gray
    }
    
    $confirm = Read-Host "`nType 'YES' to confirm removal"
    
    if ($confirm -eq 'YES') {
        Write-Host "`nRemoving user from groups..." -ForegroundColor Yellow
        $removed = 0
        $failed = 0
        
        foreach ($group in $groupsToRemove) {
            try {
                Remove-MgGroupMemberByRef -GroupId $group.Id -DirectoryObjectId $user.Id -ErrorAction Stop
                Write-Host "✓ Removed from: $($group.AdditionalProperties.displayName)" -ForegroundColor Green
                $removed++
            } catch {
                Write-Host "✗ Failed to remove from $($group.AdditionalProperties.displayName): $_" -ForegroundColor Red
                $failed++
            }
        }
        
        Write-Host "`n=============== SUMMARY ===============" -ForegroundColor Cyan
        Write-Host "Successfully removed from: $removed groups" -ForegroundColor Green
        if ($failed -gt 0) {
            Write-Host "Failed to remove from: $failed groups" -ForegroundColor Red
        }
        
        # Note about SharePoint sync
        Write-Host "`n⚠ IMPORTANT NOTES:" -ForegroundColor Yellow
        Write-Host "1. Azure AD group changes may take 1-24 hours to sync to SharePoint" -ForegroundColor White
        Write-Host "2. User may still have SharePoint-only permissions (not in Azure AD)" -ForegroundColor White
        Write-Host "3. For immediate removal, use the PnP script on specific sites" -ForegroundColor White
        
    } else {
        Write-Host "`nCancelled - no changes made" -ForegroundColor Yellow
    }
}

# Disconnect
Disconnect-MgGraph

Write-Host "`n✓ Complete!" -ForegroundColor Green