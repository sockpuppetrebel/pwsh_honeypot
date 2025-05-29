# Check SharePoint permissions using PnP PowerShell
# This will check ALL permission types including SharePoint-specific ones

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
Write-Host "Adding certificate to CurrentUser store temporarily..." -ForegroundColor Yellow
$store = New-Object System.Security.Cryptography.X509Certificates.X509Store("My", "CurrentUser")
$store.Open("ReadWrite")
$store.Add($cert)
$store.Close()

$upn = 'kaila.trapani@optimizely.com'

# Test specific sites
$testSites = @(
    "https://episerver99.sharepoint.com/sites/APACOppDusk",
    "https://episerver99.sharepoint.com/sites/NET5"
)

Write-Host "`n=============== CHECKING PERMISSIONS WITH PnP ===============" -ForegroundColor Cyan
Write-Host "User: $upn" -ForegroundColor White

foreach ($siteUrl in $testSites) {
    Write-Host "`n━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor DarkGray
    Write-Host "SITE: $siteUrl" -ForegroundColor Cyan
    Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor DarkGray
    
    try {
        # Connect to the specific site
        Write-Host "`nConnecting to site..." -ForegroundColor Yellow
        Connect-PnPOnline -Url $siteUrl `
                          -ClientId "fe2a9efe-3000-4b02-96ea-344a2583dd52" `
                          -Tenant "3ec00d79-021a-42d4-aac8-dcb35973dff2" `
                          -Thumbprint $cert.Thumbprint
        
        $web = Get-PnPWeb
        Write-Host "✓ Connected to: $($web.Title)" -ForegroundColor Green
        
        # 1. Check if user is Site Collection Administrator
        Write-Host "`n1. Checking Site Collection Administrators..." -ForegroundColor Yellow
        $siteAdmins = Get-PnPSiteCollectionAdmin
        $isAdmin = $siteAdmins | Where-Object { $_.Email -eq $upn }
        if ($isAdmin) {
            Write-Host "   ✓ USER IS SITE COLLECTION ADMINISTRATOR!" -ForegroundColor Green -BackgroundColor DarkGreen
        } else {
            Write-Host "   ✗ User is NOT a site collection admin" -ForegroundColor Gray
            Write-Host "   Current admins:" -ForegroundColor Gray
            $siteAdmins | Select-Object -First 3 | ForEach-Object {
                Write-Host "     - $($_.Title) ($($_.Email))" -ForegroundColor Gray
            }
        }
        
        # 2. Check site owners group
        Write-Host "`n2. Checking Site Owners Group..." -ForegroundColor Yellow
        try {
            $ownersGroup = Get-PnPGroup -AssociatedOwnerGroup
            $owners = Get-PnPGroupMember -Group $ownersGroup
            $isOwner = $owners | Where-Object { $_.Email -eq $upn }
            if ($isOwner) {
                Write-Host "   ✓ USER IS IN SITE OWNERS GROUP!" -ForegroundColor Green
            } else {
                Write-Host "   ✗ User is NOT in owners group" -ForegroundColor Gray
            }
        } catch {
            Write-Host "   Could not check owners group: $_" -ForegroundColor Gray
        }
        
        # 3. Check site members group
        Write-Host "`n3. Checking Site Members Group..." -ForegroundColor Yellow
        try {
            $membersGroup = Get-PnPGroup -AssociatedMemberGroup
            $members = Get-PnPGroupMember -Group $membersGroup
            $isMember = $members | Where-Object { $_.Email -eq $upn }
            if ($isMember) {
                Write-Host "   ✓ USER IS IN SITE MEMBERS GROUP!" -ForegroundColor Green
            } else {
                Write-Host "   ✗ User is NOT in members group" -ForegroundColor Gray
            }
        } catch {
            Write-Host "   Could not check members group: $_" -ForegroundColor Gray
        }
        
        # 4. Check all SharePoint groups
        Write-Host "`n4. Checking All SharePoint Groups..." -ForegroundColor Yellow
        $allGroups = Get-PnPGroup
        $userGroups = @()
        
        foreach ($group in $allGroups) {
            try {
                $groupMembers = Get-PnPGroupMember -Group $group
                $userInGroup = $groupMembers | Where-Object { $_.Email -eq $upn }
                if ($userInGroup) {
                    $userGroups += $group
                    Write-Host "   ✓ Found in group: $($group.Title)" -ForegroundColor Green
                }
            } catch {
                # Skip groups we can't read
            }
        }
        
        if ($userGroups.Count -eq 0) {
            Write-Host "   ✗ User not found in any SharePoint groups" -ForegroundColor Gray
        }
        
        # 5. Check direct permissions on site
        Write-Host "`n5. Checking Direct Web Permissions..." -ForegroundColor Yellow
        $user = Get-PnPUser -Identity $upn -ErrorAction SilentlyContinue
        if ($user) {
            Write-Host "   ✓ User found in site user information list" -ForegroundColor Green
            Write-Host "   User ID: $($user.Id)" -ForegroundColor Gray
            
            # Get role assignments
            $roleAssignments = Get-PnPWeb -Includes RoleAssignments
            foreach ($ra in $roleAssignments.RoleAssignments) {
                if ($ra.PrincipalId -eq $user.Id) {
                    Write-Host "   ✓ DIRECT PERMISSION FOUND!" -ForegroundColor Green
                    $ra.RoleDefinitionBindings | ForEach-Object {
                        Write-Host "     - Role: $($_.Name)" -ForegroundColor Green
                    }
                }
            }
        } else {
            Write-Host "   ✗ User not found in site user list" -ForegroundColor Gray
        }
        
        # 6. Check unique permissions on lists/libraries
        Write-Host "`n6. Checking List/Library Permissions..." -ForegroundColor Yellow
        $lists = Get-PnPList | Where-Object { $_.Hidden -eq $false }
        $listsWithAccess = 0
        
        foreach ($list in $lists) {
            if ($list.HasUniqueRoleAssignments) {
                try {
                    $listPerms = Get-PnPList -Identity $list -Includes RoleAssignments
                    foreach ($ra in $listPerms.RoleAssignments) {
                        if ($ra.PrincipalId -eq $user.Id) {
                            Write-Host "   ✓ Has permissions on list: $($list.Title)" -ForegroundColor Green
                            $listsWithAccess++
                        }
                    }
                } catch {
                    # Skip lists we can't check
                }
            }
        }
        
        if ($listsWithAccess -eq 0) {
            Write-Host "   No special list/library permissions found" -ForegroundColor Gray
        }
        
        # Disconnect from this site
        Disconnect-PnPOnline
        
    } catch {
        Write-Host "✗ Error checking site: $_" -ForegroundColor Red
        Write-Host "Error type: $($_.Exception.GetType().FullName)" -ForegroundColor Red
    }
}

Write-Host "`n━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor DarkGray
Write-Host "SUMMARY" -ForegroundColor Cyan
Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor DarkGray
Write-Host "PnP PowerShell can detect SharePoint-specific permissions that" -ForegroundColor Yellow
Write-Host "Microsoft Graph API cannot see, including:" -ForegroundColor Yellow
Write-Host "- Site Collection Administrators" -ForegroundColor White
Write-Host "- SharePoint Groups (different from Azure AD groups)" -ForegroundColor White
Write-Host "- Direct web/list permissions" -ForegroundColor White
Write-Host "- Classic SharePoint permission inheritance" -ForegroundColor White

# Clean up certificate
Write-Host "`nCleaning up certificate..." -ForegroundColor Yellow
try {
    $store = New-Object System.Security.Cryptography.X509Certificates.X509Store("My", "CurrentUser")
    $store.Open("ReadWrite")
    $certToRemove = $store.Certificates | Where-Object { $_.Thumbprint -eq $cert.Thumbprint }
    if ($certToRemove) {
        $store.Remove($certToRemove)
        Write-Host "Certificate removed from store" -ForegroundColor Gray
    }
    $store.Close()
} catch {
    Write-Host "Warning: Could not remove certificate: $_" -ForegroundColor Yellow
}