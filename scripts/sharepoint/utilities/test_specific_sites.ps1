# Test script to check specific sites with detailed debugging
# This will help determine if it's a permissions issue

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

# Get user info
Write-Host "`nGetting user information..." -ForegroundColor Yellow
try {
    $user = Get-MgUser -Filter "userPrincipalName eq '$upn'" -ErrorAction Stop
    Write-Host "✓ Found user: $($user.DisplayName) (ID: $($user.Id))" -ForegroundColor Green
    
    # Get user's groups
    $userGroups = Get-MgUserMemberOf -UserId $user.Id -All | Where-Object { $_.AdditionalProperties.'@odata.type' -eq '#microsoft.graph.group' }
    $userGroupIds = $userGroups | ForEach-Object { $_.Id }
    Write-Host "✓ User is member of $($userGroups.Count) groups" -ForegroundColor Green
} catch {
    Write-Host "✗ Could not find user: $_" -ForegroundColor Red
    exit 1
}

# Test specific sites
$testSites = @(
    "https://episerver99.sharepoint.com/sites/APACOppDusk",
    "https://episerver99.sharepoint.com/sites/NET5"
)

foreach ($siteUrl in $testSites) {
    Write-Host "`n=============== Testing Site ===============" -ForegroundColor Cyan
    Write-Host "URL: $siteUrl" -ForegroundColor White
    
    try {
        # Method 1: Try to get site by URL
        Write-Host "`nMethod 1: Getting site by URL..." -ForegroundColor Yellow
        $encodedUrl = [System.Web.HttpUtility]::UrlEncode($siteUrl)
        $site = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/v1.0/sites/episerver99.sharepoint.com:/sites/$(Split-Path $siteUrl -Leaf)" -Method GET
        
        Write-Host "✓ Found site: $($site.displayName)" -ForegroundColor Green
        Write-Host "  Site ID: $($site.id)" -ForegroundColor Gray
        
        # Get all permissions for this site
        Write-Host "`nGetting all permissions for this site..." -ForegroundColor Yellow
        $permissions = Get-MgSitePermission -SiteId $site.id -All -ErrorAction Stop
        Write-Host "✓ Found $($permissions.Count) total permissions on this site" -ForegroundColor Green
        
        # Check each permission in detail
        $userFound = $false
        Write-Host "`nChecking each permission:" -ForegroundColor Yellow
        
        foreach ($perm in $permissions) {
            $permDetails = "Permission ID: $($perm.Id)"
            
            # Check GrantedToIdentities
            if ($perm.GrantedToIdentities) {
                foreach ($identity in $perm.GrantedToIdentities) {
                    if ($identity.User) {
                        if ($identity.User.UserPrincipalName -eq $upn) {
                            Write-Host "✓ DIRECT PERMISSION FOUND!" -ForegroundColor Green
                            Write-Host "  User: $($identity.User.DisplayName) ($($identity.User.UserPrincipalName))" -ForegroundColor Green
                            Write-Host "  Roles: $($perm.Roles -join ', ')" -ForegroundColor Green
                            $userFound = $true
                        }
                    }
                    if ($identity.Group -and $userGroupIds -contains $identity.Group.Id) {
                        Write-Host "✓ GROUP PERMISSION FOUND!" -ForegroundColor Yellow
                        Write-Host "  Group: $($identity.Group.DisplayName)" -ForegroundColor Yellow
                        Write-Host "  Roles: $($perm.Roles -join ', ')" -ForegroundColor Yellow
                        $userFound = $true
                    }
                }
            }
            
            # Check GrantedToIdentitiesV2
            if ($perm.GrantedToIdentitiesV2) {
                foreach ($identity in $perm.GrantedToIdentitiesV2) {
                    if ($identity.User) {
                        if ($identity.User.UserPrincipalName -eq $upn) {
                            Write-Host "✓ DIRECT PERMISSION FOUND (V2)!" -ForegroundColor Green
                            Write-Host "  User: $($identity.User.DisplayName) ($($identity.User.UserPrincipalName))" -ForegroundColor Green
                            Write-Host "  Roles: $($perm.Roles -join ', ')" -ForegroundColor Green
                            $userFound = $true
                        }
                    }
                    if ($identity.Group -and $userGroupIds -contains $identity.Group.Id) {
                        Write-Host "✓ GROUP PERMISSION FOUND (V2)!" -ForegroundColor Yellow
                        Write-Host "  Group: $($identity.Group.DisplayName)" -ForegroundColor Yellow
                        Write-Host "  Roles: $($perm.Roles -join ', ')" -ForegroundColor Yellow
                        $userFound = $true
                    }
                }
            }
        }
        
        if (-not $userFound) {
            Write-Host "✗ User has NO permissions on this site" -ForegroundColor Red
            
            # List first 5 permissions to see what's there
            Write-Host "`nShowing first 5 permissions on this site:" -ForegroundColor Yellow
            $permissions | Select-Object -First 5 | ForEach-Object {
                $perm = $_
                Write-Host "`n  Permission ID: $($perm.Id)" -ForegroundColor Gray
                Write-Host "  Roles: $($perm.Roles -join ', ')" -ForegroundColor Gray
                
                if ($perm.GrantedToIdentitiesV2) {
                    $perm.GrantedToIdentitiesV2 | ForEach-Object {
                        if ($_.User) {
                            Write-Host "    Granted to user: $($_.User.DisplayName)" -ForegroundColor Gray
                        }
                        if ($_.Group) {
                            Write-Host "    Granted to group: $($_.Group.DisplayName)" -ForegroundColor Gray
                        }
                    }
                }
            }
        }
        
        # Try alternate permission check methods
        Write-Host "`nTrying alternate permission APIs..." -ForegroundColor Yellow
        
        # Check site owners
        try {
            $siteWithOwners = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/v1.0/sites/$($site.id)?`$expand=owners" -Method GET
            if ($siteWithOwners.owners) {
                Write-Host "Site has $($siteWithOwners.owners.Count) owners" -ForegroundColor Yellow
                $siteWithOwners.owners | ForEach-Object {
                    if ($_.userPrincipalName -eq $upn) {
                        Write-Host "✓ User is a SITE OWNER!" -ForegroundColor Magenta
                        $userFound = $true
                    }
                }
            }
        } catch {
            Write-Host "Could not check site owners: $_" -ForegroundColor Gray
        }
        
    } catch {
        Write-Host "✗ Error accessing site: $_" -ForegroundColor Red
        Write-Host "Error type: $($_.Exception.GetType().FullName)" -ForegroundColor Red
        
        # If it's a 403, it's definitely a permissions issue
        if ($_.Exception.Message -match "403" -or $_.Exception.Message -match "Forbidden") {
            Write-Host "`n⚠ 403 FORBIDDEN: The app does NOT have permission to access this site's permissions!" -ForegroundColor Red
            Write-Host "This confirms an app registration permission issue." -ForegroundColor Red
        }
    }
}

Write-Host "`n=============== DIAGNOSTIC SUMMARY ===============" -ForegroundColor Cyan
Write-Host "If you see 403 errors above, the app registration lacks required permissions." -ForegroundColor Yellow
Write-Host "If you see 0 permissions but no errors, the user may not have access via Graph-visible methods." -ForegroundColor Yellow
Write-Host "SharePoint permissions can also be set at levels not exposed via Graph API." -ForegroundColor Yellow

# Disconnect
Disconnect-MgGraph