# Check permissions for specific SharePoint sites - Version 2
# This script uses site IDs or relative URLs to check permissions

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

# CONFIGURATION - MODIFY THESE
$upn = 'kaila.trapani@optimizely.com'

# ADD THE SPECIFIC SITES YOU WANT TO CHECK HERE
# You can use either site names or full URLs
$specificSites = @(
    "APACOppDusk",
    "NET5"
)

$removedFrom = @()
$permissionsRemoved = 0
$sitesChecked = 0
$sitesWithPermission = 0

# Initialize CSV logging
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$csvPath = Join-Path $scriptPath "specific_sites_audit_${timestamp}.csv"
$csvData = @()

# Add CSV headers
$csvData += [PSCustomObject]@{
    SiteName = "AUDIT_STARTED"
    SiteUrl = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    PermissionFound = "N/A"
    PermissionRemoved = "N/A"
    UserPrincipalName = $upn
}

Write-Host "`nChecking $($specificSites.Count) specific sites for $upn..." -ForegroundColor Cyan
Write-Host "Audit log will be saved to: $csvPath" -ForegroundColor Cyan

# First, let's try a different approach - get the user's ID
Write-Host "`nGetting user information for $upn..." -ForegroundColor Yellow
try {
    $user = Get-MgUser -Filter "userPrincipalName eq '$upn'" -ErrorAction Stop
    Write-Host "✓ Found user: $($user.DisplayName) (ID: $($user.Id))" -ForegroundColor Green
} catch {
    Write-Host "✗ Could not find user $upn" -ForegroundColor Red
    Write-Host "Error: $_" -ForegroundColor Red
}

foreach ($siteName in $specificSites) {
    $sitesChecked++
    Write-Host "`n[$sitesChecked/$($specificSites.Count)] Checking site: $siteName" -ForegroundColor Cyan
    
    try {
        # Try to find the site by searching for it
        Write-Host "  Searching for site..." -ForegroundColor Gray
        
        # First try exact match on display name
        $searchResponse = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/v1.0/sites?search=$siteName" -Method GET
        
        $foundSite = $null
        if ($searchResponse.value -and $searchResponse.value.Count -gt 0) {
            # Look for exact match first
            $foundSite = $searchResponse.value | Where-Object { $_.name -eq $siteName -or $_.displayName -eq $siteName } | Select-Object -First 1
            
            # If no exact match, take the first result
            if (-not $foundSite) {
                $foundSite = $searchResponse.value[0]
            }
        }
        
        if ($foundSite) {
            Write-Host "  ✓ Found site: $($foundSite.displayName)" -ForegroundColor Green
            Write-Host "    URL: $($foundSite.webUrl)" -ForegroundColor Gray
            Write-Host "    ID: $($foundSite.id)" -ForegroundColor Gray
            
            # Now check permissions
            Write-Host "  Checking permissions..." -ForegroundColor Yellow
            
            try {
                # Get all permissions for the site
                $allPermissions = Get-MgSitePermission -SiteId $foundSite.id -All -ErrorAction Stop
                Write-Host "  Total permissions on site: $($allPermissions.Count)" -ForegroundColor Gray
                
                # Check for user permissions
                $userPermissions = @()
                foreach ($perm in $allPermissions) {
                    if ($perm.GrantedToIdentitiesV2) {
                        foreach ($identity in $perm.GrantedToIdentitiesV2) {
                            if ($identity.User -and $identity.User.UserPrincipalName -eq $upn) {
                                $userPermissions += $perm
                            }
                        }
                    }
                    if ($perm.GrantedToIdentities) {
                        foreach ($identity in $perm.GrantedToIdentities) {
                            if ($identity.User -and $identity.User.UserPrincipalName -eq $upn) {
                                $userPermissions += $perm
                            }
                        }
                    }
                }
                
                if ($userPermissions.Count -gt 0) {
                    $sitesWithPermission++
                    Write-Host "  ✓ Found $($userPermissions.Count) permission(s) for $upn" -ForegroundColor Green
                    
                    foreach ($perm in $userPermissions) {
                        Write-Host "    - Permission ID: $($perm.Id)" -ForegroundColor Cyan
                        if ($perm.Roles) {
                            Write-Host "      Roles: $($perm.Roles -join ', ')" -ForegroundColor Cyan
                        }
                    }
                    
                    # Log to CSV
                    $csvData += [PSCustomObject]@{
                        SiteName = $foundSite.displayName
                        SiteUrl = $foundSite.webUrl
                        PermissionFound = $true
                        PermissionRemoved = $false
                        UserPrincipalName = $upn
                    }
                } else {
                    Write-Host "  ✗ No permissions found for $upn" -ForegroundColor Yellow
                    
                    # Let's also check if the user might be part of a group with permissions
                    Write-Host "  Checking for group permissions..." -ForegroundColor Gray
                    $groupPermissions = @()
                    foreach ($perm in $allPermissions) {
                        if ($perm.GrantedToIdentitiesV2) {
                            foreach ($identity in $perm.GrantedToIdentitiesV2) {
                                if ($identity.Group) {
                                    $groupPermissions += @{
                                        GroupName = $identity.Group.DisplayName
                                        GroupId = $identity.Group.Id
                                        Roles = $perm.Roles
                                    }
                                }
                            }
                        }
                    }
                    
                    if ($groupPermissions.Count -gt 0) {
                        Write-Host "  Found $($groupPermissions.Count) group(s) with permissions:" -ForegroundColor Yellow
                        $groupPermissions | ForEach-Object {
                            Write-Host "    - $($_.GroupName) (Roles: $($_.Roles -join ', '))" -ForegroundColor Gray
                        }
                    }
                    
                    # Log to CSV
                    $csvData += [PSCustomObject]@{
                        SiteName = $foundSite.displayName
                        SiteUrl = $foundSite.webUrl
                        PermissionFound = $false
                        PermissionRemoved = $false
                        UserPrincipalName = $upn
                    }
                }
                
            } catch {
                Write-Host "  ✗ Error checking permissions: $($_.Exception.Message)" -ForegroundColor Red
                
                # Log error to CSV
                $csvData += [PSCustomObject]@{
                    SiteName = $foundSite.displayName
                    SiteUrl = $foundSite.webUrl
                    PermissionFound = "ERROR"
                    PermissionRemoved = "ERROR: $($_.Exception.Message)"
                    UserPrincipalName = $upn
                }
            }
            
        } else {
            Write-Host "  ✗ Site not found" -ForegroundColor Red
            
            # Log to CSV
            $csvData += [PSCustomObject]@{
                SiteName = $siteName
                SiteUrl = "NOT_FOUND"
                PermissionFound = "SITE_NOT_FOUND"
                PermissionRemoved = "N/A"
                UserPrincipalName = $upn
            }
        }
        
    } catch {
        Write-Host "  ✗ Error accessing site: $($_.Exception.Message)" -ForegroundColor Red
        
        # Log error to CSV
        $csvData += [PSCustomObject]@{
            SiteName = $siteName
            SiteUrl = "ERROR"
            PermissionFound = "ERROR"
            PermissionRemoved = "ERROR: $($_.Exception.Message)"
            UserPrincipalName = $upn
        }
    }
}

# Add summary to CSV
$csvData += [PSCustomObject]@{
    SiteName = "AUDIT_COMPLETED"
    SiteUrl = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    PermissionFound = "Total Sites Checked: $sitesChecked"
    PermissionRemoved = "Sites with Permission: $sitesWithPermission"
    UserPrincipalName = $upn
}

# Export to CSV
try {
    $csvData | Export-Csv -Path $csvPath -NoTypeInformation
    Write-Host "`n✔ Audit log saved to: $csvPath" -ForegroundColor Green
} catch {
    Write-Host "`n✗ Failed to save audit log: $_" -ForegroundColor Red
}

Write-Host "`n" -NoNewline
Write-Host "=============== SUMMARY ===============" -ForegroundColor Cyan
Write-Host "Total sites checked: $sitesChecked" -ForegroundColor Cyan
Write-Host "Sites where user has permissions: $sitesWithPermission" -ForegroundColor Cyan
Write-Host "User: $upn" -ForegroundColor Cyan
Write-Host "======================================" -ForegroundColor Cyan

# Disconnect
Disconnect-MgGraph