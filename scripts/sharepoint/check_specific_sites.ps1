# Check permissions for specific SharePoint sites
# This script checks only the sites you specify, not all sites

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

# ADD THE SPECIFIC SITE URLS YOU WANT TO CHECK HERE
$specificSiteUrls = @(
    # Example sites - replace with actual URLs you know she has access to
    "https://episerver99.sharepoint.com/sites/APACOppDusk",
    "https://episerver99.sharepoint.com/sites/NET5"
)

$removedFrom = @(*)
$permissionsRemoved = 0[..]
$sitesChecked = 0

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

Write-Host "Checking $($specificSiteUrls.Count) specific sites for $upn..." -ForegroundColor Cyan
Write-Host "Audit log will be saved to: $csvPath" -ForegroundColor Cyan

foreach ($siteUrl in $specificSiteUrls) {
    $sitesChecked++
    Write-Host "`nChecking site ${sitesChecked}: $siteUrl"
    
    try {
        # Get site by URL
        $encodedUrl = [System.Web.HttpUtility]::UrlEncode($siteUrl)
        $siteResponse = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/v1.0/sites?`$filter=webUrl eq '$encodedUrl'" -Method GET
        
        if ($siteResponse.value -and $siteResponse.value.Count -gt 0) {
            $site = $siteResponse.value[0]
            Write-Host "Found site: $($site.displayName) (ID: $($site.id))"
            
            # Check permissions
            try {
                $permissions = Get-MgSitePermission -SiteId $site.id -All
                Write-Host "Total permissions on site: $($permissions.Count)"
                
                $userPermissions = $permissions | Where-Object {
                    $_.GrantedToIdentities.User.UserPrincipalName -eq $upn
                }
                
                if ($userPermissions) {
                    Write-Host "✓ Found $($userPermissions.Count) permission(s) for $upn" -ForegroundColor Green
                    
                    foreach ($perm in $userPermissions) {
                        Write-Host "  - Permission ID: $($perm.Id)" -ForegroundColor Cyan
                        Write-Host "    Roles: $($perm.Roles -join ', ')" -ForegroundColor Cyan
                        if ($perm.GrantedToIdentities) {
                            foreach ($identity in $perm.GrantedToIdentities) {
                                if ($identity.User) {
                                    Write-Host "    User: $($identity.User.DisplayName) ($($identity.User.UserPrincipalName))" -ForegroundColor Cyan
                                }
                            }
                        }
                    }
                    
                    # Log to CSV
                    $csvData += [PSCustomObject]@{
                        SiteName = $site.displayName
                        SiteUrl = $site.webUrl
                        PermissionFound = $true
                        PermissionRemoved = $false
                        UserPrincipalName = $upn
                    }
                } else {
                    Write-Host "✗ No permissions found for $upn" -ForegroundColor Yellow
                    
                    # Log to CSV
                    $csvData += [PSCustomObject]@{
                        SiteName = $site.displayName
                        SiteUrl = $site.webUrl
                        PermissionFound = $false
                        PermissionRemoved = $false
                        UserPrincipalName = $upn
                    }
                }
                
            } catch {
                Write-Host "Error checking permissions: $($_.Exception.Message)" -ForegroundColor Red
                
                # Log error to CSV
                $csvData += [PSCustomObject]@{
                    SiteName = $site.displayName
                    SiteUrl = $site.webUrl
                    PermissionFound = "ERROR"
                    PermissionRemoved = "ERROR: $($_.Exception.Message)"
                    UserPrincipalName = $upn
                }
            }
            
        } else {
            Write-Host "✗ Site not found or no access to site" -ForegroundColor Red
            
            # Log to CSV
            $csvData += [PSCustomObject]@{
                SiteName = "Unknown"
                SiteUrl = $siteUrl
                PermissionFound = "SITE_NOT_FOUND"
                PermissionRemoved = "N/A"
                UserPrincipalName = $upn
            }
        }
        
    } catch {
        Write-Host "Error accessing site: $($_.Exception.Message)" -ForegroundColor Red
        
        # Log error to CSV
        $csvData += [PSCustomObject]@{
            SiteName = "Unknown"
            SiteUrl = $siteUrl
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
    PermissionRemoved = "N/A"
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
Write-Host "User: $upn" -ForegroundColor Cyan
Write-Host "======================================" -ForegroundColor Cyan

# Disconnect
Disconnect-MgGraph -NoWelcome