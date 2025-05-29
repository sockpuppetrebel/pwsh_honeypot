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

# Initialize CSV logging
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$csvPath = Join-Path $scriptPath "site_permissions_audit_${timestamp}.csv"
$csvData = @()

# Add CSV headers
$csvData += [PSCustomObject]@{
    SiteName = "AUDIT_STARTED"
    SiteUrl = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    PermissionFound = "N/A"
    PermissionRemoved = "N/A"
    UserPrincipalName = $upn
}

Write-Host "Searching sites to remove $upn from up to $maxPermissionsToRemove site permissions..." -ForegroundColor Cyan
Write-Host "Audit log will be saved to: $csvPath" -ForegroundColor Cyan

# Get all sites using pagination
$nextLink = $null
$pageCount = 0

do {
    Write-Host "Fetching page $($pageCount + 1) of sites..."
    
    if ($nextLink) {
        # Use the nextLink for subsequent pages
        $response = Invoke-MgGraphRequest -Uri $nextLink -Method GET
    } else {
        # First page
        $response = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/v1.0/sites?`$top=$pageSize" -Method GET
    }
    
    $sites = $response.value
    $nextLink = $response.'@odata.nextLink'
    $pageCount++
    
    if (-not $sites -or $sites.Count -eq 0) {
        Write-Host "No more sites to check." -ForegroundColor Yellow
        break
    }
    
    Write-Host "Processing $($sites.Count) sites from page $pageCount..."

    foreach ($site in $sites) {
        if ($permissionsRemoved -ge $maxPermissionsToRemove) {
            Write-Host "Reached limit of $maxPermissionsToRemove permissions removed. Stopping." -ForegroundColor Yellow
            break
        }

        $sitesChecked++
        Write-Host "Checking site ${sitesChecked}: $($site.displayName) ($($site.webUrl))"

        try {
            $permsToRemove = Get-MgSitePermission -SiteId $site.id -All | Where-Object {
                $_.GrantedToIdentities.User.UserPrincipalName -eq $upn
            }

            $hadPermission = $false
            $wasRemoved = $false

            if ($permsToRemove) {
                $hadPermission = $true
                foreach ($perm in $permsToRemove) {
                    if ($permissionsRemoved -ge $maxPermissionsToRemove) {
                        Write-Host "Reached limit of $maxPermissionsToRemove permissions removed. Stopping." -ForegroundColor Yellow
                        break
                    }
                    
                    Remove-MgSitePermission -SiteId $site.id -PermissionId $perm.Id -ErrorAction Stop
                    $removedFrom += "$($site.displayName) : $($site.webUrl)"
                    $permissionsRemoved++
                    $wasRemoved = $true
                    Write-Host "✔ Removed permission ID $($perm.Id) ($permissionsRemoved/$maxPermissionsToRemove)" -ForegroundColor Green
                }
            }

            # Log to CSV
            $csvData += [PSCustomObject]@{
                SiteName = $site.displayName
                SiteUrl = $site.webUrl
                PermissionFound = $hadPermission
                PermissionRemoved = $wasRemoved
                UserPrincipalName = $upn
            }

        } catch {
            Write-Host "Error checking/removing from $($site.displayName): $($_.Exception.Message)" -ForegroundColor Red
            
            # Log error to CSV
            $csvData += [PSCustomObject]@{
                SiteName = $site.displayName
                SiteUrl = $site.webUrl
                PermissionFound = "ERROR"
                PermissionRemoved = "ERROR: $($_.Exception.Message)"
                UserPrincipalName = $upn
            }
        }
    }

} while ($permissionsRemoved -lt $maxPermissionsToRemove -and $nextLink)

# Add summary to CSV
$csvData += [PSCustomObject]@{
    SiteName = "AUDIT_COMPLETED"
    SiteUrl = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    PermissionFound = "Total Sites Checked: $sitesChecked"
    PermissionRemoved = "Total Permissions Removed: $permissionsRemoved"
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
Write-Host "Completed. Removed $permissionsRemoved permission(s) for $upn" -ForegroundColor Green

if ($permissionsRemoved -gt 0) {
    Write-Host "`nSites user was removed from:" -ForegroundColor Yellow
    $removedFrom | Select-Object -Unique | ForEach-Object { Write-Host "- $_" -ForegroundColor White }
    
    if ($permissionsRemoved -eq $maxPermissionsToRemove) {
        Write-Host "`n⚠ IMPORTANT: Stopped at $maxPermissionsToRemove permissions as requested." -ForegroundColor Yellow
        Write-Host "There may be more sites where this user has permissions." -ForegroundColor Yellow
        Write-Host "To remove ALL permissions, modify the `$maxPermissionsToRemove variable or remove the limit." -ForegroundColor Yellow
    }
} else {
    Write-Host "User had no permissions on any sites checked." -ForegroundColor Yellow
}
Write-Host "======================================" -ForegroundColor Cyan