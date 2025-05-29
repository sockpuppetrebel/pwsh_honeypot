# Optimized script to remove user permissions from SharePoint sites
# Handles large numbers of sites (7000+) with minimal memory usage

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

# Initialize CSV with minimal data - only sites with permissions or errors
$csvPath = Join-Path $scriptPath "site_permissions_found.csv"
$csvData = @()

# Progress tracking
$saveInterval = 10  # Save CSV every 10 sites
$lastSaveAt = 0
$startTime = Get-Date

Write-Host "Searching for $upn permissions across all sites..." -ForegroundColor Cyan
Write-Host "Will stop after removing from $maxPermissionsToRemove sites" -ForegroundColor Cyan
Write-Host "Progress will be saved to: $csvPath" -ForegroundColor Cyan
Write-Host "Only logging sites with permissions or errors to minimize data" -ForegroundColor Cyan
Write-Host "" -NoNewline

# Function to save CSV
function Save-Progress {
    param($Data, $Path, $SitesChecked, $PermissionsRemoved)
    
    # Add summary row
    $summaryData = $Data + @([PSCustomObject]@{
        SiteName = "PROGRESS_UPDATE"
        SiteUrl = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        PermissionFound = "Sites Checked: $SitesChecked"
        PermissionRemoved = "Permissions Removed: $PermissionsRemoved"
        UserPrincipalName = ""
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
        # Use the nextLink for subsequent pages
        $response = Invoke-MgGraphRequest -Uri $nextLink -Method GET
    } else {
        # First page
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
            Write-Host "`rChecked: $sitesChecked sites | Found permissions on: $($csvData.Count) sites | Removed: $permissionsRemoved | Rate: $([math]::Round($rate, 1)) sites/min" -NoNewline
        }

        try {
            $permsToRemove = Get-MgSitePermission -SiteId $site.id -All | Where-Object {
                $_.GrantedToIdentities.User.UserPrincipalName -eq $upn
            }

            if ($permsToRemove) {
                # Found permissions - log this
                Write-Host "`n✓ Found permissions at site ${sitesChecked}: $($site.displayName)" -ForegroundColor Green
                
                foreach ($perm in $permsToRemove) {
                    if ($permissionsRemoved -ge $maxPermissionsToRemove) {
                        break
                    }
                    
                    Remove-MgSitePermission -SiteId $site.id -PermissionId $perm.Id -ErrorAction Stop
                    $removedFrom += "$($site.displayName) : $($site.webUrl)"
                    $permissionsRemoved++
                    
                    # Add to CSV data
                    $csvData += [PSCustomObject]@{
                        SiteName = $site.displayName
                        SiteUrl = $site.webUrl
                        PermissionFound = $true
                        PermissionRemoved = $true
                        UserPrincipalName = $upn
                    }
                    
                    Write-Host "  ✔ Removed permission ID $($perm.Id) ($permissionsRemoved/$maxPermissionsToRemove)" -ForegroundColor Green
                }
            }
            # Don't log sites without permissions to minimize data

        } catch {
            # Only log errors
            Write-Host "`n✗ Error at site $sitesChecked ($($site.displayName)): $($_.Exception.Message)" -ForegroundColor Red
            
            $csvData += [PSCustomObject]@{
                SiteName = $site.displayName
                SiteUrl = $site.webUrl
                PermissionFound = "ERROR"
                PermissionRemoved = "ERROR: $($_.Exception.Message)"
                UserPrincipalName = $upn
            }
        }
        
        # Save progress every N sites
        if (($sitesChecked - $lastSaveAt) -ge $saveInterval) {
            Save-Progress -Data $csvData -Path $csvPath -SitesChecked $sitesChecked -PermissionsRemoved $permissionsRemoved
            $lastSaveAt = $sitesChecked
        }
    }

} while ($permissionsRemoved -lt $maxPermissionsToRemove -and $nextLink)

# Final save
Save-Progress -Data $csvData -Path $csvPath -SitesChecked $sitesChecked -PermissionsRemoved $permissionsRemoved

# Calculate stats
$elapsed = (Get-Date) - $startTime
$rate = $sitesChecked / $elapsed.TotalMinutes

Write-Host "`n`n" -NoNewline
Write-Host "=============== SUMMARY ===============" -ForegroundColor Cyan
Write-Host "Total sites checked: $sitesChecked" -ForegroundColor Cyan
Write-Host "Sites with permissions found: $($csvData.Count - 1)" -ForegroundColor Cyan  # -1 for summary row
Write-Host "Permissions removed: $permissionsRemoved" -ForegroundColor Cyan
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
} else {
    Write-Host "`nNo permissions found for $upn in the $sitesChecked sites checked." -ForegroundColor Yellow
}
Write-Host "======================================" -ForegroundColor Cyan

# Disconnect
Disconnect-MgGraph