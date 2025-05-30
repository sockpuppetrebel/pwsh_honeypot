# Quick SharePoint permissions check - focused on likely sites
# Faster approach: Check specific sites first, then expand if needed

param(
    [string]$UserEmail = "kaila.trapani@optimizely.com",
    [int]$MaxSites = 100,
    [switch]$OnlyKnownSites = $false
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

Write-Host "`n=============== QUICK PERMISSIONS CHECK ===============" -ForegroundColor Cyan
Write-Host "User: $UserEmail" -ForegroundColor White
Write-Host "Mode: $(if ($OnlyKnownSites) { 'Known sites only' } else { "First $MaxSites sites" })" -ForegroundColor Yellow

# Known sites where user has access
$knownSites = @(
    "https://episerver99.sharepoint.com/sites/APACOppDusk",
    "https://episerver99.sharepoint.com/sites/NET5"
)

$findings = @()

# Function to check a single site
function Check-SitePermissions {
    param($SiteUrl, $UserEmail, $Cert)
    
    $result = [PSCustomObject]@{
        SiteUrl = $SiteUrl
        SiteTitle = ""
        Permissions = @()
        Error = $null
    }
    
    try {
        Connect-PnPOnline -Url $SiteUrl `
                          -ClientId "fe2a9efe-3000-4b02-96ea-344a2583dd52" `
                          -Tenant "3ec00d79-021a-42d4-aac8-dcb35973dff2" `
                          -Thumbprint $Cert.Thumbprint
        
        $web = Get-PnPWeb
        $result.SiteTitle = if ($web.Title) { $web.Title } else { Split-Path $SiteUrl -Leaf }
        
        # Check Site Collection Admin
        try {
            $siteAdmins = Get-PnPSiteCollectionAdmin
            if ($siteAdmins | Where-Object { $_.Email -eq $UserEmail }) {
                $result.Permissions += "Site Collection Administrator"
            }
        } catch { }
        
        # Check SharePoint Groups
        $allGroups = Get-PnPGroup
        foreach ($group in $allGroups) {
            try {
                $groupMembers = Get-PnPGroupMember -Group $group
                if ($groupMembers | Where-Object { $_.Email -eq $UserEmail }) {
                    $groupType = "SharePoint Group"
                    if ($group.Id -eq $web.AssociatedOwnerGroup.Id) {
                        $groupType = "Site Owner Group"
                    } elseif ($group.Id -eq $web.AssociatedMemberGroup.Id) {
                        $groupType = "Site Member Group"
                    } elseif ($group.Id -eq $web.AssociatedVisitorGroup.Id) {
                        $groupType = "Site Visitor Group"
                    }
                    $result.Permissions += "$groupType - $($group.Title)"
                }
            } catch { }
        }
        
        Disconnect-PnPOnline
        
    } catch {
        $result.Error = $_.Exception.Message
    }
    
    return $result
}

# 1. Check known sites first
Write-Host "`nChecking known sites..." -ForegroundColor Yellow
foreach ($site in $knownSites) {
    Write-Host "Checking: $site" -ForegroundColor Gray
    $result = Check-SitePermissions -SiteUrl $site -UserEmail $UserEmail -Cert $cert
    
    if ($result.Permissions.Count -gt 0) {
        Write-Host "✓ FOUND: $($result.SiteTitle)" -ForegroundColor Green
        $result.Permissions | ForEach-Object {
            Write-Host "  - $_" -ForegroundColor Green
        }
        $findings += $result
    } elseif ($result.Error) {
        Write-Host "✗ ERROR: $($result.Error)" -ForegroundColor Red
    } else {
        Write-Host "  No permissions found" -ForegroundColor Gray
    }
}

# 2. If not only known sites, check more
if (-not $OnlyKnownSites) {
    Write-Host "`nGetting additional sites to check (limit: $MaxSites)..." -ForegroundColor Yellow
    
    # Connect to admin center
    Connect-PnPOnline -Url "https://episerver99-admin.sharepoint.com" `
                      -ClientId "fe2a9efe-3000-4b02-96ea-344a2583dd52" `
                      -Tenant "3ec00d79-021a-42d4-aac8-dcb35973dff2" `
                      -Thumbprint $cert.Thumbprint
    
    # Get sites (excluding OneDrive and already checked)
    $additionalSites = Get-PnPTenantSite | 
        Where-Object { 
            $_.Template -notlike "REDIRECT*" -and 
            $_.Url -notlike "*-my.sharepoint.com*" -and
            $_.Url -notin $knownSites
        } |
        Select-Object -First ($MaxSites - $knownSites.Count) |
        Select-Object -ExpandProperty Url
    
    Disconnect-PnPOnline
    
    Write-Host "Checking $($additionalSites.Count) additional sites..." -ForegroundColor Yellow
    
    $count = 0
    foreach ($site in $additionalSites) {
        $count++
        Write-Host "`r[$count/$($additionalSites.Count)] Checking sites..." -NoNewline
        
        $result = Check-SitePermissions -SiteUrl $site -UserEmail $UserEmail -Cert $cert
        
        if ($result.Permissions.Count -gt 0) {
            Write-Host "`n✓ FOUND: $($result.SiteTitle) - $($result.SiteUrl)" -ForegroundColor Green
            $result.Permissions | ForEach-Object {
                Write-Host "  - $_" -ForegroundColor Green
            }
            $findings += $result
        }
    }
}

# Summary
Write-Host "`n`n=============== SUMMARY ===============" -ForegroundColor Cyan
Write-Host "Sites with permissions found: $($findings.Count)" -ForegroundColor White

if ($findings.Count -gt 0) {
    Write-Host "`nPermissions found:" -ForegroundColor Yellow
    $findings | ForEach-Object {
        Write-Host "`n• $($_.SiteTitle)" -ForegroundColor White
        Write-Host "  URL: $($_.SiteUrl)" -ForegroundColor Gray
        $_.Permissions | ForEach-Object {
            Write-Host "  - $_" -ForegroundColor Green
        }
    }
    
    # Save to CSV
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $csvPath = Join-Path $scriptPath "quick_permissions_${timestamp}.csv"
    
    $csvData = @()
    foreach ($finding in $findings) {
        foreach ($perm in $finding.Permissions) {
            $csvData += [PSCustomObject]@{
                SiteUrl = $finding.SiteUrl
                SiteTitle = $finding.SiteTitle
                Permission = $perm
                UserEmail = $UserEmail
            }
        }
    }
    
    $csvData | Export-Csv -Path $csvPath -NoTypeInformation
    Write-Host "`nResults saved to: $csvPath" -ForegroundColor Green
    
    # Prompt for full scan
    if (-not $OnlyKnownSites -and $findings.Count -gt 0) {
        Write-Host "`n⚠ Found permissions in sample. User likely has access to more sites." -ForegroundColor Yellow
        Write-Host "Run the full audit script overnight for complete results." -ForegroundColor Yellow
    }
} else {
    Write-Host "`nNo permissions found in checked sites." -ForegroundColor Yellow
}

# Clean up certificate
try {
    $store = New-Object System.Security.Cryptography.X509Certificates.X509Store("My", "CurrentUser")
    $store.Open("ReadWrite")
    $certToRemove = $store.Certificates | Where-Object { $_.Thumbprint -eq $cert.Thumbprint }
    if ($certToRemove) {
        $store.Remove($certToRemove)
    }
    $store.Close()
} catch { }

Write-Host "`n✓ Complete!" -ForegroundColor Green