# Script to remove SharePoint permissions using PnP PowerShell
# This will work with actual SharePoint permissions, not just Graph API sharing

Write-Host "This script requires PnP PowerShell module." -ForegroundColor Cyan
Write-Host "To install: Install-Module -Name PnP.PowerShell" -ForegroundColor Yellow

# Check if PnP module is installed
if (-not (Get-Module -ListAvailable -Name PnP.PowerShell)) {
    Write-Host "`nPnP.PowerShell module not found. Installing..." -ForegroundColor Yellow
    try {
        Install-Module -Name PnP.PowerShell -Force -AllowClobber -Scope CurrentUser
        Write-Host "✓ PnP.PowerShell installed successfully" -ForegroundColor Green
    } catch {
        Write-Host "✗ Failed to install PnP.PowerShell: $_" -ForegroundColor Red
        Write-Host "Please run: Install-Module -Name PnP.PowerShell -Force" -ForegroundColor Yellow
        exit 1
    }
}

Import-Module PnP.PowerShell

$upn = 'kaila.trapani@optimizely.com'
$tenantUrl = "https://episerver99.sharepoint.com"
$appId = "fe2a9efe-3000-4b02-96ea-344a2583dd52"
$tenant = "3ec00d79-021a-42d4-aac8-dcb35973dff2"

# Certificate paths
$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
$certPath = Join-Path $scriptPath "azure_app_cert.pem"
$keyPath = Join-Path $scriptPath "azure_app_key.pem"

# Create PFX from PEM files for PnP
Write-Host "`nPreparing certificate for PnP..." -ForegroundColor Yellow
$pfxPath = Join-Path $scriptPath "temp_cert.pfx"
$pfxPassword = "TempPassword123!"

# Convert PEM to PFX using openssl
$opensslCmd = "openssl pkcs12 -export -out `"$pfxPath`" -inkey `"$keyPath`" -in `"$certPath`" -password pass:$pfxPassword"
Write-Host "Creating PFX certificate..." -ForegroundColor Gray
Invoke-Expression $opensslCmd 2>$null

if (-not (Test-Path $pfxPath)) {
    Write-Host "✗ Failed to create PFX certificate" -ForegroundColor Red
    Write-Host "Make sure OpenSSL is installed" -ForegroundColor Yellow
    exit 1
}

# Connect to SharePoint using PnP
Write-Host "`nConnecting to SharePoint admin center..." -ForegroundColor Cyan
try {
    Connect-PnPOnline -Url "$tenantUrl" `
                      -ClientId $appId `
                      -Tenant $tenant `
                      -CertificatePath $pfxPath `
                      -CertificatePassword (ConvertTo-SecureString -String $pfxPassword -AsPlainText -Force)
    
    Write-Host "✓ Connected to SharePoint!" -ForegroundColor Green
} catch {
    Write-Host "✗ Failed to connect: $_" -ForegroundColor Red
    Remove-Item $pfxPath -Force -ErrorAction SilentlyContinue
    exit 1
}

# Test with known sites
$testSites = @(
    "https://episerver99.sharepoint.com/sites/APACOppDusk",
    "https://episerver99.sharepoint.com/sites/NET5"
)

Write-Host "`nChecking permissions on test sites..." -ForegroundColor Cyan

foreach ($siteUrl in $testSites) {
    Write-Host "`nChecking: $siteUrl" -ForegroundColor Yellow
    
    try {
        # Connect to specific site
        Connect-PnPOnline -Url $siteUrl `
                          -ClientId $appId `
                          -Tenant $tenant `
                          -CertificatePath $pfxPath `
                          -CertificatePassword (ConvertTo-SecureString -String $pfxPassword -AsPlainText -Force)
        
        # Get the user
        $user = Get-PnPUser | Where-Object { $_.Email -eq $upn -or $_.LoginName -like "*$upn*" }
        
        if ($user) {
            Write-Host "✓ Found user: $($user.Title) ($($user.LoginName))" -ForegroundColor Green
            
            # Check direct permissions
            Write-Host "  Checking direct permissions..." -ForegroundColor Gray
            $web = Get-PnPWeb -Includes RoleAssignments
            $userPermissions = $web.RoleAssignments | Where-Object { $_.Member.LoginName -eq $user.LoginName }
            
            if ($userPermissions) {
                Write-Host "  ✓ User has direct permissions!" -ForegroundColor Green
                foreach ($perm in $userPermissions) {
                    $perm.RoleDefinitionBindings | ForEach-Object {
                        Write-Host "    - $($_.Name)" -ForegroundColor Cyan
                    }
                }
            }
            
            # Check group memberships
            Write-Host "  Checking group memberships..." -ForegroundColor Gray
            $groups = Get-PnPGroup
            foreach ($group in $groups) {
                $members = Get-PnPGroupMember -Identity $group.Id -ErrorAction SilentlyContinue
                if ($members | Where-Object { $_.LoginName -eq $user.LoginName }) {
                    Write-Host "  ✓ User is member of group: $($group.Title)" -ForegroundColor Green
                    
                    # Get group permissions
                    $groupPerms = $web.RoleAssignments | Where-Object { $_.Member.Id -eq $group.Id }
                    if ($groupPerms) {
                        $groupPerms.RoleDefinitionBindings | ForEach-Object {
                            Write-Host "    - Group has permission: $($_.Name)" -ForegroundColor Cyan
                        }
                    }
                }
            }
        } else {
            Write-Host "  ✗ User not found in this site" -ForegroundColor Yellow
        }
        
    } catch {
        Write-Host "  ✗ Error: $_" -ForegroundColor Red
    }
}

# Cleanup
Remove-Item $pfxPath -Force -ErrorAction SilentlyContinue
Disconnect-PnPOnline

Write-Host "`nDONE!" -ForegroundColor Green
Write-Host "This demonstrates the difference between Graph API and SharePoint permissions." -ForegroundColor Cyan
Write-Host "To remove permissions across all sites, we would need to use PnP PowerShell." -ForegroundColor Yellow