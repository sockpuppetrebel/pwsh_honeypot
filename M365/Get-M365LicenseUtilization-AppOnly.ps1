#Requires -Modules Microsoft.Graph.Users, Microsoft.Graph.Identity.DirectoryManagement

<#
.SYNOPSIS
    Analyzes M365 license utilization using app-only certificate authentication
.DESCRIPTION
    Uses certificate-based app-only authentication for unattended scenarios
#>

param(
    [Parameter(Mandatory = $false)]
    [switch]$ExportToExcel,
    
    [Parameter(Mandatory = $false)]
    [switch]$ShowInactiveUsers,
    
    [Parameter(Mandatory = $false)]
    [int]$InactiveDays = 90
)

# App registration details
$tenantId = '3ec00d79-021a-42d4-aac8-dcb35973dff2'
$clientId = 'fe2a9efe-3000-4b02-96ea-344a2583dd52'
$certThumbprint = '5C27245587D066005B3F9B5098AD5A3BA22E69CC'  # Original thumbprint from your script

Write-Host "Connecting to Microsoft Graph using app-only authentication..." -ForegroundColor Yellow

# For macOS: Export cert from Keychain and use it
$certPath = "$HOME/.certs/optimizely-graph.pfx"

# Check if we need to export the certificate
if (-not (Test-Path $certPath)) {
    Write-Host "Certificate not found at $certPath" -ForegroundColor Yellow
    Write-Host "`nTo use certificate authentication on macOS:" -ForegroundColor Cyan
    Write-Host "1. Open Keychain Access" -ForegroundColor White
    Write-Host "2. Find your certificate (first.last@optimizely.com)" -ForegroundColor White
    Write-Host "3. Export it as a .p12/.pfx file" -ForegroundColor White
    Write-Host "4. Save it to: $certPath" -ForegroundColor White
    Write-Host "5. Run this script again" -ForegroundColor White
    Write-Host "`nAlternatively, use the interactive version: ./Get-M365LicenseUtilization.ps1" -ForegroundColor Yellow
    exit 1
}

# Load certificate from file
try {
    $certPassword = Read-Host -AsSecureString "Enter certificate password"
    $cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($certPath, $certPassword)
    
    # Connect using certificate
    Connect-MgGraph -ClientId $clientId -TenantId $tenantId -Certificate $cert -NoWelcome
    
    # Verify connection
    $context = Get-MgContext
    if ($context) {
        Write-Host "Connected successfully using certificate" -ForegroundColor Green
        Write-Host "App: $($context.AppName)" -ForegroundColor White
        Write-Host "TenantId: $($context.TenantId)" -ForegroundColor White
    } else {
        throw "Connection verification failed"
    }
} catch {
    Write-Error "Failed to connect: $_"
    exit 1
}

Write-Host "=== M365 LICENSE UTILIZATION ANALYSIS ===" -ForegroundColor Cyan

# Get all subscribed SKUs
Write-Host "Retrieving license information..." -ForegroundColor Yellow
$subscribedSkus = Get-MgSubscribedSku -All

# Get all users with license assignments
Write-Host "Retrieving user license assignments..." -ForegroundColor Yellow
$users = Get-MgUser -All -Property "Id,UserPrincipalName,DisplayName,AssignedLicenses,AccountEnabled,SignInActivity,CreatedDateTime,UserType" -Filter "assignedLicenses/`$count ne 0"

# License SKU mapping for friendly names
$skuMapping = @{
    "ENTERPRISEPACK" = "Office 365 E3"
    "ENTERPRISEPREMIUM" = "Office 365 E5"
    "ENTERPRISEPACK_GOV" = "Office 365 E3 (Government)"
    "ENTERPRISEPREMIUM_GOV" = "Office 365 E5 (Government)"
    "POWER_BI_PRO" = "Power BI Pro"
    "FLOW_FREE" = "Power Automate (Free)"
    "TEAMS_EXPLORATORY" = "Microsoft Teams Exploratory"
    "MICROSOFT_365_E3" = "Microsoft 365 E3"
    "MICROSOFT_365_E5" = "Microsoft 365 E5"
    "MICROSOFT_365_F1" = "Microsoft 365 F1"
    "MICROSOFT_365_F3" = "Microsoft 365 F3"
    "SPB" = "Microsoft 365 Business Premium"
    "SPE_E3" = "Microsoft 365 E3"
    "SPE_E5" = "Microsoft 365 E5"
    "VISIOCLIENT" = "Visio Plan 2"
    "PROJECTPREMIUM" = "Project Plan 5"
    "EMS" = "Enterprise Mobility + Security E3"
    "EMSPREMIUM" = "Enterprise Mobility + Security E5"
}

# Analyze license utilization
$licenseAnalysis = @()

foreach ($sku in $subscribedSkus) {
    $friendlyName = if ($skuMapping.ContainsKey($sku.SkuPartNumber)) { 
        $skuMapping[$sku.SkuPartNumber] 
    } else { 
        $sku.SkuPartNumber 
    }
    
    $assignedUsers = $users | Where-Object { 
        $_.AssignedLicenses.SkuId -contains $sku.SkuId 
    }
    
    $enabledUsers = $assignedUsers | Where-Object AccountEnabled -eq $true
    $disabledUsers = $assignedUsers | Where-Object AccountEnabled -eq $false
    
    # Calculate inactive users if requested
    $inactiveUsers = @()
    if ($ShowInactiveUsers) {
        $cutoffDate = (Get-Date).AddDays(-$InactiveDays)
        $inactiveUsers = $enabledUsers | Where-Object {
            if ($_.SignInActivity.LastSignInDateTime) {
                (Get-Date $_.SignInActivity.LastSignInDateTime) -lt $cutoffDate
            } else {
                $true  # Never signed in
            }
        }
    }
    
    $utilizationPercent = if ($sku.PrepaidUnits.Enabled -gt 0) {
        [math]::Round(($sku.ConsumedUnits / $sku.PrepaidUnits.Enabled) * 100, 2)
    } else {
        0
    }
    
    $analysis = [PSCustomObject]@{
        LicenseName = $friendlyName
        SkuPartNumber = $sku.SkuPartNumber
        SkuId = $sku.SkuId
        TotalLicenses = $sku.PrepaidUnits.Enabled
        ConsumedLicenses = $sku.ConsumedUnits
        AvailableLicenses = $sku.PrepaidUnits.Enabled - $sku.ConsumedUnits
        UtilizationPercent = $utilizationPercent
        AssignedToEnabledUsers = $enabledUsers.Count
        AssignedToDisabledUsers = $disabledUsers.Count
        InactiveUsers = $inactiveUsers.Count
        SuspendedLicenses = $sku.PrepaidUnits.Suspended
        WarningLicenses = $sku.PrepaidUnits.Warning
        CapabilityStatus = $sku.CapabilityStatus
        ServicePlans = ($sku.ServicePlans | Where-Object ProvisioningStatus -eq "Success").Count
    }
    
    $licenseAnalysis += $analysis
}

# Display summary
Write-Host "`n--- LICENSE UTILIZATION SUMMARY ---" -ForegroundColor Cyan

$totalLicenses = ($licenseAnalysis | Measure-Object TotalLicenses -Sum).Sum
$totalConsumed = ($licenseAnalysis | Measure-Object ConsumedLicenses -Sum).Sum
$totalAvailable = ($licenseAnalysis | Measure-Object AvailableLicenses -Sum).Sum

Write-Host "Total Licenses: $totalLicenses" -ForegroundColor White
Write-Host "Total Consumed: $totalConsumed" -ForegroundColor White
Write-Host "Total Available: $totalAvailable" -ForegroundColor White
Write-Host "Overall Utilization: $([math]::Round(($totalConsumed / $totalLicenses) * 100, 2))%" -ForegroundColor White

# Show licenses with low utilization
$lowUtilization = $licenseAnalysis | Where-Object { $_.UtilizationPercent -lt 50 -and $_.TotalLicenses -gt 0 }
if ($lowUtilization) {
    Write-Host "`n--- LOW UTILIZATION LICENSES ---" -ForegroundColor Yellow
    $lowUtilization | Select-Object LicenseName, TotalLicenses, ConsumedLicenses, UtilizationPercent | 
        Sort-Object UtilizationPercent | Format-Table -AutoSize
}

# Show licenses nearing capacity
$nearCapacity = $licenseAnalysis | Where-Object UtilizationPercent -gt 90
if ($nearCapacity) {
    Write-Host "`n--- LICENSES NEARING CAPACITY ---" -ForegroundColor Red
    $nearCapacity | Select-Object LicenseName, TotalLicenses, ConsumedLicenses, AvailableLicenses, UtilizationPercent | 
        Format-Table -AutoSize
}

# Analyze disabled user licenses
$disabledUserLicenses = $licenseAnalysis | Where-Object AssignedToDisabledUsers -gt 0
if ($disabledUserLicenses) {
    Write-Host "`n--- LICENSES ASSIGNED TO DISABLED USERS ---" -ForegroundColor Yellow
    $disabledUserLicenses | Select-Object LicenseName, AssignedToDisabledUsers | 
        Sort-Object AssignedToDisabledUsers -Descending | Format-Table -AutoSize
    
    $totalDisabledLicenses = ($disabledUserLicenses | Measure-Object AssignedToDisabledUsers -Sum).Sum
    Write-Host "Total licenses assigned to disabled users: $totalDisabledLicenses" -ForegroundColor Yellow
}

# Analyze inactive users if requested
if ($ShowInactiveUsers) {
    $inactiveUserLicenses = $licenseAnalysis | Where-Object InactiveUsers -gt 0
    if ($inactiveUserLicenses) {
        Write-Host "`n--- LICENSES ASSIGNED TO INACTIVE USERS (>$InactiveDays days) ---" -ForegroundColor Yellow
        $inactiveUserLicenses | Select-Object LicenseName, InactiveUsers | 
            Sort-Object InactiveUsers -Descending | Format-Table -AutoSize
        
        $totalInactiveLicenses = ($inactiveUserLicenses | Measure-Object InactiveUsers -Sum).Sum
        Write-Host "Total licenses assigned to inactive users: $totalInactiveLicenses" -ForegroundColor Yellow
    }
}

# Service plan analysis
Write-Host "`n--- SERVICE PLAN ANALYSIS ---" -ForegroundColor Cyan

$servicePlanData = @()
foreach ($sku in $subscribedSkus) {
    foreach ($servicePlan in $sku.ServicePlans) {
        $servicePlanData += [PSCustomObject]@{
            LicenseName = if ($skuMapping.ContainsKey($sku.SkuPartNumber)) { $skuMapping[$sku.SkuPartNumber] } else { $sku.SkuPartNumber }
            ServicePlanName = $servicePlan.ServicePlanName
            ServicePlanId = $servicePlan.ServicePlanId
            ProvisioningStatus = $servicePlan.ProvisioningStatus
            AppliesTo = $servicePlan.AppliesTo
        }
    }
}

# Find disabled service plans
$disabledServices = $servicePlanData | Where-Object ProvisioningStatus -ne "Success" | Group-Object ServicePlanName
if ($disabledServices) {
    Write-Host "Service plans with issues:" -ForegroundColor Yellow
    $disabledServices | Select-Object Name, Count | Sort-Object Count -Descending | Format-Table -AutoSize
}

# Cost optimization recommendations
Write-Host "`n--- COST OPTIMIZATION RECOMMENDATIONS ---" -ForegroundColor Cyan

$recommendations = @()

# Unused licenses
$unusedLicenses = $licenseAnalysis | Where-Object { $_.AvailableLicenses -gt 0 -and $_.UtilizationPercent -lt 80 }
foreach ($license in $unusedLicenses) {
    $recommendations += "Consider reducing $($license.LicenseName) by $($license.AvailableLicenses) licenses"
}

# Disabled user licenses
if ($totalDisabledLicenses -gt 0) {
    $recommendations += "Remove $totalDisabledLicenses licenses from disabled users"
}

# Inactive user licenses
if ($ShowInactiveUsers -and $totalInactiveLicenses -gt 0) {
    $recommendations += "Review $totalInactiveLicenses licenses assigned to inactive users"
}

# Display recommendations
if ($recommendations.Count -gt 0) {
    $recommendations | ForEach-Object { Write-Host "• $_" -ForegroundColor Cyan }
} else {
    Write-Host "• License utilization appears optimized" -ForegroundColor Green
}

# Export data if requested
if ($ExportToExcel) {
    $timestamp = Get-Date -Format "yyyy-MM-dd-HHmm"
    $outputPath = ".\M365-License-Analysis-$timestamp.xlsx"
    
    if (Get-Module -ListAvailable -Name ImportExcel) {
        # Main analysis
        $licenseAnalysis | Export-Excel -Path $outputPath -WorksheetName "License Analysis" -AutoSize -FreezeTopRow
        
        # User details
        $userDetails = foreach ($user in $users) {
            $userLicenses = foreach ($license in $user.AssignedLicenses) {
                $sku = $subscribedSkus | Where-Object SkuId -eq $license.SkuId
                if ($skuMapping.ContainsKey($sku.SkuPartNumber)) { 
                    $skuMapping[$sku.SkuPartNumber] 
                } else { 
                    $sku.SkuPartNumber 
                }
            }
            
            [PSCustomObject]@{
                UserPrincipalName = $user.UserPrincipalName
                DisplayName = $user.DisplayName
                AccountEnabled = $user.AccountEnabled
                UserType = $user.UserType
                CreatedDateTime = $user.CreatedDateTime
                LastSignIn = $user.SignInActivity.LastSignInDateTime
                DaysSinceLastSignIn = if ($user.SignInActivity.LastSignInDateTime) {
                    (Get-Date) - (Get-Date $user.SignInActivity.LastSignInDateTime) | Select-Object -ExpandProperty Days
                } else {
                    "Never"
                }
                AssignedLicenses = ($userLicenses -join "; ")
                LicenseCount = $user.AssignedLicenses.Count
            }
        }
        
        $userDetails | Export-Excel -Path $outputPath -WorksheetName "User License Details" -AutoSize -FreezeTopRow
        
        # Service plans
        $servicePlanData | Export-Excel -Path $outputPath -WorksheetName "Service Plans" -AutoSize -FreezeTopRow
        
        Write-Host "`nDetailed report exported to: $outputPath" -ForegroundColor Green
    } else {
        Write-Host "ImportExcel module not available. Exporting to CSV instead..." -ForegroundColor Yellow
        $licenseAnalysis | Export-Csv -Path ".\M365-License-Analysis-$timestamp.csv" -NoTypeInformation
        Write-Host "Report exported to: .\M365-License-Analysis-$timestamp.csv" -ForegroundColor Green
    }
}

Write-Host "`n--- NEXT STEPS ---" -ForegroundColor Cyan
Write-Host "• Review disabled user accounts and remove unnecessary licenses" -ForegroundColor White
Write-Host "• Monitor usage patterns for inactive users" -ForegroundColor White
Write-Host "• Consider license pooling or auto-assignment for future growth" -ForegroundColor White
Write-Host "• Set up automated alerts for license capacity thresholds" -ForegroundColor White

Disconnect-MgGraph
Write-Host "Disconnected from Microsoft Graph" -ForegroundColor Green