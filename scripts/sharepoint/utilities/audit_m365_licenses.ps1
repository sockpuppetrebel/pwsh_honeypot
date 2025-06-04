#Requires -Modules Microsoft.Graph, Microsoft.Graph.Users, Microsoft.Graph.Reports

<#
.SYNOPSIS
    Comprehensive M365 License Audit for Optimizely Tenant
.DESCRIPTION
    This script audits all M365 licenses in your tenant and exports detailed reports
.NOTES
    Requires Microsoft Graph PowerShell modules
    Install with: Install-Module Microsoft.Graph -Scope CurrentUser
#>

# Configuration
$TenantName = "episerver99"
$OutputPath = "C:\Temp\M365_License_Audit_$(Get-Date -Format 'yyyy-MM-dd_HHmmss')"
$ExportFormat = "CSV" # Can be CSV or Excel

# Create output directory
if (!(Test-Path $OutputPath)) {
    New-Item -ItemType Directory -Path $OutputPath | Out-Null
}

Write-Host "=== M365 LICENSE AUDIT ===" -ForegroundColor Cyan
Write-Host "Tenant: $TenantName" -ForegroundColor Yellow
Write-Host "Output: $OutputPath" -ForegroundColor Yellow
Write-Host ""

# Connect to Microsoft Graph
Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Green
Connect-MgGraph -Scopes "User.Read.All", "Directory.Read.All", "Reports.Read.All"

# Get subscription SKUs (available licenses)
Write-Host "Fetching subscription information..." -ForegroundColor Green
$Skus = Get-MgSubscribedSku

# Create SKU summary
$SkuSummary = @()
foreach ($Sku in $Skus) {
    $SkuSummary += [PSCustomObject]@{
        SkuPartNumber = $Sku.SkuPartNumber
        SkuId = $Sku.SkuId
        DisplayName = $Sku.SkuPartNumber
        TotalLicenses = $Sku.PrepaidUnits.Enabled
        ConsumedLicenses = $Sku.ConsumedUnits
        AvailableLicenses = $Sku.PrepaidUnits.Enabled - $Sku.ConsumedUnits
        SuspendedLicenses = $Sku.PrepaidUnits.Suspended
        WarningLicenses = $Sku.PrepaidUnits.Warning
    }
}

# Export SKU summary
$SkuSummary | Export-Csv -Path "$OutputPath\License_Summary.csv" -NoTypeInformation
Write-Host "Exported license summary to: $OutputPath\License_Summary.csv" -ForegroundColor Green

# Get all users with licenses
Write-Host "Fetching all licensed users..." -ForegroundColor Green
$Users = Get-MgUser -All -Property Id,DisplayName,UserPrincipalName,AssignedLicenses,AssignedPlans,Department,JobTitle,CreatedDateTime,AccountEnabled

# Create detailed user license report
$UserLicenseDetails = @()
$counter = 0
$totalUsers = $Users.Count

foreach ($User in $Users) {
    $counter++
    Write-Progress -Activity "Processing Users" -Status "User $counter of $totalUsers" -PercentComplete (($counter / $totalUsers) * 100)
    
    if ($User.AssignedLicenses.Count -gt 0) {
        # Get license names for this user
        $LicenseNames = @()
        foreach ($License in $User.AssignedLicenses) {
            $SkuName = ($Skus | Where-Object { $_.SkuId -eq $License.SkuId }).SkuPartNumber
            if ($SkuName) {
                $LicenseNames += $SkuName
            }
        }
        
        $UserLicenseDetails += [PSCustomObject]@{
            DisplayName = $User.DisplayName
            UserPrincipalName = $User.UserPrincipalName
            Department = $User.Department
            JobTitle = $User.JobTitle
            AccountEnabled = $User.AccountEnabled
            CreatedDate = $User.CreatedDateTime
            Licenses = $LicenseNames -join "; "
            LicenseCount = $User.AssignedLicenses.Count
        }
    }
}

Write-Progress -Activity "Processing Users" -Completed

# Export user details
$UserLicenseDetails | Export-Csv -Path "$OutputPath\User_License_Details.csv" -NoTypeInformation
Write-Host "Exported user details to: $OutputPath\User_License_Details.csv" -ForegroundColor Green

# Create license assignment report (who has what)
Write-Host "Creating license assignment matrix..." -ForegroundColor Green
$LicenseMatrix = @()

foreach ($Sku in $Skus) {
    if ($Sku.ConsumedUnits -gt 0) {
        $UsersWithLicense = $Users | Where-Object { 
            $_.AssignedLicenses.SkuId -contains $Sku.SkuId 
        }
        
        foreach ($User in $UsersWithLicense) {
            $LicenseMatrix += [PSCustomObject]@{
                License = $Sku.SkuPartNumber
                UserPrincipalName = $User.UserPrincipalName
                DisplayName = $User.DisplayName
                Department = $User.Department
                AccountEnabled = $User.AccountEnabled
            }
        }
    }
}

$LicenseMatrix | Export-Csv -Path "$OutputPath\License_Assignment_Matrix.csv" -NoTypeInformation
Write-Host "Exported license matrix to: $OutputPath\License_Assignment_Matrix.csv" -ForegroundColor Green

# Create department summary
Write-Host "Creating department summary..." -ForegroundColor Green
$DeptSummary = $UserLicenseDetails | Group-Object Department | ForEach-Object {
    [PSCustomObject]@{
        Department = if ($_.Name) { $_.Name } else { "(No Department)" }
        UserCount = $_.Count
        TotalLicenses = ($_.Group | Measure-Object -Property LicenseCount -Sum).Sum
    }
} | Sort-Object UserCount -Descending

$DeptSummary | Export-Csv -Path "$OutputPath\Department_License_Summary.csv" -NoTypeInformation
Write-Host "Exported department summary to: $OutputPath\Department_License_Summary.csv" -ForegroundColor Green

# Find users with multiple licenses
$MultiLicenseUsers = $UserLicenseDetails | Where-Object { $_.LicenseCount -gt 1 } | Sort-Object LicenseCount -Descending
if ($MultiLicenseUsers) {
    $MultiLicenseUsers | Export-Csv -Path "$OutputPath\Users_With_Multiple_Licenses.csv" -NoTypeInformation
    Write-Host "Exported multi-license users to: $OutputPath\Users_With_Multiple_Licenses.csv" -ForegroundColor Green
}

# Find disabled users with licenses (potential savings)
$DisabledWithLicenses = $UserLicenseDetails | Where-Object { $_.AccountEnabled -eq $false }
if ($DisabledWithLicenses) {
    $DisabledWithLicenses | Export-Csv -Path "$OutputPath\Disabled_Users_With_Licenses.csv" -NoTypeInformation
    Write-Host "Exported disabled users with licenses to: $OutputPath\Disabled_Users_With_Licenses.csv" -ForegroundColor Green
}

# Create summary report
$SummaryReport = @"
M365 LICENSE AUDIT SUMMARY
==========================
Generated: $(Get-Date)
Tenant: $TenantName

OVERALL STATISTICS
------------------
Total License Types: $($Skus.Count)
Total Licensed Users: $($UserLicenseDetails.Count)
Total Licenses Consumed: $(($Skus | Measure-Object -Property ConsumedUnits -Sum).Sum)
Total Licenses Available: $(($Skus | ForEach-Object { $_.PrepaidUnits.Enabled } | Measure-Object -Sum).Sum)

POTENTIAL SAVINGS
-----------------
Disabled Users with Licenses: $($DisabledWithLicenses.Count)
Estimated Monthly Savings: Calculate based on license costs

TOP LICENSE USAGE
-----------------
$($SkuSummary | Sort-Object ConsumedLicenses -Descending | Select-Object -First 5 | Format-Table -AutoSize | Out-String)

Files Generated:
- License_Summary.csv
- User_License_Details.csv
- License_Assignment_Matrix.csv
- Department_License_Summary.csv
$(if ($MultiLicenseUsers) { "- Users_With_Multiple_Licenses.csv" })
$(if ($DisabledWithLicenses) { "- Disabled_Users_With_Licenses.csv" })
"@

$SummaryReport | Out-File -Path "$OutputPath\AUDIT_SUMMARY.txt"
Write-Host ""
Write-Host $SummaryReport -ForegroundColor Cyan

Write-Host ""
Write-Host "=== AUDIT COMPLETE ===" -ForegroundColor Green
Write-Host "All reports saved to: $OutputPath" -ForegroundColor Green
Write-Host ""

# Disconnect
Disconnect-MgGraph