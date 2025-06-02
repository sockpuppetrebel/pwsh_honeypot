#Requires -Modules Microsoft.Graph.DeviceManagement

<#
.SYNOPSIS
    Retrieves Intune device compliance status and generates detailed reports
.DESCRIPTION
    Connects to Microsoft Graph and pulls compliance data for all managed devices,
    including policy details and remediation suggestions
.PARAMETER TenantId
    Azure AD Tenant ID
.PARAMETER OutputPath
    Path for CSV export (optional)
.EXAMPLE
    .\Get-IntuneDeviceCompliance.ps1 -TenantId "12345678-1234-1234-1234-123456789012"
#>

param(
    [Parameter(Mandatory = $true)]
    [string]$TenantId,
    
    [Parameter(Mandatory = $false)]
    [string]$OutputPath = ".\compliance-report-$(Get-Date -Format 'yyyy-MM-dd').csv"
)

# Connect to Microsoft Graph
try {
    Connect-MgGraph -TenantId $TenantId -Scopes "DeviceManagementManagedDevices.Read.All", "DeviceManagementConfiguration.Read.All"
    Write-Host "Connected to Microsoft Graph successfully" -ForegroundColor Green
}
catch {
    Write-Error "Failed to connect to Microsoft Graph: $($_.Exception.Message)"
    exit 1
}

# Get all managed devices
Write-Host "Retrieving managed devices..." -ForegroundColor Yellow
$devices = Get-MgDeviceManagementManagedDevice -All

# Get compliance policies
Write-Host "Retrieving compliance policies..." -ForegroundColor Yellow
$compliancePolicies = Get-MgDeviceManagementDeviceCompliancePolicy -All

$results = @()

foreach ($device in $devices) {
    Write-Progress -Activity "Processing devices" -Status $device.DeviceName -PercentComplete (($devices.IndexOf($device) / $devices.Count) * 100)
    
    # Get device compliance status
    $complianceStatus = Get-MgDeviceManagementManagedDeviceComplianceState -ManagedDeviceId $device.Id
    
    $deviceInfo = [PSCustomObject]@{
        DeviceName = $device.DeviceName
        UserPrincipalName = $device.UserPrincipalName
        DeviceType = $device.DeviceType
        OperatingSystem = $device.OperatingSystem
        OSVersion = $device.OsVersion
        ComplianceState = $device.ComplianceState
        LastSyncDateTime = $device.LastSyncDateTime
        EnrolledDateTime = $device.EnrolledDateTime
        ManagementAgent = $device.ManagementAgent
        DeviceId = $device.Id
        SerialNumber = $device.SerialNumber
        Model = $device.Model
        Manufacturer = $device.Manufacturer
        IsEncrypted = $device.IsEncrypted
        JailBroken = $device.JailBroken
        PartnerReportedThreatState = $device.PartnerReportedThreatState
    }
    
    $results += $deviceInfo
}

# Generate summary statistics
$totalDevices = $results.Count
$compliantDevices = ($results | Where-Object ComplianceState -eq "Compliant").Count
$nonCompliantDevices = ($results | Where-Object ComplianceState -eq "NonCompliant").Count
$unknownDevices = ($results | Where-Object ComplianceState -eq "Unknown").Count

Write-Host "`n=== COMPLIANCE SUMMARY ===" -ForegroundColor Cyan
Write-Host "Total Devices: $totalDevices" -ForegroundColor White
Write-Host "Compliant: $compliantDevices ($([math]::Round(($compliantDevices/$totalDevices)*100,2))%)" -ForegroundColor Green
Write-Host "Non-Compliant: $nonCompliantDevices ($([math]::Round(($nonCompliantDevices/$totalDevices)*100,2))%)" -ForegroundColor Red
Write-Host "Unknown: $unknownDevices ($([math]::Round(($unknownDevices/$totalDevices)*100,2))%)" -ForegroundColor Yellow

# Show devices that haven't synced in 7+ days
$staleDevices = $results | Where-Object { 
    $_.LastSyncDateTime -and 
    (Get-Date $_.LastSyncDateTime) -lt (Get-Date).AddDays(-7) 
}

if ($staleDevices) {
    Write-Host "`n=== STALE DEVICES (No sync in 7+ days) ===" -ForegroundColor Yellow
    $staleDevices | Select-Object DeviceName, UserPrincipalName, LastSyncDateTime | Format-Table -AutoSize
}

# Export to CSV
$results | Export-Csv -Path $OutputPath -NoTypeInformation
Write-Host "`nReport exported to: $OutputPath" -ForegroundColor Green

# Disconnect
Disconnect-MgGraph
Write-Host "Disconnected from Microsoft Graph" -ForegroundColor Green