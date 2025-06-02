#Requires -Modules Microsoft.Graph.DeviceManagement

<#
.SYNOPSIS
    Comprehensive Intune device inventory export for reporting and analysis
.DESCRIPTION
    Exports detailed device information including hardware specs, compliance status,
    installed apps, and policy assignments. Useful for asset management and auditing.
.PARAMETER OutputFormat
    Export format: CSV, JSON, or Excel (requires ImportExcel module)
.PARAMETER IncludeApps
    Include installed applications data (increases runtime)
.PARAMETER FilterByOS
    Filter devices by operating system (Windows, iOS, Android, macOS)
.EXAMPLE
    .\Export-IntuneInventory.ps1 -OutputFormat Excel -IncludeApps -FilterByOS Windows
#>

param(
    [Parameter(Mandatory = $false)]
    [ValidateSet("CSV", "JSON", "Excel")]
    [string]$OutputFormat = "CSV",
    
    [Parameter(Mandatory = $false)]
    [switch]$IncludeApps,
    
    [Parameter(Mandatory = $false)]
    [ValidateSet("Windows", "iOS", "Android", "macOS", "All")]
    [string]$FilterByOS = "All"
)

# Connect to Microsoft Graph
try {
    $scopes = @(
        "DeviceManagementManagedDevices.Read.All",
        "DeviceManagementConfiguration.Read.All",
        "DeviceManagementApps.Read.All"
    )
    
    Connect-MgGraph -Scopes $scopes
    Write-Host "Connected to Microsoft Graph" -ForegroundColor Green
}
catch {
    Write-Error "Failed to connect to Microsoft Graph: $($_.Exception.Message)"
    exit 1
}

Write-Host "=== INTUNE DEVICE INVENTORY EXPORT ===" -ForegroundColor Cyan

# Get all managed devices
Write-Host "Retrieving managed devices..." -ForegroundColor Yellow
$devices = Get-MgDeviceManagementManagedDevice -All

# Apply OS filter
if ($FilterByOS -ne "All") {
    $devices = $devices | Where-Object OperatingSystem -like "*$FilterByOS*"
    Write-Host "Filtered to $($devices.Count) $FilterByOS devices" -ForegroundColor Yellow
}

if ($devices.Count -eq 0) {
    Write-Host "No devices found matching criteria" -ForegroundColor Yellow
    Disconnect-MgGraph
    exit 0
}

Write-Host "Processing $($devices.Count) devices..." -ForegroundColor Yellow

# Get compliance policies for reference
$compliancePolicies = Get-MgDeviceManagementDeviceCompliancePolicy -All
$configurationPolicies = Get-MgDeviceManagementDeviceConfiguration -All

$inventory = @()
$deviceCounter = 0

foreach ($device in $devices) {
    $deviceCounter++
    Write-Progress -Activity "Processing devices" -Status "$($device.DeviceName) ($deviceCounter/$($devices.Count))" -PercentComplete (($deviceCounter / $devices.Count) * 100)
    
    # Basic device information
    $deviceInfo = [PSCustomObject]@{
        DeviceName = $device.DeviceName
        UserPrincipalName = $device.UserPrincipalName
        UserDisplayName = $device.UserDisplayName
        DeviceId = $device.Id
        AzureADDeviceId = $device.AzureAdDeviceId
        SerialNumber = $device.SerialNumber
        IMEI = $device.Imei
        WiFiMacAddress = $device.WiFiMacAddress
        EthernetMacAddress = $device.EthernetMacAddress
        
        # Operating System Info
        OperatingSystem = $device.OperatingSystem
        OSVersion = $device.OsVersion
        AndroidSecurityPatchLevel = $device.AndroidSecurityPatchLevel
        
        # Hardware Info
        Manufacturer = $device.Manufacturer
        Model = $device.Model
        TotalStorageSpaceInBytes = $device.TotalStorageSpaceInBytes
        FreeStorageSpaceInBytes = $device.FreeStorageSpaceInBytes
        PhysicalMemoryInBytes = $device.PhysicalMemoryInBytes
        
        # Management Info
        EnrolledDateTime = $device.EnrolledDateTime
        LastSyncDateTime = $device.LastSyncDateTime
        ComplianceState = $device.ComplianceState
        ManagementState = $device.ManagementState
        ManagementAgent = $device.ManagementAgent
        IsEncrypted = $device.IsEncrypted
        IsSupervised = $device.IsSupervised
        JailBroken = $device.JailBroken
        
        # Security Info
        PartnerReportedThreatState = $device.PartnerReportedThreatState
        RequireUserEnrollmentApproval = $device.RequireUserEnrollmentApproval
        ActivationLockBypassCode = $device.ActivationLockBypassCode
        
        # Additional Properties
        DeviceType = $device.DeviceType
        DeviceRegistrationState = $device.DeviceRegistrationState
        ExchangeAccessState = $device.ExchangeAccessState
        ExchangeAccessStateReason = $device.ExchangeAccessStateReason
        RemoteAssistanceSessionErrorDetails = $device.RemoteAssistanceSessionErrorDetails
        
        # Calculated Fields
        StorageUsedPercent = if ($device.TotalStorageSpaceInBytes -gt 0) { 
            [math]::Round((($device.TotalStorageSpaceInBytes - $device.FreeStorageSpaceInBytes) / $device.TotalStorageSpaceInBytes) * 100, 2) 
        } else { 
            $null 
        }
        DaysSinceLastSync = if ($device.LastSyncDateTime) { 
            (Get-Date) - (Get-Date $device.LastSyncDateTime) | Select-Object -ExpandProperty Days 
        } else { 
            $null 
        }
        DaysSinceEnrollment = if ($device.EnrolledDateTime) { 
            (Get-Date) - (Get-Date $device.EnrolledDateTime) | Select-Object -ExpandProperty Days 
        } else { 
            $null 
        }
    }
    
    # Get device compliance policy assignments
    try {
        $complianceDetails = Get-MgDeviceManagementManagedDeviceComplianceState -ManagedDeviceId $device.Id -ErrorAction SilentlyContinue
        if ($complianceDetails) {
            $deviceInfo | Add-Member -NotePropertyName "CompliancePolicyCount" -NotePropertyValue $complianceDetails.Count
        }
    }
    catch {
        # Compliance details not available
    }
    
    # Get installed applications if requested
    if ($IncludeApps) {
        try {
            $installedApps = Get-MgDeviceManagementManagedDeviceDetectedApp -ManagedDeviceId $device.Id -ErrorAction SilentlyContinue
            $appNames = ($installedApps | Select-Object -ExpandProperty DisplayName | Sort-Object -Unique) -join "; "
            $deviceInfo | Add-Member -NotePropertyName "InstalledApps" -NotePropertyValue $appNames
            $deviceInfo | Add-Member -NotePropertyName "InstalledAppCount" -NotePropertyValue $installedApps.Count
        }
        catch {
            $deviceInfo | Add-Member -NotePropertyName "InstalledApps" -NotePropertyValue "Error retrieving apps"
            $deviceInfo | Add-Member -NotePropertyName "InstalledAppCount" -NotePropertyValue 0
        }
    }
    
    $inventory += $deviceInfo
}

Write-Progress -Activity "Processing devices" -Completed

# Generate summary statistics
Write-Host "`n--- INVENTORY SUMMARY ---" -ForegroundColor Cyan

$totalDevices = $inventory.Count
$compliantDevices = ($inventory | Where-Object ComplianceState -eq "Compliant").Count
$nonCompliantDevices = ($inventory | Where-Object ComplianceState -eq "NonCompliant").Count
$unknownCompliance = ($inventory | Where-Object ComplianceState -eq "Unknown").Count

Write-Host "Total Devices: $totalDevices" -ForegroundColor White
Write-Host "Compliant: $compliantDevices ($([math]::Round(($compliantDevices/$totalDevices)*100,1))%)" -ForegroundColor Green
Write-Host "Non-Compliant: $nonCompliantDevices ($([math]::Round(($nonCompliantDevices/$totalDevices)*100,1))%)" -ForegroundColor Red
Write-Host "Unknown: $unknownCompliance ($([math]::Round(($unknownCompliance/$totalDevices)*100,1))%)" -ForegroundColor Yellow

# OS Distribution
Write-Host "`nOperating System Distribution:" -ForegroundColor Yellow
$inventory | Group-Object OperatingSystem | Sort-Object Count -Descending | ForEach-Object {
    $percentage = [math]::Round(($_.Count / $totalDevices) * 100, 1)
    Write-Host "  $($_.Name): $($_.Count) ($percentage%)" -ForegroundColor White
}

# Manufacturer Distribution
Write-Host "`nManufacturer Distribution:" -ForegroundColor Yellow
$inventory | Group-Object Manufacturer | Sort-Object Count -Descending | Select-Object -First 5 | ForEach-Object {
    $percentage = [math]::Round(($_.Count / $totalDevices) * 100, 1)
    Write-Host "  $($_.Name): $($_.Count) ($percentage%)" -ForegroundColor White
}

# Devices with issues
$staleDevices = $inventory | Where-Object { $_.DaysSinceLastSync -gt 7 }
$lowStorageDevices = $inventory | Where-Object { $_.StorageUsedPercent -gt 90 }
$jailbrokenDevices = $inventory | Where-Object JailBroken -eq "True"

if ($staleDevices) {
    Write-Host "`nStale Devices (>7 days since sync): $($staleDevices.Count)" -ForegroundColor Yellow
}

if ($lowStorageDevices) {
    Write-Host "Low Storage Devices (>90% full): $($lowStorageDevices.Count)" -ForegroundColor Yellow
}

if ($jailbrokenDevices) {
    Write-Host "Jailbroken/Rooted Devices: $($jailbrokenDevices.Count)" -ForegroundColor Red
}

# Export data
$timestamp = Get-Date -Format "yyyy-MM-dd-HHmm"
$baseFileName = "Intune-Inventory-$FilterByOS-$timestamp"

switch ($OutputFormat) {
    "CSV" {
        $outputPath = ".\$baseFileName.csv"
        $inventory | Export-Csv -Path $outputPath -NoTypeInformation
        Write-Host "`nInventory exported to: $outputPath" -ForegroundColor Green
    }
    
    "JSON" {
        $outputPath = ".\$baseFileName.json"
        $exportData = @{
            GeneratedDate = Get-Date
            TotalDevices = $totalDevices
            FilterApplied = $FilterByOS
            IncludeApps = $IncludeApps.IsPresent
            Devices = $inventory
            Summary = @{
                Compliant = $compliantDevices
                NonCompliant = $nonCompliantDevices
                Unknown = $unknownCompliance
            }
        }
        $exportData | ConvertTo-Json -Depth 5 | Out-File -FilePath $outputPath
        Write-Host "`nInventory exported to: $outputPath" -ForegroundColor Green
    }
    
    "Excel" {
        if (Get-Module -ListAvailable -Name ImportExcel) {
            $outputPath = ".\$baseFileName.xlsx"
            
            # Main inventory sheet
            $inventory | Export-Excel -Path $outputPath -WorksheetName "Device Inventory" -AutoSize -FreezeTopRow
            
            # Summary sheet
            $summaryData = @(
                [PSCustomObject]@{ Metric = "Total Devices"; Value = $totalDevices }
                [PSCustomObject]@{ Metric = "Compliant"; Value = "$compliantDevices ($([math]::Round(($compliantDevices/$totalDevices)*100,1))%)" }
                [PSCustomObject]@{ Metric = "Non-Compliant"; Value = "$nonCompliantDevices ($([math]::Round(($nonCompliantDevices/$totalDevices)*100,1))%)" }
                [PSCustomObject]@{ Metric = "Unknown"; Value = "$unknownCompliance ($([math]::Round(($unknownCompliance/$totalDevices)*100,1))%)" }
                [PSCustomObject]@{ Metric = "Stale Devices"; Value = $staleDevices.Count }
                [PSCustomObject]@{ Metric = "Low Storage"; Value = $lowStorageDevices.Count }
                [PSCustomObject]@{ Metric = "Jailbroken"; Value = $jailbrokenDevices.Count }
            )
            
            $summaryData | Export-Excel -Path $outputPath -WorksheetName "Summary" -AutoSize
            
            # OS Distribution sheet
            $osData = $inventory | Group-Object OperatingSystem | Sort-Object Count -Descending | ForEach-Object {
                [PSCustomObject]@{
                    OperatingSystem = $_.Name
                    Count = $_.Count
                    Percentage = [math]::Round(($_.Count / $totalDevices) * 100, 1)
                }
            }
            $osData | Export-Excel -Path $outputPath -WorksheetName "OS Distribution" -AutoSize
            
            Write-Host "`nInventory exported to: $outputPath" -ForegroundColor Green
        } else {
            Write-Host "ImportExcel module not found. Installing..." -ForegroundColor Yellow
            Install-Module -Name ImportExcel -Force -Scope CurrentUser
            Write-Host "Please run the script again to export to Excel format" -ForegroundColor Yellow
        }
    }
}

# Generate quick analysis
Write-Host "`n--- QUICK ANALYSIS ---" -ForegroundColor Cyan

# Find devices that need attention
$needsAttention = $inventory | Where-Object {
    $_.ComplianceState -ne "Compliant" -or 
    $_.DaysSinceLastSync -gt 3 -or
    $_.StorageUsedPercent -gt 85 -or
    $_.JailBroken -eq "True"
}

if ($needsAttention) {
    Write-Host "Devices requiring attention: $($needsAttention.Count)" -ForegroundColor Yellow
    Write-Host "Review the exported data for detailed information" -ForegroundColor Yellow
} else {
    Write-Host "All devices appear to be in good condition!" -ForegroundColor Green
}

Disconnect-MgGraph
Write-Host "Disconnected from Microsoft Graph" -ForegroundColor Green