#Requires -Modules Microsoft.Graph.DeviceManagement

<#
.SYNOPSIS
    Forces policy sync on Intune-managed devices and monitors compliance
.DESCRIPTION
    Triggers immediate sync for devices that are out of compliance or haven't 
    checked in recently. Useful for troubleshooting policy deployment issues.
.PARAMETER DeviceFilter
    Filter devices by name pattern (wildcards supported)
.PARAMETER DaysStale
    Consider devices stale if not synced in X days (default: 3)
.PARAMETER WaitForSync
    Wait and monitor sync completion
.EXAMPLE
    .\Sync-IntuneDevicePolicies.ps1 -DeviceFilter "*LAPTOP*" -DaysStale 1 -WaitForSync
#>

param(
    [Parameter(Mandatory = $false)]
    [string]$DeviceFilter = "*",
    
    [Parameter(Mandatory = $false)]
    [int]$DaysStale = 3,
    
    [Parameter(Mandatory = $false)]
    [switch]$WaitForSync
)

# Connect to Microsoft Graph
try {
    Connect-MgGraph -Scopes "DeviceManagementManagedDevices.ReadWrite.All"
    Write-Host "Connected to Microsoft Graph" -ForegroundColor Green
}
catch {
    Write-Error "Failed to connect to Microsoft Graph: $($_.Exception.Message)"
    exit 1
}

Write-Host "=== INTUNE DEVICE POLICY SYNC TOOL ===" -ForegroundColor Cyan

# Get devices that need sync
Write-Host "Retrieving managed devices..." -ForegroundColor Yellow
$allDevices = Get-MgDeviceManagementManagedDevice -All

# Filter devices based on criteria
$devicesToSync = $allDevices | Where-Object {
    ($_.DeviceName -like $DeviceFilter) -and
    (
        ($_.ComplianceState -ne "Compliant") -or
        ($_.LastSyncDateTime -and (Get-Date $_.LastSyncDateTime) -lt (Get-Date).AddDays(-$DaysStale)) -or
        (-not $_.LastSyncDateTime)
    )
}

if ($devicesToSync.Count -eq 0) {
    Write-Host "No devices found matching criteria" -ForegroundColor Green
    Disconnect-MgGraph
    exit 0
}

Write-Host "`nDevices requiring sync: $($devicesToSync.Count)" -ForegroundColor Yellow

# Display devices to be synced
$devicesToSync | Select-Object DeviceName, UserPrincipalName, ComplianceState, LastSyncDateTime, OperatingSystem | 
    Format-Table -AutoSize

# Confirm action
$confirmation = Read-Host "`nProceed with sync? (y/N)"
if ($confirmation -ne 'y' -and $confirmation -ne 'Y') {
    Write-Host "Operation cancelled" -ForegroundColor Yellow
    Disconnect-MgGraph
    exit 0
}

# Track sync operations
$syncResults = @()

foreach ($device in $devicesToSync) {
    Write-Progress -Activity "Syncing devices" -Status $device.DeviceName -PercentComplete (($devicesToSync.IndexOf($device) / $devicesToSync.Count) * 100)
    
    try {
        # Trigger sync
        Invoke-MgDeviceManagementManagedDeviceSync -ManagedDeviceId $device.Id
        
        $result = [PSCustomObject]@{
            DeviceName = $device.DeviceName
            DeviceId = $device.Id
            UserPrincipalName = $device.UserPrincipalName
            SyncTriggered = $true
            SyncTime = Get-Date
            PreviousSync = $device.LastSyncDateTime
            Error = $null
        }
        
        Write-Host "✓ Sync triggered for $($device.DeviceName)" -ForegroundColor Green
    }
    catch {
        $result = [PSCustomObject]@{
            DeviceName = $device.DeviceName
            DeviceId = $device.Id
            UserPrincipalName = $device.UserPrincipalName
            SyncTriggered = $false
            SyncTime = Get-Date
            PreviousSync = $device.LastSyncDateTime
            Error = $_.Exception.Message
        }
        
        Write-Host "✗ Failed to sync $($device.DeviceName): $($_.Exception.Message)" -ForegroundColor Red
    }
    
    $syncResults += $result
    
    # Brief pause to avoid rate limiting
    Start-Sleep -Milliseconds 500
}

Write-Progress -Activity "Syncing devices" -Completed

# Summary
$successfulSyncs = ($syncResults | Where-Object SyncTriggered -eq $true).Count
$failedSyncs = ($syncResults | Where-Object SyncTriggered -eq $false).Count

Write-Host "`n=== SYNC SUMMARY ===" -ForegroundColor Cyan
Write-Host "Successful: $successfulSyncs" -ForegroundColor Green
Write-Host "Failed: $failedSyncs" -ForegroundColor Red

if ($failedSyncs -gt 0) {
    Write-Host "`nFailed devices:" -ForegroundColor Red
    $syncResults | Where-Object SyncTriggered -eq $false | 
        Select-Object DeviceName, Error | Format-Table -AutoSize
}

# Monitor sync completion if requested
if ($WaitForSync -and $successfulSyncs -gt 0) {
    Write-Host "`nMonitoring sync completion..." -ForegroundColor Yellow
    $monitorDevices = $syncResults | Where-Object SyncTriggered -eq $true
    
    $timeout = (Get-Date).AddMinutes(10)
    $completed = @()
    
    do {
        Start-Sleep -Seconds 30
        
        foreach ($device in $monitorDevices) {
            if ($device.DeviceId -in $completed) { continue }
            
            try {
                $currentDevice = Get-MgDeviceManagementManagedDevice -ManagedDeviceId $device.DeviceId
                
                if ($currentDevice.LastSyncDateTime -and 
                    (Get-Date $currentDevice.LastSyncDateTime) -gt $device.SyncTime) {
                    
                    Write-Host "✓ $($device.DeviceName) completed sync at $($currentDevice.LastSyncDateTime)" -ForegroundColor Green
                    $completed += $device.DeviceId
                }
            }
            catch {
                Write-Host "⚠ Error checking $($device.DeviceName): $($_.Exception.Message)" -ForegroundColor Yellow
            }
        }
        
        $remainingDevices = $monitorDevices.Count - $completed.Count
        if ($remainingDevices -gt 0) {
            Write-Host "Waiting for $remainingDevices devices to complete sync..." -ForegroundColor Yellow
        }
        
    } while ($completed.Count -lt $monitorDevices.Count -and (Get-Date) -lt $timeout)
    
    if ((Get-Date) -ge $timeout) {
        Write-Host "⚠ Timeout reached. Some devices may still be syncing." -ForegroundColor Yellow
    }
    
    # Final compliance check
    Write-Host "`nChecking updated compliance status..." -ForegroundColor Yellow
    
    foreach ($device in $monitorDevices) {
        try {
            $updatedDevice = Get-MgDeviceManagementManagedDevice -ManagedDeviceId $device.DeviceId
            
            $complianceChange = if ($updatedDevice.ComplianceState -ne $devicesToSync | Where-Object Id -eq $device.DeviceId | Select-Object -ExpandProperty ComplianceState) {
                "→ $($updatedDevice.ComplianceState)"
            } else {
                "(no change)"
            }
            
            Write-Host "$($device.DeviceName): $($updatedDevice.ComplianceState) $complianceChange" -ForegroundColor Cyan
        }
        catch {
            Write-Host "$($device.DeviceName): Error retrieving updated status" -ForegroundColor Red
        }
    }
}

# Export results
$reportPath = ".\IntuneSync-Report-$(Get-Date -Format 'yyyy-MM-dd-HHmm').csv"
$syncResults | Export-Csv -Path $reportPath -NoTypeInformation
Write-Host "`nSync report exported to: $reportPath" -ForegroundColor Green

# Additional troubleshooting info
Write-Host "`n=== TROUBLESHOOTING TIPS ===" -ForegroundColor Cyan
Write-Host "• If devices don't sync: Check device connectivity and Intune agent status"
Write-Host "• For persistent non-compliance: Review policy assignments and conflicts"
Write-Host "• iOS devices: May require user interaction for certain policy types"
Write-Host "• Android devices: Check work profile status if using Android Enterprise"

Disconnect-MgGraph
Write-Host "Disconnected from Microsoft Graph" -ForegroundColor Green