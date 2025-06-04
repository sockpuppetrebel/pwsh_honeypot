<#
.SYNOPSIS
Generates comprehensive system health report for Windows endpoints

.DESCRIPTION
Performs detailed system health assessment including hardware, performance,
storage, services, and event log analysis. Provides recommendations for
optimization and maintenance. Essential for proactive system monitoring.

.PARAMETER ComputerName
Target computer(s) to analyze. Defaults to local machine.

.PARAMETER ExportPath
Path to save the detailed report. Creates timestamped folder if not specified.

.PARAMETER IncludePerformanceCounters
Include detailed performance counter analysis (may take longer).

.PARAMETER Days
Number of days to analyze for event logs and performance trends. Default: 7

.EXAMPLE
Get-SystemHealthReport
Analyzes local system health

.EXAMPLE
Get-SystemHealthReport -ComputerName "SERVER01","WORKSTATION02" -Days 30
Analyzes multiple systems with 30-day historical data

.NOTES
Author: Enterprise PowerShell Collection
Requires: Administrative privileges for full analysis
Version: 1.0
#>

[CmdletBinding()]
param(
    [Parameter(ValueFromPipeline = $true)]
    [string[]]$ComputerName = $env:COMPUTERNAME,
    
    [Parameter()]
    [string]$ExportPath,
    
    [Parameter()]
    [switch]$IncludePerformanceCounters,
    
    [Parameter()]
    [int]$Days = 7
)

begin {
    Write-Host "Starting System Health Assessment..." -ForegroundColor Cyan
    
    # Setup export directory
    if (-not $ExportPath) {
        $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
        $ExportPath = "SystemHealth_$timestamp"
    }
    
    if (-not (Test-Path $ExportPath)) {
        New-Item -Path $ExportPath -ItemType Directory -Force | Out-Null
    }
    
    $results = @()
    $startTime = (Get-Date).AddDays(-$Days)
}

process {
    foreach ($computer in $ComputerName) {
        Write-Host "Analyzing system: $computer" -ForegroundColor Yellow
        
        try {
            $computerResult = [PSCustomObject]@{
                ComputerName = $computer
                Timestamp = Get-Date
                SystemInfo = $null
                Hardware = $null
                Storage = $null
                Services = $null
                Performance = $null
                EventLogs = $null
                SecurityStatus = $null
                Recommendations = @()
                OverallHealth = "Unknown"
            }
            
            # System Information
            Write-Progress -Activity "Health Check: $computer" -Status "Gathering system information" -PercentComplete 10
            $computerResult.SystemInfo = Get-CimInstance -ComputerName $computer -ClassName Win32_ComputerSystem -ErrorAction Stop | Select-Object Name, Manufacturer, Model, TotalPhysicalMemory, NumberOfProcessors
            
            $osInfo = Get-CimInstance -ComputerName $computer -ClassName Win32_OperatingSystem -ErrorAction Stop
            $computerResult.SystemInfo | Add-Member -NotePropertyName "OS" -NotePropertyValue "$($osInfo.Caption) $($osInfo.Version)"
            $computerResult.SystemInfo | Add-Member -NotePropertyName "LastBootTime" -NotePropertyValue $osInfo.LastBootUpTime
            $computerResult.SystemInfo | Add-Member -NotePropertyName "UptimeDays" -NotePropertyValue ([math]::Round((Get-Date - $osInfo.LastBootUpTime).TotalDays, 2))
            
            # Hardware Health
            Write-Progress -Activity "Health Check: $computer" -Status "Checking hardware health" -PercentComplete 25
            $hardware = @{
                CPU = Get-CimInstance -ComputerName $computer -ClassName Win32_Processor | Select-Object Name, MaxClockSpeed, NumberOfCores, NumberOfLogicalProcessors
                Memory = @{
                    TotalGB = [math]::Round($computerResult.SystemInfo.TotalPhysicalMemory / 1GB, 2)
                    Available = try { [math]::Round((Get-CimInstance -ComputerName $computer -ClassName Win32_PerfRawData_PerfOS_Memory).AvailableBytes / 1GB, 2) } catch { "N/A" }
                }
                Disks = Get-CimInstance -ComputerName $computer -ClassName Win32_LogicalDisk -Filter "DriveType=3" | 
                    Select-Object DeviceID, @{N="SizeGB";E={[math]::Round($_.Size/1GB,2)}}, @{N="FreeGB";E={[math]::Round($_.FreeSpace/1GB,2)}}, @{N="PercentFree";E={[math]::Round(($_.FreeSpace/$_.Size)*100,2)}}
            }
            $computerResult.Hardware = $hardware
            
            # Storage Analysis
            Write-Progress -Activity "Health Check: $computer" -Status "Analyzing storage" -PercentComplete 40
            $storageIssues = @()
            foreach ($disk in $hardware.Disks) {
                if ($disk.PercentFree -lt 10) {
                    $storageIssues += "Drive $($disk.DeviceID) critically low on space ($($disk.PercentFree)% free)"
                    $computerResult.Recommendations += "Immediate action required: Free up space on drive $($disk.DeviceID)"
                } elseif ($disk.PercentFree -lt 20) {
                    $storageIssues += "Drive $($disk.DeviceID) low on space ($($disk.PercentFree)% free)"
                    $computerResult.Recommendations += "Warning: Monitor space usage on drive $($disk.DeviceID)"
                }
            }
            $computerResult.Storage = @{ Issues = $storageIssues; Details = $hardware.Disks }
            
            # Critical Services Status
            Write-Progress -Activity "Health Check: $computer" -Status "Checking critical services" -PercentComplete 55
            $criticalServices = @("Themes", "DHCP", "DNS", "Eventlog", "LanmanServer", "LanmanWorkstation", "RpcSs", "SamSs", "Schedule", "Spooler", "Winmgmt")
            $serviceStatus = foreach ($serviceName in $criticalServices) {
                try {
                    $service = Get-CimInstance -ComputerName $computer -ClassName Win32_Service -Filter "Name='$serviceName'" -ErrorAction Stop
                    [PSCustomObject]@{
                        Name = $service.Name
                        DisplayName = $service.DisplayName
                        State = $service.State
                        StartMode = $service.StartMode
                        Status = if ($service.State -eq "Running") { "OK" } else { "Issue" }
                    }
                } catch {
                    [PSCustomObject]@{ Name = $serviceName; DisplayName = "Unknown"; State = "Unknown"; StartMode = "Unknown"; Status = "Error" }
                }
            }
            $computerResult.Services = $serviceStatus
            
            # Event Log Analysis
            Write-Progress -Activity "Health Check: $computer" -Status "Analyzing event logs" -PercentComplete 70
            $eventAnalysis = @{
                SystemErrors = 0
                ApplicationErrors = 0
                SecurityWarnings = 0
                RecentCritical = @()
            }
            
            try {
                $systemErrors = Get-WinEvent -ComputerName $computer -FilterHashtable @{LogName='System'; Level=1,2; StartTime=$startTime} -ErrorAction SilentlyContinue
                $eventAnalysis.SystemErrors = ($systemErrors | Measure-Object).Count
                
                $appErrors = Get-WinEvent -ComputerName $computer -FilterHashtable @{LogName='Application'; Level=1,2; StartTime=$startTime} -ErrorAction SilentlyContinue
                $eventAnalysis.ApplicationErrors = ($appErrors | Measure-Object).Count
                
                $criticalEvents = $systemErrors | Where-Object Level -eq 1 | Select-Object -First 5 TimeCreated, Id, LevelDisplayName, ProviderName, Message
                $eventAnalysis.RecentCritical = $criticalEvents
                
                if ($eventAnalysis.SystemErrors -gt 50) {
                    $computerResult.Recommendations += "High number of system errors detected ($($eventAnalysis.SystemErrors)). Review system event log."
                }
                if ($eventAnalysis.ApplicationErrors -gt 100) {
                    $computerResult.Recommendations += "High number of application errors detected ($($eventAnalysis.ApplicationErrors)). Review application event log."
                }
            } catch {
                Write-Warning "Could not analyze event logs for $computer"
            }
            $computerResult.EventLogs = $eventAnalysis
            
            # Performance Counters (if requested)
            if ($IncludePerformanceCounters) {
                Write-Progress -Activity "Health Check: $computer" -Status "Gathering performance data" -PercentComplete 85
                try {
                    $perfCounters = @{
                        CPUUsage = (Get-CimInstance -ComputerName $computer -ClassName Win32_PerfRawData_PerfOS_Processor -Filter "Name='_Total'").PercentProcessorTime
                        MemoryUsage = [math]::Round(((Get-CimInstance -ComputerName $computer -ClassName Win32_PerfRawData_PerfOS_Memory).CommittedBytes / (Get-CimInstance -ComputerName $computer -ClassName Win32_ComputerSystem).TotalPhysicalMemory) * 100, 2)
                        DiskQueue = (Get-CimInstance -ComputerName $computer -ClassName Win32_PerfRawData_PerfDisk_LogicalDisk -Filter "Name='_Total'").CurrentDiskQueueLength
                    }
                    $computerResult.Performance = $perfCounters
                } catch {
                    $computerResult.Performance = @{ Error = "Could not gather performance data" }
                }
            }
            
            # Overall Health Assessment
            Write-Progress -Activity "Health Check: $computer" -Status "Calculating health score" -PercentComplete 95
            $healthScore = 100
            
            # Deduct points for issues
            if ($computerResult.SystemInfo.UptimeDays -gt 30) { $healthScore -= 10; $computerResult.Recommendations += "Consider rebooting system (uptime: $($computerResult.SystemInfo.UptimeDays) days)" }
            if ($storageIssues.Count -gt 0) { $healthScore -= 20 }
            if (($serviceStatus | Where-Object Status -eq "Issue").Count -gt 0) { $healthScore -= 15 }
            if ($eventAnalysis.SystemErrors -gt 20) { $healthScore -= 15 }
            if ($eventAnalysis.ApplicationErrors -gt 50) { $healthScore -= 10 }
            
            $computerResult.OverallHealth = switch ($healthScore) {
                {$_ -ge 90} { "Excellent" }
                {$_ -ge 80} { "Good" }
                {$_ -ge 70} { "Fair" }
                {$_ -ge 60} { "Poor" }
                default { "Critical" }
            }
            
            $results += $computerResult
            
            Write-Host "✓ Completed analysis for $computer - Health: $($computerResult.OverallHealth) ($healthScore%)" -ForegroundColor Green
            
        } catch {
            Write-Error "Failed to analyze $computer : $_"
            $results += [PSCustomObject]@{
                ComputerName = $computer
                Error = $_.Exception.Message
                OverallHealth = "Error"
            }
        }
        
        Write-Progress -Activity "Health Check: $computer" -Completed
    }
}

end {
    # Export detailed results
    Write-Host "`nExporting detailed results..." -ForegroundColor Cyan
    
    $results | Export-Clixml -Path "$ExportPath\SystemHealthResults.xml"
    $results | ConvertTo-Json -Depth 5 | Out-File "$ExportPath\SystemHealthResults.json"
    
    # Create summary report
    $summary = $results | Select-Object ComputerName, OverallHealth, @{N="UptimeDays";E={$_.SystemInfo.UptimeDays}}, 
        @{N="StorageIssues";E={$_.Storage.Issues.Count}}, @{N="ServiceIssues";E={($_.Services | Where-Object Status -eq "Issue").Count}},
        @{N="SystemErrors";E={$_.EventLogs.SystemErrors}}, @{N="Recommendations";E={$_.Recommendations.Count}}
    
    $summary | Export-Csv -Path "$ExportPath\SystemHealthSummary.csv" -NoTypeInformation
    
    # Display summary
    Write-Host "`n" + "="*80 -ForegroundColor Cyan
    Write-Host "SYSTEM HEALTH ASSESSMENT SUMMARY" -ForegroundColor Cyan
    Write-Host "="*80 -ForegroundColor Cyan
    
    $summary | Format-Table -AutoSize
    
    # Health distribution
    $healthDistribution = $results | Group-Object OverallHealth | Sort-Object Name
    Write-Host "`nHealth Distribution:" -ForegroundColor Yellow
    $healthDistribution | ForEach-Object {
        $color = switch ($_.Name) {
            "Excellent" { "Green" }
            "Good" { "Green" }
            "Fair" { "Yellow" }
            "Poor" { "Red" }
            "Critical" { "Red" }
            default { "White" }
        }
        Write-Host "  $($_.Name): $($_.Count) systems" -ForegroundColor $color
    }
    
    # Top recommendations
    $allRecommendations = $results | ForEach-Object { $_.Recommendations } | Group-Object | Sort-Object Count -Descending | Select-Object -First 5
    if ($allRecommendations) {
        Write-Host "`nTop Recommendations:" -ForegroundColor Yellow
        $allRecommendations | ForEach-Object {
            Write-Host "  • $($_.Name) ($($_.Count) systems)" -ForegroundColor White
        }
    }
    
    Write-Host "`nDetailed reports saved to: $ExportPath" -ForegroundColor Green
    Write-Host "Analysis completed for $($results.Count) systems" -ForegroundColor Cyan
}