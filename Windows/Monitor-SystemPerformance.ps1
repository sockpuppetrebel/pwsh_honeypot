<#
.SYNOPSIS
Comprehensive system performance monitoring and baseline establishment

.DESCRIPTION
Monitors system performance metrics including CPU, memory, disk, and network usage.
Establishes performance baselines and identifies anomalies. Provides trending
analysis and capacity planning recommendations.

.PARAMETER ComputerName
Target computer(s) to monitor. Defaults to local machine.

.PARAMETER Duration
Monitoring duration in minutes. Default: 60

.PARAMETER Interval
Sample interval in seconds. Default: 30

.PARAMETER ExportPath
Path to save performance data and reports

.PARAMETER EstablishBaseline
Create new performance baseline for comparison

.PARAMETER CompareToBaseline
Compare current performance to existing baseline

.PARAMETER AlertThresholds
Custom alert thresholds (JSON format or hashtable)

.EXAMPLE
Monitor-SystemPerformance -Duration 120 -Interval 15
Monitors local system for 2 hours with 15-second intervals

.EXAMPLE
Monitor-SystemPerformance -ComputerName "SERVER01" -EstablishBaseline -Duration 1440
Creates 24-hour baseline for SERVER01

.NOTES
Author: Enterprise PowerShell Collection
Requires: Administrative privileges for full metrics
Version: 1.0
#>

[CmdletBinding()]
param(
    [Parameter(ValueFromPipeline = $true)]
    [string[]]$ComputerName = $env:COMPUTERNAME,
    
    [Parameter()]
    [int]$Duration = 60,
    
    [Parameter()]
    [int]$Interval = 30,
    
    [Parameter()]
    [string]$ExportPath,
    
    [Parameter()]
    [switch]$EstablishBaseline,
    
    [Parameter()]
    [switch]$CompareToBaseline,
    
    [Parameter()]
    [hashtable]$AlertThresholds = @{
        CPUPercent = 80
        MemoryPercent = 85
        DiskPercent = 90
        DiskQueueLength = 2
        NetworkUtilization = 80
    }
)

begin {
    Write-Host "Starting System Performance Monitoring..." -ForegroundColor Cyan
    
    # Setup export directory
    if (-not $ExportPath) {
        $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
        $ExportPath = "PerfMonitoring_$timestamp"
    }
    
    if (-not (Test-Path $ExportPath)) {
        New-Item -Path $ExportPath -ItemType Directory -Force | Out-Null
    }
    
    $results = @()
    $samples = [math]::Ceiling($Duration * 60 / $Interval)
    
    Write-Host "Configuration:" -ForegroundColor Yellow
    Write-Host "  Duration: $Duration minutes" -ForegroundColor White
    Write-Host "  Interval: $Interval seconds" -ForegroundColor White
    Write-Host "  Samples: $samples" -ForegroundColor White
    Write-Host "  Export Path: $ExportPath" -ForegroundColor White
}

process {
    foreach ($computer in $ComputerName) {
        Write-Host "`nStarting performance monitoring for: $computer" -ForegroundColor Yellow
        
        try {
            $performanceData = @()
            $alerts = @()
            
            # Initialize performance counters
            $counters = @(
                "\\$computer\Processor(_Total)\% Processor Time",
                "\\$computer\Memory\Available MBytes",
                "\\$computer\Memory\Committed Bytes",
                "\\$computer\PhysicalDisk(_Total)\% Disk Time",
                "\\$computer\PhysicalDisk(_Total)\Current Disk Queue Length",
                "\\$computer\PhysicalDisk(_Total)\Disk Bytes/sec",
                "\\$computer\Network Interface(*)\Bytes Total/sec",
                "\\$computer\System\Processor Queue Length",
                "\\$computer\System\Context Switches/sec"
            )
            
            # Get baseline memory info for calculations
            $memoryInfo = Get-CimInstance -ComputerName $computer -ClassName Win32_ComputerSystem
            $totalMemoryMB = [math]::Round($memoryInfo.TotalPhysicalMemory / 1MB, 0)
            
            Write-Host "  Collecting $samples samples over $Duration minutes..." -ForegroundColor Cyan
            
            for ($i = 1; $i -le $samples; $i++) {
                $sampleTime = Get-Date
                Write-Progress -Activity "Performance Monitoring: $computer" -Status "Sample $i of $samples" -PercentComplete (($i / $samples) * 100)
                
                try {
                    # Get performance counter data
                    $counterData = Get-Counter -ComputerName $computer -Counter $counters -MaxSamples 1 -ErrorAction Stop
                    
                    # Parse counter values
                    $cpuUsage = [math]::Round(($counterData.CounterSamples | Where-Object Path -like "*Processor Time").CookedValue, 2)
                    $availableMemoryMB = [math]::Round(($counterData.CounterSamples | Where-Object Path -like "*Available MBytes").CookedValue, 0)
                    $memoryUsagePercent = [math]::Round((($totalMemoryMB - $availableMemoryMB) / $totalMemoryMB) * 100, 2)
                    $diskTime = [math]::Round(($counterData.CounterSamples | Where-Object Path -like "*% Disk Time").CookedValue, 2)
                    $diskQueue = [math]::Round(($counterData.CounterSamples | Where-Object Path -like "*Current Disk Queue Length").CookedValue, 2)
                    $diskBytesPerSec = [math]::Round(($counterData.CounterSamples | Where-Object Path -like "*Disk Bytes/sec").CookedValue / 1MB, 2)
                    $processorQueue = [math]::Round(($counterData.CounterSamples | Where-Object Path -like "*Processor Queue Length").CookedValue, 0)
                    $contextSwitches = [math]::Round(($counterData.CounterSamples | Where-Object Path -like "*Context Switches/sec").CookedValue, 0)
                    
                    # Network utilization (sum all interfaces)
                    $networkBytes = ($counterData.CounterSamples | Where-Object Path -like "*Network Interface*Bytes Total/sec").CookedValue | Measure-Object -Sum
                    $networkMBps = [math]::Round($networkBytes.Sum / 1MB, 2)
                    
                    $sample = [PSCustomObject]@{
                        ComputerName = $computer
                        Timestamp = $sampleTime
                        SampleNumber = $i
                        CPUPercent = $cpuUsage
                        MemoryUsedPercent = $memoryUsagePercent
                        MemoryAvailableMB = $availableMemoryMB
                        DiskTimePercent = $diskTime
                        DiskQueueLength = $diskQueue
                        DiskMBPerSec = $diskBytesPerSec
                        NetworkMBPerSec = $networkMBps
                        ProcessorQueueLength = $processorQueue
                        ContextSwitchesPerSec = $contextSwitches
                    }
                    
                    $performanceData += $sample
                    
                    # Check alert thresholds
                    if ($cpuUsage -gt $AlertThresholds.CPUPercent) {
                        $alerts += [PSCustomObject]@{ Time = $sampleTime; Metric = "CPU"; Value = $cpuUsage; Threshold = $AlertThresholds.CPUPercent }
                    }
                    if ($memoryUsagePercent -gt $AlertThresholds.MemoryPercent) {
                        $alerts += [PSCustomObject]@{ Time = $sampleTime; Metric = "Memory"; Value = $memoryUsagePercent; Threshold = $AlertThresholds.MemoryPercent }
                    }
                    if ($diskTime -gt $AlertThresholds.DiskPercent) {
                        $alerts += [PSCustomObject]@{ Time = $sampleTime; Metric = "Disk"; Value = $diskTime; Threshold = $AlertThresholds.DiskPercent }
                    }
                    if ($diskQueue -gt $AlertThresholds.DiskQueueLength) {
                        $alerts += [PSCustomObject]@{ Time = $sampleTime; Metric = "DiskQueue"; Value = $diskQueue; Threshold = $AlertThresholds.DiskQueueLength }
                    }
                    
                    # Real-time display every 10 samples
                    if ($i % 10 -eq 0) {
                        Write-Host "    Sample $i - CPU: $cpuUsage%, Memory: $memoryUsagePercent%, Disk: $diskTime%" -ForegroundColor Gray
                    }
                    
                } catch {
                    Write-Warning "Failed to collect sample $i for $computer : $_"
                }
                
                if ($i -lt $samples) {
                    Start-Sleep -Seconds $Interval
                }
            }
            
            Write-Progress -Activity "Performance Monitoring: $computer" -Completed
            
            # Calculate statistics
            $stats = [PSCustomObject]@{
                ComputerName = $computer
                MonitoringPeriod = @{
                    Start = ($performanceData | Measure-Object Timestamp -Minimum).Minimum
                    End = ($performanceData | Measure-Object Timestamp -Maximum).Maximum
                    Duration = $Duration
                    Samples = $performanceData.Count
                }
                CPU = @{
                    Average = [math]::Round(($performanceData | Measure-Object CPUPercent -Average).Average, 2)
                    Maximum = [math]::Round(($performanceData | Measure-Object CPUPercent -Maximum).Maximum, 2)
                    Minimum = [math]::Round(($performanceData | Measure-Object CPUPercent -Minimum).Minimum, 2)
                    Above80Percent = ($performanceData | Where-Object CPUPercent -gt 80).Count
                }
                Memory = @{
                    AverageUsedPercent = [math]::Round(($performanceData | Measure-Object MemoryUsedPercent -Average).Average, 2)
                    MaximumUsedPercent = [math]::Round(($performanceData | Measure-Object MemoryUsedPercent -Maximum).Maximum, 2)
                    MinimumAvailableMB = [math]::Round(($performanceData | Measure-Object MemoryAvailableMB -Minimum).Minimum, 0)
                    TotalMemoryMB = $totalMemoryMB
                }
                Disk = @{
                    AverageTimePercent = [math]::Round(($performanceData | Measure-Object DiskTimePercent -Average).Average, 2)
                    MaximumTimePercent = [math]::Round(($performanceData | Measure-Object DiskTimePercent -Maximum).Maximum, 2)
                    AverageQueueLength = [math]::Round(($performanceData | Measure-Object DiskQueueLength -Average).Average, 2)
                    MaximumQueueLength = [math]::Round(($performanceData | Measure-Object DiskQueueLength -Maximum).Maximum, 2)
                    AverageMBPerSec = [math]::Round(($performanceData | Measure-Object DiskMBPerSec -Average).Average, 2)
                }
                Network = @{
                    AverageMBPerSec = [math]::Round(($performanceData | Measure-Object NetworkMBPerSec -Average).Average, 2)
                    MaximumMBPerSec = [math]::Round(($performanceData | Measure-Object NetworkMBPerSec -Maximum).Maximum, 2)
                }
                System = @{
                    AverageProcessorQueue = [math]::Round(($performanceData | Measure-Object ProcessorQueueLength -Average).Average, 2)
                    AverageContextSwitches = [math]::Round(($performanceData | Measure-Object ContextSwitchesPerSec -Average).Average, 0)
                }
                Alerts = $alerts
                RawData = $performanceData
            }
            
            # Performance assessment
            $issues = @()
            $recommendations = @()
            
            if ($stats.CPU.Average -gt 70) {
                $issues += "High average CPU utilization ($($stats.CPU.Average)%)"
                $recommendations += "Investigate high CPU usage processes and consider hardware upgrade"
            }
            
            if ($stats.Memory.AverageUsedPercent -gt 80) {
                $issues += "High average memory utilization ($($stats.Memory.AverageUsedPercent)%)"
                $recommendations += "Consider adding more RAM or optimizing memory usage"
            }
            
            if ($stats.Disk.AverageQueueLength -gt 2) {
                $issues += "High disk queue length ($($stats.Disk.AverageQueueLength))"
                $recommendations += "Disk I/O bottleneck detected - consider faster storage or I/O optimization"
            }
            
            if ($stats.System.AverageProcessorQueue -gt 2) {
                $issues += "High processor queue length ($($stats.System.AverageProcessorQueue))"
                $recommendations += "CPU bottleneck detected - consider processor upgrade or workload optimization"
            }
            
            $stats | Add-Member -NotePropertyName "Issues" -NotePropertyValue $issues
            $stats | Add-Member -NotePropertyName "Recommendations" -NotePropertyValue $recommendations
            
            # Baseline operations
            if ($EstablishBaseline) {
                $baselineFile = "$ExportPath\Baseline_$computer.json"
                $stats | ConvertTo-Json -Depth 5 | Out-File $baselineFile
                Write-Host "  ✓ Baseline established and saved to $baselineFile" -ForegroundColor Green
            }
            
            if ($CompareToBaseline) {
                $baselineFile = "$ExportPath\Baseline_$computer.json"
                if (Test-Path $baselineFile) {
                    $baseline = Get-Content $baselineFile | ConvertFrom-Json
                    $comparison = @{
                        CPUChange = $stats.CPU.Average - $baseline.CPU.Average
                        MemoryChange = $stats.Memory.AverageUsedPercent - $baseline.Memory.AverageUsedPercent
                        DiskChange = $stats.Disk.AverageTimePercent - $baseline.Disk.AverageTimePercent
                    }
                    $stats | Add-Member -NotePropertyName "BaselineComparison" -NotePropertyValue $comparison
                    Write-Host "  ✓ Compared to baseline" -ForegroundColor Green
                } else {
                    Write-Warning "No baseline found for $computer"
                }
            }
            
            $results += $stats
            
            Write-Host "  ✓ Performance monitoring completed for $computer" -ForegroundColor Green
            Write-Host "    Average CPU: $($stats.CPU.Average)%, Memory: $($stats.Memory.AverageUsedPercent)%, Disk: $($stats.Disk.AverageTimePercent)%" -ForegroundColor White
            
        } catch {
            Write-Error "Failed to monitor $computer : $_"
            $results += [PSCustomObject]@{
                ComputerName = $computer
                Error = $_.Exception.Message
                Status = "Error"
            }
        }
    }
}

end {
    # Export detailed results
    Write-Host "`nExporting performance data..." -ForegroundColor Cyan
    
    $results | Export-Clixml -Path "$ExportPath\PerformanceResults.xml"
    $results | ConvertTo-Json -Depth 6 | Out-File "$ExportPath\PerformanceResults.json"
    
    # Export raw performance data
    $allRawData = $results | ForEach-Object { $_.RawData }
    $allRawData | Export-Csv -Path "$ExportPath\RawPerformanceData.csv" -NoTypeInformation
    
    # Create executive summary
    $summary = $results | Where-Object Error -eq $null | Select-Object ComputerName,
        @{N="AvgCPU%";E={$_.CPU.Average}},
        @{N="MaxCPU%";E={$_.CPU.Maximum}},
        @{N="AvgMemory%";E={$_.Memory.AverageUsedPercent}},
        @{N="AvgDisk%";E={$_.Disk.AverageTimePercent}},
        @{N="DiskQueue";E={$_.Disk.AverageQueueLength}},
        @{N="AlertCount";E={$_.Alerts.Count}},
        @{N="IssueCount";E={$_.Issues.Count}}
    
    $summary | Export-Csv -Path "$ExportPath\PerformanceSummary.csv" -NoTypeInformation
    
    # Display summary
    Write-Host "`n" + "="*100 -ForegroundColor Cyan
    Write-Host "SYSTEM PERFORMANCE MONITORING SUMMARY" -ForegroundColor Cyan
    Write-Host "="*100 -ForegroundColor Cyan
    
    $summary | Format-Table -AutoSize
    
    # Overall statistics
    $overallStats = $summary | Measure-Object "AvgCPU%", "AvgMemory%", "AvgDisk%" -Average
    Write-Host "Overall Averages:" -ForegroundColor Yellow
    Write-Host "  CPU Usage: $([math]::Round($overallStats[0].Average, 1))%" -ForegroundColor White
    Write-Host "  Memory Usage: $([math]::Round($overallStats[1].Average, 1))%" -ForegroundColor White
    Write-Host "  Disk Usage: $([math]::Round($overallStats[2].Average, 1))%" -ForegroundColor White
    
    # Alert summary
    $totalAlerts = ($results | ForEach-Object { $_.Alerts }).Count
    if ($totalAlerts -gt 0) {
        Write-Host "`nAlert Summary: $totalAlerts total alerts" -ForegroundColor Red
        $alertsByMetric = $results | ForEach-Object { $_.Alerts } | Group-Object Metric
        $alertsByMetric | ForEach-Object {
            Write-Host "  $($_.Name): $($_.Count) alerts" -ForegroundColor Yellow
        }
    }
    
    # Recommendations
    $allRecommendations = $results | ForEach-Object { $_.Recommendations } | Group-Object | Sort-Object Count -Descending
    if ($allRecommendations) {
        Write-Host "`nTop Recommendations:" -ForegroundColor Yellow
        $allRecommendations | Select-Object -First 5 | ForEach-Object {
            Write-Host "  • $($_.Name) ($($_.Count) systems)" -ForegroundColor Cyan
        }
    }
    
    Write-Host "`nDetailed performance data saved to: $ExportPath" -ForegroundColor Green
    Write-Host "Performance monitoring completed for $($results.Count) systems" -ForegroundColor Cyan
}