<#
.SYNOPSIS
Advanced event log analysis and alerting for Windows endpoints

.DESCRIPTION
Performs comprehensive event log analysis including error pattern detection,
security event monitoring, performance correlation, and trend analysis.
Generates actionable reports and alerts for system administrators.

.PARAMETER ComputerName
Target computer(s) to analyze. Defaults to local machine.

.PARAMETER LogNames
Event log names to analyze. Default: System, Application, Security

.PARAMETER Hours
Number of hours to analyze from current time. Default: 24

.PARAMETER ErrorThreshold
Minimum number of errors to trigger investigation. Default: 10

.PARAMETER ExportPath
Path to save analysis reports and alerts

.PARAMETER IncludeSecurity
Include security event analysis (requires elevated privileges)

.PARAMETER AlertEmail
Email address to send critical alerts

.EXAMPLE
Analyze-EventLogs -Hours 72 -ErrorThreshold 5
Analyzes local system logs for past 3 days with low error threshold

.EXAMPLE
Analyze-EventLogs -ComputerName "SERVER01","SERVER02" -IncludeSecurity -AlertEmail "admin@company.com"
Analyzes multiple servers including security events with email alerts

.NOTES
Author: Enterprise PowerShell Collection
Requires: Administrative privileges for Security log access
Version: 1.0
#>

[CmdletBinding()]
param(
    [Parameter(ValueFromPipeline = $true)]
    [string[]]$ComputerName = $env:COMPUTERNAME,
    
    [Parameter()]
    [string[]]$LogNames = @("System", "Application", "Security"),
    
    [Parameter()]
    [int]$Hours = 24,
    
    [Parameter()]
    [int]$ErrorThreshold = 10,
    
    [Parameter()]
    [string]$ExportPath,
    
    [Parameter()]
    [switch]$IncludeSecurity,
    
    [Parameter()]
    [string]$AlertEmail
)

begin {
    Write-Host "Starting Event Log Analysis..." -ForegroundColor Cyan
    
    # Setup export directory
    if (-not $ExportPath) {
        $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
        $ExportPath = "EventLogAnalysis_$timestamp"
    }
    
    if (-not (Test-Path $ExportPath)) {
        New-Item -Path $ExportPath -ItemType Directory -Force | Out-Null
    }
    
    $startTime = (Get-Date).AddHours(-$Hours)
    $results = @()
    
    # Define critical event patterns
    $criticalPatterns = @{
        SystemFailure = @(41, 1074, 1076, 6008, 6009, 6013)  # System crashes, shutdowns
        ServiceFailure = @(7000, 7001, 7009, 7011, 7023, 7024, 7031, 7034)  # Service errors
        DiskIssues = @(7, 9, 11, 15, 51, 52, 153)  # Disk/storage errors
        MemoryIssues = @(1001, 1003, 1019, 1020, 1021)  # Memory errors
        NetworkIssues = @(4201, 4202, 4210, 5719, 5723)  # Network errors
        SecurityAlerts = @(4624, 4625, 4648, 4672, 4720, 4732, 4756)  # Security events
    }
    
    Write-Host "Analysis Configuration:" -ForegroundColor Yellow
    Write-Host "  Time Range: Last $Hours hours (since $startTime)" -ForegroundColor White
    Write-Host "  Error Threshold: $ErrorThreshold events" -ForegroundColor White
    Write-Host "  Target Logs: $($LogNames -join ', ')" -ForegroundColor White
}

process {
    foreach ($computer in $ComputerName) {
        Write-Host "`nAnalyzing event logs for: $computer" -ForegroundColor Yellow
        
        try {
            $computerResult = [PSCustomObject]@{
                ComputerName = $computer
                AnalysisTime = Get-Date
                TimeRange = @{ Start = $startTime; End = Get-Date; Hours = $Hours }
                LogAnalysis = @{}
                CriticalEvents = @()
                ErrorPatterns = @()
                SecurityEvents = @()
                Alerts = @()
                Summary = @{}
                Status = "Success"
            }
            
            foreach ($logName in $LogNames) {
                # Skip Security log if not explicitly requested or no privileges
                if ($logName -eq "Security" -and -not $IncludeSecurity) {
                    continue
                }
                
                Write-Progress -Activity "Event Log Analysis: $computer" -Status "Analyzing $logName log" -PercentComplete ((([array]::IndexOf($LogNames, $logName)) / $LogNames.Count) * 100)
                
                try {
                    # Get events from the specified time range
                    $events = Get-WinEvent -ComputerName $computer -FilterHashtable @{
                        LogName = $logName
                        StartTime = $startTime
                        Level = 1,2,3  # Critical, Error, Warning
                    } -ErrorAction SilentlyContinue
                    
                    if ($events) {
                        # Analyze event levels
                        $logStats = @{
                            LogName = $logName
                            TotalEvents = $events.Count
                            Critical = ($events | Where-Object Level -eq 1).Count
                            Error = ($events | Where-Object Level -eq 2).Count
                            Warning = ($events | Where-Object Level -eq 3).Count
                            TopSources = ($events | Group-Object ProviderName | Sort-Object Count -Descending | Select-Object -First 10)
                            TopEventIDs = ($events | Group-Object Id | Sort-Object Count -Descending | Select-Object -First 10)
                            TimeDistribution = @{}
                        }
                        
                        # Time distribution analysis (hourly buckets)
                        for ($h = 0; $h -lt $Hours; $h++) {
                            $bucketStart = $startTime.AddHours($h)
                            $bucketEnd = $bucketStart.AddHours(1)
                            $bucketEvents = $events | Where-Object { $_.TimeCreated -ge $bucketStart -and $_.TimeCreated -lt $bucketEnd }
                            $logStats.TimeDistribution["Hour_$h"] = $bucketEvents.Count
                        }
                        
                        $computerResult.LogAnalysis[$logName] = $logStats
                        
                        # Pattern detection
                        foreach ($patternName in $criticalPatterns.Keys) {
                            $patternEvents = $events | Where-Object { $_.Id -in $criticalPatterns[$patternName] }
                            if ($patternEvents.Count -ge $ErrorThreshold) {
                                $pattern = [PSCustomObject]@{
                                    Pattern = $patternName
                                    EventCount = $patternEvents.Count
                                    LogName = $logName
                                    Events = $patternEvents | Select-Object -First 5 TimeCreated, Id, LevelDisplayName, ProviderName, Message
                                    Severity = switch ($patternName) {
                                        "SystemFailure" { "Critical" }
                                        "SecurityAlerts" { "High" }
                                        default { "Medium" }
                                    }
                                }
                                $computerResult.ErrorPatterns += $pattern
                                
                                # Generate alerts for critical patterns
                                if ($pattern.Severity -in @("Critical", "High")) {
                                    $alert = [PSCustomObject]@{
                                        Computer = $computer
                                        AlertType = $patternName
                                        Severity = $pattern.Severity
                                        Count = $patternEvents.Count
                                        FirstOccurrence = ($patternEvents | Sort-Object TimeCreated | Select-Object -First 1).TimeCreated
                                        LastOccurrence = ($patternEvents | Sort-Object TimeCreated -Descending | Select-Object -First 1).TimeCreated
                                        Message = "Pattern '$patternName' detected: $($patternEvents.Count) events in $logName log"
                                    }
                                    $computerResult.Alerts += $alert
                                }
                            }
                        }
                        
                        # Collect critical events for detailed analysis
                        $criticalEvents = $events | Where-Object Level -eq 1 | Sort-Object TimeCreated -Descending | Select-Object -First 20
                        $computerResult.CriticalEvents += $criticalEvents | Select-Object LogName, TimeCreated, Id, LevelDisplayName, ProviderName, Message
                        
                        # Security event analysis (if Security log)
                        if ($logName -eq "Security" -and $IncludeSecurity) {
                            $securityAnalysis = @{
                                FailedLogons = ($events | Where-Object Id -eq 4625).Count
                                SuccessfulLogons = ($events | Where-Object Id -eq 4624).Count
                                PrivilegeUse = ($events | Where-Object Id -eq 4672).Count
                                AccountCreation = ($events | Where-Object Id -eq 4720).Count
                                GroupChanges = ($events | Where-Object Id -eq 4732).Count
                                PolicyChanges = ($events | Where-Object Id -eq 4756).Count
                            }
                            
                            # Check for suspicious activity
                            if ($securityAnalysis.FailedLogons -gt 50) {
                                $computerResult.Alerts += [PSCustomObject]@{
                                    Computer = $computer
                                    AlertType = "SuspiciousLogonActivity"
                                    Severity = "High"
                                    Count = $securityAnalysis.FailedLogons
                                    Message = "High number of failed logon attempts: $($securityAnalysis.FailedLogons)"
                                }
                            }
                            
                            $computerResult.SecurityEvents = $securityAnalysis
                        }
                        
                        Write-Host "    $logName: $($events.Count) events ($($logStats.Critical) critical, $($logStats.Error) errors)" -ForegroundColor Gray
                        
                    } else {
                        Write-Host "    $logName: No events found in time range" -ForegroundColor Gray
                        $computerResult.LogAnalysis[$logName] = @{ LogName = $logName; TotalEvents = 0; Message = "No events in time range" }
                    }
                    
                } catch {
                    Write-Warning "Could not access $logName log on $computer : $_"
                    $computerResult.LogAnalysis[$logName] = @{ LogName = $logName; Error = $_.Exception.Message }
                }
            }
            
            # Generate summary
            $totalEvents = ($computerResult.LogAnalysis.Values | Where-Object TotalEvents | Measure-Object TotalEvents -Sum).Sum
            $totalCritical = ($computerResult.LogAnalysis.Values | Where-Object Critical | Measure-Object Critical -Sum).Sum
            $totalErrors = ($computerResult.LogAnalysis.Values | Where-Object Error | Measure-Object Error -Sum).Sum
            $totalWarnings = ($computerResult.LogAnalysis.Values | Where-Object Warning | Measure-Object Warning -Sum).Sum
            
            $computerResult.Summary = @{
                TotalEvents = $totalEvents
                Critical = $totalCritical
                Errors = $totalErrors
                Warnings = $totalWarnings
                AlertCount = $computerResult.Alerts.Count
                PatternCount = $computerResult.ErrorPatterns.Count
                HealthScore = [Math]::Max(100 - ($totalCritical * 10) - ($totalErrors * 2) - $computerResult.Alerts.Count * 5, 0)
                Status = if ($computerResult.Alerts.Count -eq 0) { "Healthy" } 
                        elseif ($computerResult.Alerts | Where-Object Severity -eq "Critical") { "Critical" }
                        elseif ($computerResult.Alerts | Where-Object Severity -eq "High") { "Warning" }
                        else { "Attention" }
            }
            
            $results += $computerResult
            
            Write-Host "  ✓ Analysis completed - Health: $($computerResult.Summary.Status) (Score: $($computerResult.Summary.HealthScore))" -ForegroundColor Green
            Write-Host "    Events: $totalEvents total, $totalCritical critical, $totalErrors errors, $($computerResult.Alerts.Count) alerts" -ForegroundColor White
            
        } catch {
            Write-Error "Failed to analyze $computer : $_"
            $results += [PSCustomObject]@{
                ComputerName = $computer
                Status = "Error"
                Error = $_.Exception.Message
                AnalysisTime = Get-Date
            }
        }
        
        Write-Progress -Activity "Event Log Analysis: $computer" -Completed
    }
}

end {
    # Export detailed results
    Write-Host "`nExporting analysis results..." -ForegroundColor Cyan
    
    $results | Export-Clixml -Path "$ExportPath\EventLogAnalysisResults.xml"
    $results | ConvertTo-Json -Depth 6 | Out-File "$ExportPath\EventLogAnalysisResults.json"
    
    # Create summary report
    $summary = $results | Where-Object Status -eq "Success" | Select-Object ComputerName,
        @{N="HealthStatus";E={$_.Summary.Status}},
        @{N="HealthScore";E={$_.Summary.HealthScore}},
        @{N="TotalEvents";E={$_.Summary.TotalEvents}},
        @{N="Critical";E={$_.Summary.Critical}},
        @{N="Errors";E={$_.Summary.Errors}},
        @{N="Warnings";E={$_.Summary.Warnings}},
        @{N="Alerts";E={$_.Summary.AlertCount}},
        @{N="Patterns";E={$_.Summary.PatternCount}}
    
    $summary | Export-Csv -Path "$ExportPath\EventLogSummary.csv" -NoTypeInformation
    
    # Export alerts separately
    $allAlerts = $results | ForEach-Object { $_.Alerts } | Where-Object { $_ }
    if ($allAlerts) {
        $allAlerts | Export-Csv -Path "$ExportPath\CriticalAlerts.csv" -NoTypeInformation
    }
    
    # Display summary
    Write-Host "`n" + "="*100 -ForegroundColor Cyan
    Write-Host "EVENT LOG ANALYSIS SUMMARY" -ForegroundColor Cyan
    Write-Host "="*100 -ForegroundColor Cyan
    
    $summary | Format-Table -AutoSize
    
    # Health distribution
    $healthDistribution = $summary | Group-Object HealthStatus | Sort-Object Name
    Write-Host "System Health Distribution:" -ForegroundColor Yellow
    $healthDistribution | ForEach-Object {
        $color = switch ($_.Name) {
            "Healthy" { "Green" }
            "Attention" { "Yellow" }
            "Warning" { "DarkYellow" }
            "Critical" { "Red" }
            default { "White" }
        }
        Write-Host "  $($_.Name): $($_.Count) systems" -ForegroundColor $color
    }
    
    # Critical alerts summary
    if ($allAlerts) {
        Write-Host "`nCritical Alerts Summary:" -ForegroundColor Red
        $alertsByType = $allAlerts | Group-Object AlertType | Sort-Object Count -Descending
        $alertsByType | ForEach-Object {
            Write-Host "  $($_.Name): $($_.Count) alerts" -ForegroundColor Yellow
        }
        
        # Email alerts if configured
        if ($AlertEmail -and ($allAlerts | Where-Object Severity -in @("Critical", "High"))) {
            $criticalAlerts = $allAlerts | Where-Object Severity -in @("Critical", "High")
            $alertBody = "Critical Event Log Alerts Detected`n`n"
            $alertBody += ($criticalAlerts | ForEach-Object { "$($_.Computer): $($_.Message)" }) -join "`n"
            
            # Note: Email sending would require Send-MailMessage or equivalent
            Write-Host "  ⚠ $($criticalAlerts.Count) critical alerts would be sent to $AlertEmail" -ForegroundColor Red
            $alertBody | Out-File "$ExportPath\CriticalAlertsEmail.txt"
        }
    }
    
    # Top error patterns
    $allPatterns = $results | ForEach-Object { $_.ErrorPatterns } | Group-Object Pattern | Sort-Object Count -Descending | Select-Object -First 5
    if ($allPatterns) {
        Write-Host "`nMost Common Error Patterns:" -ForegroundColor Yellow
        $allPatterns | ForEach-Object {
            Write-Host "  • $($_.Name): $($_.Count) systems affected" -ForegroundColor Red
        }
    }
    
    # Overall statistics
    $overallStats = $summary | Measure-Object TotalEvents, Critical, Errors, Warnings -Sum
    Write-Host "`nOverall Statistics:" -ForegroundColor Yellow
    Write-Host "  Total Events Analyzed: $($overallStats[0].Sum)" -ForegroundColor White
    Write-Host "  Critical Events: $($overallStats[1].Sum)" -ForegroundColor Red
    Write-Host "  Error Events: $($overallStats[2].Sum)" -ForegroundColor Red
    Write-Host "  Warning Events: $($overallStats[3].Sum)" -ForegroundColor Yellow
    
    $avgHealthScore = [math]::Round(($summary | Measure-Object HealthScore -Average).Average, 1)
    Write-Host "  Average Health Score: $avgHealthScore%" -ForegroundColor $(if ($avgHealthScore -ge 80) {"Green"} elseif ($avgHealthScore -ge 60) {"Yellow"} else {"Red"})
    
    Write-Host "`nDetailed analysis results saved to: $ExportPath" -ForegroundColor Green
    Write-Host "Event log analysis completed for $($results.Count) systems" -ForegroundColor Cyan
}