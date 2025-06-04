<#
.SYNOPSIS
Comprehensive network connectivity testing and diagnostics

.DESCRIPTION
Performs advanced network connectivity testing including ping, port testing,
DNS resolution, traceroute analysis, and network performance measurements.
Provides detailed diagnostics for troubleshooting network issues.

.PARAMETER TargetHosts
Target hosts/IPs to test connectivity. Can include hostnames, IPs, or URLs.

.PARAMETER Ports
Specific ports to test (comma-separated). Common ports tested by default.

.PARAMETER IncludePerformance
Include network performance testing (latency, throughput)

.PARAMETER ExportPath
Path to save connectivity test results

.PARAMETER Continuous
Run continuous monitoring (Ctrl+C to stop)

.PARAMETER Interval
Interval between continuous tests in seconds. Default: 60

.EXAMPLE
Test-NetworkConnectivity -TargetHosts "google.com","8.8.8.8","microsoft.com"
Tests basic connectivity to multiple targets

.EXAMPLE
Test-NetworkConnectivity -TargetHosts "server01.domain.com" -Ports 80,443,3389 -IncludePerformance
Comprehensive testing including specific ports and performance metrics

.NOTES
Author: Enterprise PowerShell Collection
Requires: Administrative privileges for some tests
Version: 1.0
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory, ValueFromPipeline = $true)]
    [string[]]$TargetHosts,
    
    [Parameter()]
    [int[]]$Ports = @(53, 80, 443, 3389, 5985, 5986),
    
    [Parameter()]
    [switch]$IncludePerformance,
    
    [Parameter()]
    [string]$ExportPath,
    
    [Parameter()]
    [switch]$Continuous,
    
    [Parameter()]
    [int]$Interval = 60
)

begin {
    Write-Host "Starting Network Connectivity Testing..." -ForegroundColor Cyan
    
    # Setup export directory
    if ($ExportPath -and -not (Test-Path $ExportPath)) {
        New-Item -Path $ExportPath -ItemType Directory -Force | Out-Null
    }
    
    # Common service ports for reference
    $commonPorts = @{
        53 = "DNS"
        80 = "HTTP"
        443 = "HTTPS"
        25 = "SMTP"
        110 = "POP3"
        143 = "IMAP"
        993 = "IMAPS"
        995 = "POP3S"
        21 = "FTP"
        22 = "SSH"
        23 = "Telnet"
        3389 = "RDP"
        5985 = "WinRM HTTP"
        5986 = "WinRM HTTPS"
        135 = "RPC"
        445 = "SMB"
        389 = "LDAP"
        636 = "LDAPS"
    }
    
    $results = @()
    $testCount = 0
}

process {
    do {
        $testCount++
        $testTime = Get-Date
        
        if ($Continuous) {
            Write-Host "`n--- Test Run #$testCount at $testTime ---" -ForegroundColor Yellow
        }
        
        foreach ($target in $TargetHosts) {
            Write-Host "`nTesting connectivity to: $target" -ForegroundColor Yellow
            
            try {
                $targetResult = [PSCustomObject]@{
                    Target = $target
                    TestTime = $testTime
                    TestNumber = $testCount
                    BasicConnectivity = $null
                    DNSResolution = $null
                    PortTests = @()
                    Performance = $null
                    Traceroute = $null
                    Status = "Unknown"
                    Issues = @()
                    Recommendations = @()
                }
                
                # Basic ICMP connectivity test
                Write-Progress -Activity "Network Test: $target" -Status "Testing ICMP connectivity" -PercentComplete 10
                
                $pingResults = Test-Connection -ComputerName $target -Count 4 -ErrorAction SilentlyContinue
                if ($pingResults) {
                    $avgLatency = [math]::Round(($pingResults | Measure-Object ResponseTime -Average).Average, 2)
                    $packetLoss = [math]::Round(((4 - $pingResults.Count) / 4) * 100, 2)
                    
                    $targetResult.BasicConnectivity = @{
                        Success = $true
                        PacketsSent = 4
                        PacketsReceived = $pingResults.Count
                        PacketLoss = $packetLoss
                        AverageLatency = $avgLatency
                        MinLatency = ($pingResults | Measure-Object ResponseTime -Minimum).Minimum
                        MaxLatency = ($pingResults | Measure-Object ResponseTime -Maximum).Maximum
                    }
                    
                    if ($packetLoss -gt 0) {
                        $targetResult.Issues += "Packet loss detected: $packetLoss%"
                    }
                    if ($avgLatency -gt 100) {
                        $targetResult.Issues += "High latency: $avgLatency ms"
                        $targetResult.Recommendations += "Investigate network path for latency issues"
                    }
                } else {
                    $targetResult.BasicConnectivity = @{ Success = $false; Error = "ICMP ping failed" }
                    $targetResult.Issues += "ICMP connectivity failed"
                }
                
                # DNS Resolution test
                Write-Progress -Activity "Network Test: $target" -Status "Testing DNS resolution" -PercentComplete 25
                
                try {
                    $dnsResult = Resolve-DnsName -Name $target -ErrorAction Stop
                    $resolvedIPs = $dnsResult | Where-Object Section -eq "Answer" | Select-Object -ExpandProperty IPAddress
                    
                    $targetResult.DNSResolution = @{
                        Success = $true
                        ResolvedIPs = $resolvedIPs
                        ResponseTime = (Measure-Command { Resolve-DnsName -Name $target -ErrorAction SilentlyContinue }).TotalMilliseconds
                        TTL = ($dnsResult | Select-Object -First 1).TTL
                    }
                } catch {
                    $targetResult.DNSResolution = @{ Success = $false; Error = $_.Exception.Message }
                    $targetResult.Issues += "DNS resolution failed"
                    $targetResult.Recommendations += "Check DNS server configuration and network connectivity"
                }
                
                # Port connectivity tests
                Write-Progress -Activity "Network Test: $target" -Status "Testing port connectivity" -PercentComplete 50
                
                foreach ($port in $Ports) {
                    $portTest = @{
                        Port = $port
                        Service = if ($commonPorts.ContainsKey($port)) { $commonPorts[$port] } else { "Unknown" }
                        Success = $false
                        ResponseTime = $null
                        Error = $null
                    }
                    
                    try {
                        $tcpClient = New-Object System.Net.Sockets.TcpClient
                        $connectTime = Measure-Command {
                            $result = $tcpClient.BeginConnect($target, $port, $null, $null)
                            $success = $result.AsyncWaitHandle.WaitOne(5000, $false)
                            if ($success) {
                                $tcpClient.EndConnect($result)
                                $portTest.Success = $true
                            }
                        }
                        $tcpClient.Close()
                        
                        if ($portTest.Success) {
                            $portTest.ResponseTime = [math]::Round($connectTime.TotalMilliseconds, 2)
                        } else {
                            $portTest.Error = "Connection timeout"
                        }
                    } catch {
                        $portTest.Error = $_.Exception.Message
                    }
                    
                    $targetResult.PortTests += $portTest
                    
                    $status = if ($portTest.Success) { "✓" } else { "✗" }
                    Write-Host "    Port $port ($($portTest.Service)): $status" -ForegroundColor $(if ($portTest.Success) {"Green"} else {"Red"})
                }
                
                # Performance testing (if requested)
                if ($IncludePerformance) {
                    Write-Progress -Activity "Network Test: $target" -Status "Performance testing" -PercentComplete 75
                    
                    # Extended ping test for jitter analysis
                    $extendedPing = Test-Connection -ComputerName $target -Count 20 -ErrorAction SilentlyContinue
                    if ($extendedPing) {
                        $latencies = $extendedPing | Select-Object -ExpandProperty ResponseTime
                        $avgLatency = [math]::Round(($latencies | Measure-Object -Average).Average, 2)
                        $jitter = [math]::Round((($latencies | ForEach-Object { [math]::Abs($_ - $avgLatency) } | Measure-Object -Average).Average), 2)
                        
                        $targetResult.Performance = @{
                            ExtendedPingCount = 20
                            AverageLatency = $avgLatency
                            Jitter = $jitter
                            StandardDeviation = [math]::Round([math]::Sqrt((($latencies | ForEach-Object { [math]::Pow($_ - $avgLatency, 2) } | Measure-Object -Sum).Sum / $latencies.Count)), 2)
                        }
                        
                        if ($jitter -gt 10) {
                            $targetResult.Issues += "High network jitter: $jitter ms"
                            $targetResult.Recommendations += "Network stability issues detected - investigate network equipment"
                        }
                    }
                }
                
                # Traceroute (basic implementation)
                Write-Progress -Activity "Network Test: $target" -Status "Traceroute analysis" -PercentComplete 90
                
                try {
                    $tracert = tracert.exe $target
                    $hops = $tracert | Where-Object { $_ -match '^\s*\d+' } | Measure-Object
                    $targetResult.Traceroute = @{
                        Success = $true
                        HopCount = $hops.Count
                        Route = ($tracert | Where-Object { $_ -match '^\s*\d+' }) -join "; "
                    }
                    
                    if ($hops.Count -gt 15) {
                        $targetResult.Issues += "High hop count: $($hops.Count) hops"
                        $targetResult.Recommendations += "Review network routing - possible inefficient path"
                    }
                } catch {
                    $targetResult.Traceroute = @{ Success = $false; Error = "Traceroute failed" }
                }
                
                # Overall status assessment
                $successfulPorts = ($targetResult.PortTests | Where-Object Success -eq $true).Count
                $totalPorts = $targetResult.PortTests.Count
                
                $targetResult.Status = if ($targetResult.BasicConnectivity.Success -and $targetResult.DNSResolution.Success -and $successfulPorts -gt 0) {
                    if ($targetResult.Issues.Count -eq 0) { "Excellent" }
                    elseif ($targetResult.Issues.Count -le 2) { "Good" }
                    else { "Fair" }
                } elseif ($targetResult.BasicConnectivity.Success -or $successfulPorts -gt 0) {
                    "Limited"
                } else {
                    "Failed"
                }
                
                $results += $targetResult
                
                Write-Host "  ✓ Test completed - Status: $($targetResult.Status)" -ForegroundColor Green
                Write-Host "    ICMP: $(if ($targetResult.BasicConnectivity.Success) {'✓'} else {'✗'}) | DNS: $(if ($targetResult.DNSResolution.Success) {'✓'} else {'✗'}) | Ports: $successfulPorts/$totalPorts" -ForegroundColor White
                
            } catch {
                Write-Error "Failed to test $target : $_"
                $results += [PSCustomObject]@{
                    Target = $target
                    TestTime = $testTime
                    Status = "Error"
                    Error = $_.Exception.Message
                }
            }
            
            Write-Progress -Activity "Network Test: $target" -Completed
        }
        
        # Export results if path specified
        if ($ExportPath) {
            $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
            $results | Export-Csv -Path "$ExportPath\NetworkTest_$timestamp.csv" -NoTypeInformation
            $results | Export-Clixml -Path "$ExportPath\NetworkTest_$timestamp.xml"
        }
        
        # Display summary for this test run
        if ($Continuous -or $testCount -eq 1) {
            Write-Host "`n" + "-"*60 -ForegroundColor Cyan
            Write-Host "CONNECTIVITY TEST SUMMARY - Run #$testCount" -ForegroundColor Cyan
            Write-Host "-"*60 -ForegroundColor Cyan
            
            $summary = $results | Where-Object TestNumber -eq $testCount | Select-Object Target, Status,
                @{N="ICMP";E={if($_.BasicConnectivity.Success){"✓"}else{"✗"}}},
                @{N="DNS";E={if($_.DNSResolution.Success){"✓"}else{"✗"}}},
                @{N="AvgLatency";E={$_.BasicConnectivity.AverageLatency}},
                @{N="OpenPorts";E={($_.PortTests | Where-Object Success -eq $true).Count}},
                @{N="Issues";E={$_.Issues.Count}}
            
            $summary | Format-Table -AutoSize
            
            # Status distribution
            $statusDist = $results | Where-Object TestNumber -eq $testCount | Group-Object Status
            Write-Host "Status Distribution:" -ForegroundColor Yellow
            $statusDist | ForEach-Object {
                $color = switch ($_.Name) {
                    "Excellent" { "Green" }
                    "Good" { "Green" }
                    "Fair" { "Yellow" }
                    "Limited" { "DarkYellow" }
                    "Failed" { "Red" }
                    default { "White" }
                }
                Write-Host "  $($_.Name): $($_.Count)" -ForegroundColor $color
            }
        }
        
        if ($Continuous) {
            Write-Host "`nWaiting $Interval seconds for next test... (Ctrl+C to stop)" -ForegroundColor Gray
            Start-Sleep -Seconds $Interval
        }
        
    } while ($Continuous)
}

end {
    if (-not $Continuous) {
        # Final summary for single run
        Write-Host "`n" + "="*80 -ForegroundColor Cyan
        Write-Host "NETWORK CONNECTIVITY TEST RESULTS" -ForegroundColor Cyan
        Write-Host "="*80 -ForegroundColor Cyan
        
        $finalSummary = $results | Select-Object Target, Status,
            @{N="ICMP";E={if($_.BasicConnectivity.Success){"✓"}else{"✗"}}},
            @{N="DNS";E={if($_.DNSResolution.Success){"✓"}else{"✗"}}},
            @{N="AvgLatency";E={$_.BasicConnectivity.AverageLatency}},
            @{N="OpenPorts";E={($_.PortTests | Where-Object Success -eq $true).Count}},
            @{N="Issues";E={$_.Issues.Count}}
        
        $finalSummary | Format-Table -AutoSize
        
        # Issue summary
        $allIssues = $results | ForEach-Object { $_.Issues } | Group-Object | Sort-Object Count -Descending
        if ($allIssues) {
            Write-Host "Common Issues:" -ForegroundColor Yellow
            $allIssues | Select-Object -First 5 | ForEach-Object {
                Write-Host "  • $($_.Name) ($($_.Count) targets)" -ForegroundColor Red
            }
        }
        
        # Recommendations
        $allRecommendations = $results | ForEach-Object { $_.Recommendations } | Group-Object | Sort-Object Count -Descending
        if ($allRecommendations) {
            Write-Host "`nRecommendations:" -ForegroundColor Yellow
            $allRecommendations | Select-Object -First 5 | ForEach-Object {
                Write-Host "  • $($_.Name) ($($_.Count) targets)" -ForegroundColor Cyan
            }
        }
        
        if ($ExportPath) {
            Write-Host "`nDetailed results saved to: $ExportPath" -ForegroundColor Green
        }
    }
    
    Write-Host "`nNetwork connectivity testing completed" -ForegroundColor Cyan
    if ($Continuous) {
        Write-Host "Total test runs: $testCount" -ForegroundColor White
    }
}