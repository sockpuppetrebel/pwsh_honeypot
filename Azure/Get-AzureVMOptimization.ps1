#Requires -Modules Az.Accounts, Az.Compute, Az.Monitor

<#
.SYNOPSIS
    Analyzes Azure VMs for optimization opportunities including rightsizing and cost reduction
.DESCRIPTION
    Evaluates VM performance metrics, utilization patterns, and costs to identify
    optimization opportunities like rightsizing, deallocating unused VMs, and cost savings.
.PARAMETER SubscriptionId
    Target Azure subscription ID
.PARAMETER ResourceGroupName
    Specific resource group to analyze (optional)
.PARAMETER DaysToAnalyze
    Number of days of metrics to analyze (default: 14)
.PARAMETER ExportToExcel
    Export detailed results to Excel format
.PARAMETER IncludeMetrics
    Include detailed performance metrics in analysis
.EXAMPLE
    .\Get-AzureVMOptimization.ps1 -SubscriptionId "12345678-1234-1234-1234-123456789012" -DaysToAnalyze 30 -IncludeMetrics -ExportToExcel
#>

param(
    [Parameter(Mandatory = $false)]
    [string]$SubscriptionId,
    
    [Parameter(Mandatory = $false)]
    [string]$ResourceGroupName,
    
    [Parameter(Mandatory = $false)]
    [int]$DaysToAnalyze = 14,
    
    [Parameter(Mandatory = $false)]
    [switch]$ExportToExcel,
    
    [Parameter(Mandatory = $false)]
    [switch]$IncludeMetrics
)

# Connect to Azure
try {
    $context = Get-AzContext
    if (-not $context) {
        Write-Host "Connecting to Azure..." -ForegroundColor Yellow
        Connect-AzAccount
    }
    
    if ($SubscriptionId) {
        Set-AzContext -SubscriptionId $SubscriptionId
    }
    
    $currentSubscription = Get-AzContext
    Write-Host "Connected to subscription: $($currentSubscription.Subscription.Name)" -ForegroundColor Green
}
catch {
    Write-Error "Failed to connect to Azure: $($_.Exception.Message)"
    exit 1
}

Write-Host "=== AZURE VM OPTIMIZATION ANALYSIS ===" -ForegroundColor Cyan

# Get VMs
Write-Host "Retrieving virtual machines..." -ForegroundColor Yellow
if ($ResourceGroupName) {
    $vms = Get-AzVM -ResourceGroupName $ResourceGroupName
} else {
    $vms = Get-AzVM
}

if ($vms.Count -eq 0) {
    Write-Host "No VMs found" -ForegroundColor Yellow
    exit 0
}

Write-Host "Analyzing $($vms.Count) virtual machines..." -ForegroundColor Yellow

$vmOptimizationData = @()
$vmCounter = 0
$endTime = Get-Date
$startTime = $endTime.AddDays(-$DaysToAnalyze)

# VM size cost mapping (approximate monthly costs in USD)
$vmSizeCosts = @{
    "Standard_B1s" = 7.88
    "Standard_B1ms" = 15.77
    "Standard_B2s" = 31.54
    "Standard_B2ms" = 63.07
    "Standard_B4ms" = 126.14
    "Standard_B8ms" = 252.29
    "Standard_D2s_v3" = 96.36
    "Standard_D4s_v3" = 192.72
    "Standard_D8s_v3" = 385.44
    "Standard_D16s_v3" = 770.88
    "Standard_D32s_v3" = 1541.76
    "Standard_E2s_v3" = 121.69
    "Standard_E4s_v3" = 243.38
    "Standard_E8s_v3" = 486.77
    "Standard_E16s_v3" = 973.53
    "Standard_E32s_v3" = 1947.06
    "Standard_F2s_v2" = 83.22
    "Standard_F4s_v2" = 166.44
    "Standard_F8s_v2" = 332.88
    "Standard_F16s_v2" = 665.76
    "Standard_F32s_v2" = 1331.52
}

foreach ($vm in $vms) {
    $vmCounter++
    Write-Progress -Activity "Analyzing VMs" -Status "$($vm.Name) ($vmCounter/$($vms.Count))" -PercentComplete (($vmCounter / $vms.Count) * 100)
    
    try {
        # Get VM instance details
        $vmInstance = Get-AzVM -ResourceGroupName $vm.ResourceGroupName -Name $vm.Name -Status
        $vmSize = Get-AzVMSize -Location $vm.Location | Where-Object Name -eq $vm.HardwareProfile.VmSize
        
        # Calculate estimated monthly cost
        $estimatedMonthlyCost = if ($vmSizeCosts.ContainsKey($vm.HardwareProfile.VmSize)) {
            $vmSizeCosts[$vm.HardwareProfile.VmSize]
        } else {
            # Estimate based on cores and memory
            $cores = $vmSize.NumberOfCores
            $memoryGB = $vmSize.MemoryInMB / 1024
            ($cores * 20) + ($memoryGB * 5)  # Rough estimation
        }
        
        # Get power state
        $powerState = ($vmInstance.Statuses | Where-Object Code -like "PowerState/*").DisplayStatus
        $isRunning = $powerState -eq "VM running"
        
        # Initialize metrics variables
        $avgCpuPercent = $null
        $avgMemoryPercent = $null
        $maxCpuPercent = $null
        $maxMemoryPercent = $null
        $networkInMB = $null
        $networkOutMB = $null
        $diskReadMB = $null
        $diskWriteMB = $null
        
        # Get performance metrics if requested and VM is running
        if ($IncludeMetrics -and $isRunning) {
            try {
                # CPU metrics
                $cpuMetrics = Get-AzMetric -ResourceId $vm.Id -MetricName "Percentage CPU" -StartTime $startTime -EndTime $endTime -TimeGrain 01:00:00 -ErrorAction SilentlyContinue
                if ($cpuMetrics.Data) {
                    $cpuValues = $cpuMetrics.Data | Where-Object Average -ne $null | ForEach-Object { $_.Average }
                    if ($cpuValues) {
                        $avgCpuPercent = [math]::Round(($cpuValues | Measure-Object -Average).Average, 2)
                        $maxCpuPercent = [math]::Round(($cpuValues | Measure-Object -Maximum).Maximum, 2)
                    }
                }
                
                # Network metrics
                $networkInMetrics = Get-AzMetric -ResourceId $vm.Id -MetricName "Network In Total" -StartTime $startTime -EndTime $endTime -TimeGrain 01:00:00 -ErrorAction SilentlyContinue
                if ($networkInMetrics.Data) {
                    $networkInBytes = ($networkInMetrics.Data | Where-Object Total -ne $null | Measure-Object Total -Sum).Sum
                    $networkInMB = [math]::Round($networkInBytes / 1MB, 2)
                }
                
                $networkOutMetrics = Get-AzMetric -ResourceId $vm.Id -MetricName "Network Out Total" -StartTime $startTime -EndTime $endTime -TimeGrain 01:00:00 -ErrorAction SilentlyContinue
                if ($networkOutMetrics.Data) {
                    $networkOutBytes = ($networkOutMetrics.Data | Where-Object Total -ne $null | Measure-Object Total -Sum).Sum
                    $networkOutMB = [math]::Round($networkOutBytes / 1MB, 2)
                }
                
                # Disk metrics
                $diskReadMetrics = Get-AzMetric -ResourceId $vm.Id -MetricName "Disk Read Bytes" -StartTime $startTime -EndTime $endTime -TimeGrain 01:00:00 -ErrorAction SilentlyContinue
                if ($diskReadMetrics.Data) {
                    $diskReadBytes = ($diskReadMetrics.Data | Where-Object Total -ne $null | Measure-Object Total -Sum).Sum
                    $diskReadMB = [math]::Round($diskReadBytes / 1MB, 2)
                }
                
                $diskWriteMetrics = Get-AzMetric -ResourceId $vm.Id -MetricName "Disk Write Bytes" -StartTime $startTime -EndTime $endTime -TimeGrain 01:00:00 -ErrorAction SilentlyContinue
                if ($diskWriteMetrics.Data) {
                    $diskWriteBytes = ($diskWriteMetrics.Data | Where-Object Total -ne $null | Measure-Object Total -Sum).Sum
                    $diskWriteMB = [math]::Round($diskWriteBytes / 1MB, 2)
                }
            }
            catch {
                Write-Host "Warning: Could not retrieve metrics for $($vm.Name): $($_.Exception.Message)" -ForegroundColor Yellow
            }
        }
        
        # Analyze optimization opportunities
        $optimizationRecommendations = @()
        $potentialSavings = 0
        
        # Check if VM is stopped/deallocated
        if (-not $isRunning) {
            if ($powerState -eq "VM deallocated") {
                $optimizationRecommendations += "VM is deallocated - no compute costs"
            } else {
                $optimizationRecommendations += "VM is stopped but not deallocated - still incurring costs"
                $potentialSavings += $estimatedMonthlyCost * 0.8  # 80% savings from deallocation
            }
        }
        
        # CPU utilization analysis
        if ($avgCpuPercent -ne $null) {
            if ($avgCpuPercent -lt 5) {
                $optimizationRecommendations += "Very low CPU utilization ($avgCpuPercent%) - consider smaller VM size"
                $potentialSavings += $estimatedMonthlyCost * 0.5  # 50% savings from downsizing
            } elseif ($avgCpuPercent -lt 15) {
                $optimizationRecommendations += "Low CPU utilization ($avgCpuPercent%) - consider downsizing"
                $potentialSavings += $estimatedMonthlyCost * 0.3  # 30% savings from downsizing
            } elseif ($avgCpuPercent -gt 80) {
                $optimizationRecommendations += "High CPU utilization ($avgCpuPercent%) - consider larger VM size"
            }
        }
        
        # VM size analysis
        if ($vmSize) {
            # Check for over-provisioned memory (assuming 4GB+ per core is typical)
            $memoryPerCore = $vmSize.MemoryInMB / $vmSize.NumberOfCores / 1024
            if ($memoryPerCore -gt 8) {
                $optimizationRecommendations += "High memory-to-core ratio ($([math]::Round($memoryPerCore, 1))GB/core) - consider compute-optimized VM"
            }
        }
        
        # Check for premium disks on development VMs
        $osDisks = Get-AzDisk -ResourceGroupName $vm.ResourceGroupName | Where-Object { $_.ManagedBy -eq $vm.Id -and $_.Name -like "*OsDisk*" }
        foreach ($disk in $osDisks) {
            if ($disk.Sku.Name -like "*Premium*" -and $vm.Tags.Environment -eq "Development") {
                $optimizationRecommendations += "Premium disk on development VM - consider Standard SSD"
                $potentialSavings += 50  # Estimated monthly savings
            }
        }
        
        # Calculate optimization score
        $optimizationScore = 100
        if ($optimizationRecommendations.Count -gt 0) {
            $optimizationScore -= ($optimizationRecommendations.Count * 20)
        }
        if ($avgCpuPercent -and $avgCpuPercent -lt 10) {
            $optimizationScore -= 30
        }
        $optimizationScore = [math]::Max($optimizationScore, 0)
        
        $vmOptimization = [PSCustomObject]@{
            VMName = $vm.Name
            ResourceGroup = $vm.ResourceGroupName
            Location = $vm.Location
            PowerState = $powerState
            IsRunning = $isRunning
            
            # VM Configuration
            VMSize = $vm.HardwareProfile.VmSize
            Cores = $vmSize.NumberOfCores
            MemoryGB = [math]::Round($vmSize.MemoryInMB / 1024, 2)
            MaxDataDiskCount = $vmSize.MaxDataDiskCount
            OSType = $vm.StorageProfile.OSDisk.OSType
            
            # Cost Analysis
            EstimatedMonthlyCostUSD = [math]::Round($estimatedMonthlyCost, 2)
            PotentialMonthlySavingsUSD = [math]::Round($potentialSavings, 2)
            SavingsPercentage = if ($estimatedMonthlyCost -gt 0) { 
                [math]::Round(($potentialSavings / $estimatedMonthlyCost) * 100, 1) 
            } else { 0 }
            
            # Performance Metrics
            AvgCpuPercent = $avgCpuPercent
            MaxCpuPercent = $maxCpuPercent
            AvgMemoryPercent = $avgMemoryPercent
            MaxMemoryPercent = $maxMemoryPercent
            NetworkInMB = $networkInMB
            NetworkOutMB = $networkOutMB
            DiskReadMB = $diskReadMB
            DiskWriteMB = $diskWriteMB
            
            # Analysis
            OptimizationScore = $optimizationScore
            OptimizationRecommendations = ($optimizationRecommendations -join "; ")
            HasOptimizationOpportunity = ($optimizationRecommendations.Count -gt 0)
            
            # Categories
            IsUnderutilized = ($avgCpuPercent -and $avgCpuPercent -lt 15)
            IsOverutilized = ($avgCpuPercent -and $avgCpuPercent -gt 80)
            IsIdle = (-not $isRunning)
            IsExpensive = ($estimatedMonthlyCost -gt 500)
            
            # Tags
            Environment = $vm.Tags.Environment
            Department = $vm.Tags.Department
            Owner = $vm.Tags.Owner
            CostCenter = $vm.Tags.CostCenter
        }
        
        $vmOptimizationData += $vmOptimization
    }
    catch {
        Write-Host "Error analyzing VM $($vm.Name): $($_.Exception.Message)" -ForegroundColor Yellow
        
        # Add basic info even if detailed analysis fails
        $basicOptimization = [PSCustomObject]@{
            VMName = $vm.Name
            ResourceGroup = $vm.ResourceGroupName
            Location = $vm.Location
            VMSize = $vm.HardwareProfile.VmSize
            OptimizationScore = 0
            OptimizationRecommendations = "Analysis failed: $($_.Exception.Message)"
            HasOptimizationOpportunity = $true
        }
        
        $vmOptimizationData += $basicOptimization
    }
}

Write-Progress -Activity "Analyzing VMs" -Completed

# Generate summary report
Write-Host "`n--- VM OPTIMIZATION SUMMARY ---" -ForegroundColor Cyan

$totalVMs = $vmOptimizationData.Count
$runningVMs = ($vmOptimizationData | Where-Object IsRunning -eq $true).Count
$idleVMs = ($vmOptimizationData | Where-Object IsIdle -eq $true).Count
$underutilizedVMs = ($vmOptimizationData | Where-Object IsUnderutilized -eq $true).Count
$overutilizedVMs = ($vmOptimizationData | Where-Object IsOverutilized -eq $true).Count
$vmsWithOpportunities = ($vmOptimizationData | Where-Object HasOptimizationOpportunity -eq $true).Count

Write-Host "Total VMs: $totalVMs" -ForegroundColor White
Write-Host "Running VMs: $runningVMs" -ForegroundColor Green
Write-Host "Idle VMs: $idleVMs" -ForegroundColor Yellow
Write-Host "Underutilized VMs: $underutilizedVMs" -ForegroundColor Yellow
Write-Host "Overutilized VMs: $overutilizedVMs" -ForegroundColor Red
Write-Host "VMs with optimization opportunities: $vmsWithOpportunities" -ForegroundColor Cyan

# Cost analysis
$totalMonthlyCost = ($vmOptimizationData | Where-Object EstimatedMonthlyCostUSD | Measure-Object EstimatedMonthlyCostUSD -Sum).Sum
$totalPotentialSavings = ($vmOptimizationData | Where-Object PotentialMonthlySavingsUSD | Measure-Object PotentialMonthlySavingsUSD -Sum).Sum

Write-Host "`nCost Analysis:" -ForegroundColor Yellow
Write-Host "Total Estimated Monthly Cost: `$$([math]::Round($totalMonthlyCost, 2))" -ForegroundColor White
Write-Host "Total Potential Monthly Savings: `$$([math]::Round($totalPotentialSavings, 2))" -ForegroundColor Green
if ($totalMonthlyCost -gt 0) {
    Write-Host "Potential Savings Percentage: $([math]::Round(($totalPotentialSavings / $totalMonthlyCost) * 100, 1))%" -ForegroundColor Green
}

# Top optimization opportunities
if ($vmsWithOpportunities -gt 0) {
    Write-Host "`n--- TOP OPTIMIZATION OPPORTUNITIES ---" -ForegroundColor Cyan
    $vmOptimizationData | Where-Object HasOptimizationOpportunity -eq $true | 
        Sort-Object PotentialMonthlySavingsUSD -Descending | 
        Select-Object -First 10 | 
        Select-Object VMName, VMSize, EstimatedMonthlyCostUSD, PotentialMonthlySavingsUSD, OptimizationRecommendations | 
        Format-Table -Wrap
}

# Underutilized VMs
if ($underutilizedVMs -gt 0) {
    Write-Host "`n--- UNDERUTILIZED VMS ---" -ForegroundColor Yellow
    $vmOptimizationData | Where-Object IsUnderutilized -eq $true | 
        Sort-Object AvgCpuPercent | 
        Select-Object VMName, VMSize, AvgCpuPercent, MaxCpuPercent, EstimatedMonthlyCostUSD | 
        Format-Table
}

# Idle VMs
if ($idleVMs -gt 0) {
    Write-Host "`n--- IDLE VMS ---" -ForegroundColor Red
    $vmOptimizationData | Where-Object IsIdle -eq $true | 
        Select-Object VMName, PowerState, VMSize, EstimatedMonthlyCostUSD, Environment | 
        Format-Table
}

# VM size distribution
Write-Host "`n--- VM SIZE DISTRIBUTION ---" -ForegroundColor Yellow
$vmOptimizationData | Group-Object VMSize | Sort-Object Count -Descending | 
    Select-Object -First 10 | 
    ForEach-Object {
        $avgCost = ($_.Group | Measure-Object EstimatedMonthlyCostUSD -Average).Average
        [PSCustomObject]@{
            VMSize = $_.Name
            Count = $_.Count
            AvgMonthlyCost = [math]::Round($avgCost, 2)
        }
    } | Format-Table

# Recommendations
Write-Host "`n--- OPTIMIZATION RECOMMENDATIONS ---" -ForegroundColor Cyan

$recommendations = @()

if ($idleVMs -gt 0) {
    $recommendations += "Deallocate $idleVMs idle VMs to eliminate compute costs"
}

if ($underutilizedVMs -gt 0) {
    $recommendations += "Rightsize $underutilizedVMs underutilized VMs to reduce costs"
}

if ($overutilizedVMs -gt 0) {
    $recommendations += "Scale up $overutilizedVMs overutilized VMs to improve performance"
}

$untaggedVMs = ($vmOptimizationData | Where-Object { -not $_.Environment -and -not $_.Department }).Count
if ($untaggedVMs -gt 0) {
    $recommendations += "Add cost center tags to $untaggedVMs VMs for better cost allocation"
}

$expensiveVMs = ($vmOptimizationData | Where-Object IsExpensive -eq $true).Count
if ($expensiveVMs -gt 0) {
    $recommendations += "Review $expensiveVMs expensive VMs (>`$500/month) for optimization"
}

if ($recommendations.Count -gt 0) {
    $recommendations | ForEach-Object { Write-Host "• $_" -ForegroundColor Cyan }
} else {
    Write-Host "• VM utilization appears optimized!" -ForegroundColor Green
}

# Export results
if ($ExportToExcel) {
    $timestamp = Get-Date -Format "yyyy-MM-dd-HHmm"
    $outputPath = ".\Azure-VM-Optimization-$timestamp.xlsx"
    
    if (Get-Module -ListAvailable -Name ImportExcel) {
        # Main optimization data
        $vmOptimizationData | Export-Excel -Path $outputPath -WorksheetName "VM Optimization" -AutoSize -FreezeTopRow
        
        # Summary statistics
        $summaryData = @(
            [PSCustomObject]@{ Metric = "Total VMs"; Value = $totalVMs }
            [PSCustomObject]@{ Metric = "Running VMs"; Value = $runningVMs }
            [PSCustomObject]@{ Metric = "Idle VMs"; Value = $idleVMs }
            [PSCustomObject]@{ Metric = "Underutilized VMs"; Value = $underutilizedVMs }
            [PSCustomObject]@{ Metric = "Overutilized VMs"; Value = $overutilizedVMs }
            [PSCustomObject]@{ Metric = "VMs with Opportunities"; Value = $vmsWithOpportunities }
            [PSCustomObject]@{ Metric = "Total Monthly Cost"; Value = "`$$([math]::Round($totalMonthlyCost, 2))" }
            [PSCustomObject]@{ Metric = "Potential Monthly Savings"; Value = "`$$([math]::Round($totalPotentialSavings, 2))" }
        )
        
        $summaryData | Export-Excel -Path $outputPath -WorksheetName "Summary" -AutoSize
        
        # Optimization opportunities
        if ($vmsWithOpportunities -gt 0) {
            $vmOptimizationData | Where-Object HasOptimizationOpportunity -eq $true | 
                Export-Excel -Path $outputPath -WorksheetName "Optimization Opportunities" -AutoSize
        }
        
        Write-Host "`nDetailed report exported to: $outputPath" -ForegroundColor Green
    } else {
        Write-Host "ImportExcel module not available. Exporting to CSV..." -ForegroundColor Yellow
        $vmOptimizationData | Export-Csv -Path ".\Azure-VM-Optimization-$timestamp.csv" -NoTypeInformation
        Write-Host "Report exported to: .\Azure-VM-Optimization-$timestamp.csv" -ForegroundColor Green
    }
}

Write-Host "`n--- NEXT STEPS ---" -ForegroundColor Cyan
Write-Host "• Implement auto-shutdown schedules for development VMs" -ForegroundColor White
Write-Host "• Set up Azure Advisor recommendations monitoring" -ForegroundColor White
Write-Host "• Consider Azure Reserved Instances for predictable workloads" -ForegroundColor White
Write-Host "• Implement proper tagging strategy for cost allocation" -ForegroundColor White

Write-Host "VM optimization analysis complete" -ForegroundColor Green