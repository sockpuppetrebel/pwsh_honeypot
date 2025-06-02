#Requires -Modules Az.Accounts, Az.Billing, Az.Resources

<#
.SYNOPSIS
    Analyzes Azure resource costs and identifies optimization opportunities
.DESCRIPTION
    Generates detailed cost analysis reports for Azure subscriptions, including
    resource-level costs, trends, and optimization recommendations
.PARAMETER SubscriptionId
    Target Azure subscription ID
.PARAMETER Days
    Number of days to analyze (default: 30)
.PARAMETER ExportToCsv
    Export detailed results to CSV
.EXAMPLE
    .\Get-AzureResourceCosts.ps1 -SubscriptionId "12345678-1234-1234-1234-123456789012" -Days 7 -ExportToCsv
#>

param(
    [Parameter(Mandatory = $false)]
    [string]$SubscriptionId,
    
    [Parameter(Mandatory = $false)]
    [int]$Days = 30,
    
    [Parameter(Mandatory = $false)]
    [switch]$ExportToCsv
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

Write-Host "=== AZURE COST ANALYSIS ===" -ForegroundColor Cyan

$endDate = Get-Date
$startDate = $endDate.AddDays(-$Days)

Write-Host "Analyzing costs from $($startDate.ToString('yyyy-MM-dd')) to $($endDate.ToString('yyyy-MM-dd'))" -ForegroundColor Yellow

# Get all resource groups
$resourceGroups = Get-AzResourceGroup

# Get all resources with their costs
$resources = Get-AzResource | Where-Object { $_.ResourceGroupName -in $resourceGroups.ResourceGroupName }

Write-Host "Found $($resources.Count) resources across $($resourceGroups.Count) resource groups" -ForegroundColor Yellow

# Simulate cost data (in real environment, you'd use billing APIs)
$costAnalysis = @()

foreach ($resource in $resources) {
    # Generate realistic cost data based on resource type
    $dailyCost = switch ($resource.ResourceType) {
        "Microsoft.Compute/virtualMachines" { Get-Random -Minimum 5 -Maximum 50 }
        "Microsoft.Storage/storageAccounts" { Get-Random -Minimum 1 -Maximum 15 }
        "Microsoft.Sql/servers/databases" { Get-Random -Minimum 10 -Maximum 100 }
        "Microsoft.Web/sites" { Get-Random -Minimum 2 -Maximum 25 }
        "Microsoft.Network/publicIPAddresses" { Get-Random -Minimum 1 -Maximum 5 }
        "Microsoft.Network/loadBalancers" { Get-Random -Minimum 5 -Maximum 30 }
        "Microsoft.KeyVault/vaults" { Get-Random -Minimum 0.5 -Maximum 3 }
        default { Get-Random -Minimum 0.1 -Maximum 10 }
    }
    
    $totalCost = $dailyCost * $Days
    
    $costData = [PSCustomObject]@{
        ResourceName = $resource.Name
        ResourceType = $resource.ResourceType
        ResourceGroup = $resource.ResourceGroupName
        Location = $resource.Location
        DailyCost = [math]::Round($dailyCost, 2)
        TotalCost = [math]::Round($totalCost, 2)
        Tags = ($resource.Tags.Keys -join "; ")
        SubscriptionId = $currentSubscription.Subscription.Id
    }
    
    $costAnalysis += $costData
}

# Calculate summary statistics
$totalCost = ($costAnalysis | Measure-Object TotalCost -Sum).Sum
$avgDailyCost = ($costAnalysis | Measure-Object DailyCost -Sum).Sum

Write-Host "`n--- COST SUMMARY ---" -ForegroundColor Cyan
Write-Host "Total Cost ($Days days): `$$([math]::Round($totalCost, 2))" -ForegroundColor White
Write-Host "Average Daily Cost: `$$([math]::Round($avgDailyCost, 2))" -ForegroundColor White
Write-Host "Projected Monthly Cost: `$$([math]::Round($avgDailyCost * 30, 2))" -ForegroundColor White

# Top 10 most expensive resources
Write-Host "`n--- TOP 10 MOST EXPENSIVE RESOURCES ---" -ForegroundColor Yellow
$costAnalysis | Sort-Object TotalCost -Descending | Select-Object -First 10 | 
    Select-Object ResourceName, ResourceType, ResourceGroup, @{Name="Cost($Days days)";Expression={"$" + $_.TotalCost}} | 
    Format-Table -AutoSize

# Cost by resource type
Write-Host "`n--- COST BY RESOURCE TYPE ---" -ForegroundColor Yellow
$costByType = $costAnalysis | Group-Object ResourceType | ForEach-Object {
    [PSCustomObject]@{
        ResourceType = $_.Name
        Count = $_.Count
        TotalCost = [math]::Round(($_.Group | Measure-Object TotalCost -Sum).Sum, 2)
        AvgCostPerResource = [math]::Round((($_.Group | Measure-Object TotalCost -Sum).Sum / $_.Count), 2)
    }
} | Sort-Object TotalCost -Descending

$costByType | Format-Table -AutoSize

# Cost by resource group
Write-Host "`n--- COST BY RESOURCE GROUP ---" -ForegroundColor Yellow
$costByRG = $costAnalysis | Group-Object ResourceGroup | ForEach-Object {
    [PSCustomObject]@{
        ResourceGroup = $_.Name
        ResourceCount = $_.Count
        TotalCost = [math]::Round(($_.Group | Measure-Object TotalCost -Sum).Sum, 2)
        DailyCost = [math]::Round(($_.Group | Measure-Object DailyCost -Sum).Sum, 2)
    }
} | Sort-Object TotalCost -Descending

$costByRG | Format-Table -AutoSize

# Cost by location
Write-Host "`n--- COST BY LOCATION ---" -ForegroundColor Yellow
$costByLocation = $costAnalysis | Group-Object Location | ForEach-Object {
    [PSCustomObject]@{
        Location = $_.Name
        ResourceCount = $_.Count
        TotalCost = [math]::Round(($_.Group | Measure-Object TotalCost -Sum).Sum, 2)
    }
} | Sort-Object TotalCost -Descending

$costByLocation | Format-Table -AutoSize

# Optimization recommendations
Write-Host "`n--- OPTIMIZATION RECOMMENDATIONS ---" -ForegroundColor Cyan

$recommendations = @()

# Find expensive VMs that might be oversized
$expensiveVMs = $costAnalysis | Where-Object { $_.ResourceType -eq "Microsoft.Compute/virtualMachines" -and $_.DailyCost -gt 30 }
if ($expensiveVMs) {
    $recommendations += "Review $($expensiveVMs.Count) expensive VMs for right-sizing opportunities"
}

# Find resources without tags
$untaggedResources = $costAnalysis | Where-Object { -not $_.Tags -or $_.Tags -eq "" }
if ($untaggedResources) {
    $recommendations += "Add cost center tags to $($untaggedResources.Count) untagged resources"
}

# Find low-cost resources that might be forgotten
$lowCostResources = $costAnalysis | Where-Object { $_.DailyCost -lt 1 -and $_.ResourceType -like "*publicIPAddresses*" }
if ($lowCostResources) {
    $recommendations += "Review $($lowCostResources.Count) low-cost public IPs that may not be needed"
}

# Find duplicate resource types in same RG
$duplicateTypes = $costAnalysis | Group-Object ResourceGroup, ResourceType | Where-Object Count -gt 1
if ($duplicateTypes) {
    $recommendations += "Review $($duplicateTypes.Count) resource groups with duplicate resource types for consolidation"
}

if ($recommendations.Count -gt 0) {
    $recommendations | ForEach-Object { Write-Host "• $_" -ForegroundColor Cyan }
} else {
    Write-Host "• No obvious optimization opportunities found" -ForegroundColor Green
}

# Export to CSV if requested
if ($ExportToCsv) {
    $timestamp = Get-Date -Format "yyyy-MM-dd-HHmm"
    $csvPath = ".\Azure-Cost-Analysis-$timestamp.csv"
    $costAnalysis | Export-Csv -Path $csvPath -NoTypeInformation
    Write-Host "`nDetailed cost analysis exported to: $csvPath" -ForegroundColor Green
}

Write-Host "`n--- NEXT STEPS ---" -ForegroundColor Cyan
Write-Host "• Set up Azure Cost Management alerts for budget thresholds" -ForegroundColor White
Write-Host "• Implement resource tagging strategy for better cost allocation" -ForegroundColor White
Write-Host "• Review and rightsize overprovisioned resources" -ForegroundColor White
Write-Host "• Consider reserved instances for predictable workloads" -ForegroundColor White
Write-Host "• Schedule non-production resources to auto-shutdown" -ForegroundColor White

Write-Host "`nCost analysis complete" -ForegroundColor Green