#Requires -Modules Microsoft.Graph.Users

<#
.SYNOPSIS
    Analyzes M365 users with limited permissions (User.Read.All only)
.DESCRIPTION
    Works with just User.Read.All permission - shows user license assignments
#>

param(
    [Parameter(Mandatory = $false)]
    [switch]$ExportToCsv,
    
    [Parameter(Mandatory = $false)]
    [int]$InactiveDays = 90
)

# Connect to Microsoft Graph
Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Yellow
try {
    Connect-MgGraph -Scopes "User.Read.All" -NoWelcome
    $context = Get-MgContext
    Write-Host "Connected as: $($context.Account)" -ForegroundColor Green
} catch {
    Write-Error "Failed to connect: $_"
    exit 1
}

Write-Host "`n=== M365 USER LICENSE ANALYSIS ===" -ForegroundColor Cyan

# Get all users with license assignments
Write-Host "Retrieving users with licenses..." -ForegroundColor Yellow
try {
    $users = Get-MgUser -All -Property "Id,UserPrincipalName,DisplayName,AssignedLicenses,AccountEnabled,SignInActivity,CreatedDateTime,UserType" -Filter "assignedLicenses/`$count ne 0"
    Write-Host "Found $($users.Count) licensed users" -ForegroundColor Green
} catch {
    Write-Error "Failed to retrieve users: $_"
    exit 1
}

# Analyze users
$enabledUsers = $users | Where-Object AccountEnabled -eq $true
$disabledUsers = $users | Where-Object AccountEnabled -eq $false

# Find inactive users
$cutoffDate = (Get-Date).AddDays(-$InactiveDays)
$inactiveUsers = $enabledUsers | Where-Object {
    if ($_.SignInActivity.LastSignInDateTime) {
        (Get-Date $_.SignInActivity.LastSignInDateTime) -lt $cutoffDate
    } else {
        $true  # Never signed in
    }
}

# Summary
Write-Host "`n--- USER SUMMARY ---" -ForegroundColor Cyan
Write-Host "Total licensed users: $($users.Count)" -ForegroundColor White
Write-Host "Enabled users: $($enabledUsers.Count)" -ForegroundColor Green
Write-Host "Disabled users: $($disabledUsers.Count)" -ForegroundColor Yellow
Write-Host "Inactive users (>$InactiveDays days): $($inactiveUsers.Count)" -ForegroundColor Yellow

# Show disabled users with licenses
if ($disabledUsers.Count -gt 0) {
    Write-Host "`n--- DISABLED USERS WITH LICENSES ---" -ForegroundColor Yellow
    $disabledUsers | Select-Object UserPrincipalName, DisplayName, @{N='LicenseCount';E={$_.AssignedLicenses.Count}} | 
        Sort-Object LicenseCount -Descending | Format-Table -AutoSize
}

# Show inactive users
if ($inactiveUsers.Count -gt 0) {
    Write-Host "`n--- INACTIVE USERS (No sign-in for >$InactiveDays days) ---" -ForegroundColor Yellow
    $inactiveUsers | Select-Object UserPrincipalName, 
        @{N='LastSignIn';E={$_.SignInActivity.LastSignInDateTime}},
        @{N='DaysInactive';E={
            if ($_.SignInActivity.LastSignInDateTime) {
                ((Get-Date) - (Get-Date $_.SignInActivity.LastSignInDateTime)).Days
            } else { "Never" }
        }},
        @{N='LicenseCount';E={$_.AssignedLicenses.Count}} | 
        Sort-Object DaysInactive -Descending | Select-Object -First 20 | Format-Table -AutoSize
}

# Export if requested
if ($ExportToCsv) {
    $timestamp = Get-Date -Format "yyyy-MM-dd-HHmm"
    $outputPath = ".\M365-User-Analysis-$timestamp.csv"
    
    $exportData = foreach ($user in $users) {
        [PSCustomObject]@{
            UserPrincipalName = $user.UserPrincipalName
            DisplayName = $user.DisplayName
            AccountEnabled = $user.AccountEnabled
            UserType = $user.UserType
            CreatedDateTime = $user.CreatedDateTime
            LastSignIn = $user.SignInActivity.LastSignInDateTime
            DaysSinceLastSignIn = if ($user.SignInActivity.LastSignInDateTime) {
                ((Get-Date) - (Get-Date $user.SignInActivity.LastSignInDateTime)).Days
            } else { "Never" }
            LicenseCount = $user.AssignedLicenses.Count
            AssignedLicenseSkuIds = ($user.AssignedLicenses.SkuId -join "; ")
        }
    }
    
    $exportData | Export-Csv -Path $outputPath -NoTypeInformation
    Write-Host "`nData exported to: $outputPath" -ForegroundColor Green
}

Write-Host "`n--- RECOMMENDATIONS ---" -ForegroundColor Cyan
Write-Host "• Remove licenses from $($disabledUsers.Count) disabled users" -ForegroundColor White
Write-Host "• Review licenses for $($inactiveUsers.Count) inactive users" -ForegroundColor White
Write-Host "• Note: Full license details require Organization.Read.All permission" -ForegroundColor Yellow

Disconnect-MgGraph
Write-Host "`nDisconnected from Microsoft Graph" -ForegroundColor Green