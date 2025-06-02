#Requires -Modules Microsoft.Graph.Users, Microsoft.Graph.Groups, ExchangeOnlineManagement

<#
.SYNOPSIS
    Comprehensive M365 user troubleshooting and repair tool
.DESCRIPTION
    Diagnoses and attempts to resolve common M365 user issues across multiple services:
    - License assignment problems
    - Group membership issues  
    - Exchange Online mailbox problems
    - OneDrive sync issues
    - Teams access problems
.PARAMETER UserPrincipalName
    UPN of the user to troubleshoot
.PARAMETER FixIssues
    Switch to automatically attempt repairs (use with caution)
.EXAMPLE
    .\Repair-M365UserIssues.ps1 -UserPrincipalName "john.doe@company.com" -FixIssues
#>

param(
    [Parameter(Mandatory = $true)]
    [string]$UserPrincipalName,
    
    [Parameter(Mandatory = $false)]
    [switch]$FixIssues
)

# Initialize results tracking
$issuesFound = @()
$repairActions = @()

function Write-Issue {
    param([string]$Category, [string]$Description, [string]$Severity = "Warning")
    
    $issue = [PSCustomObject]@{
        Category = $Category
        Description = $Description
        Severity = $Severity
        Timestamp = Get-Date
    }
    
    $script:issuesFound += $issue
    
    $color = switch ($Severity) {
        "Critical" { "Red" }
        "Warning" { "Yellow" }
        "Info" { "Cyan" }
        default { "White" }
    }
    
    Write-Host "[$Severity] $Category: $Description" -ForegroundColor $color
}

function Write-Repair {
    param([string]$Action, [string]$Result)
    
    $repair = [PSCustomObject]@{
        Action = $Action
        Result = $Result
        Timestamp = Get-Date
    }
    
    $script:repairActions += $repair
    Write-Host "[REPAIR] $Action -> $Result" -ForegroundColor Green
}

# Connect to required services
try {
    Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Yellow
    Connect-MgGraph -Scopes "User.Read.All", "Group.Read.All", "Directory.Read.All"
    
    Write-Host "Connecting to Exchange Online..." -ForegroundColor Yellow
    Connect-ExchangeOnline -ShowProgress:$false
}
catch {
    Write-Error "Failed to connect to required services: $($_.Exception.Message)"
    exit 1
}

Write-Host "`n=== STARTING M365 USER DIAGNOSTICS FOR $UserPrincipalName ===" -ForegroundColor Cyan

# 1. Basic User Information
try {
    $user = Get-MgUser -UserId $UserPrincipalName -Property "Id,UserPrincipalName,DisplayName,AccountEnabled,AssignedLicenses,UserType,CreatedDateTime,OnPremisesSyncEnabled"
    Write-Host "`nUser found: $($user.DisplayName)" -ForegroundColor Green
    
    if (-not $user.AccountEnabled) {
        Write-Issue -Category "Account" -Description "User account is disabled" -Severity "Critical"
    }
    
    if ($user.OnPremisesSyncEnabled) {
        Write-Issue -Category "Sync" -Description "User is synced from on-premises (changes must be made in AD)" -Severity "Info"
    }
}
catch {
    Write-Issue -Category "User" -Description "User not found or inaccessible" -Severity "Critical"
    exit 1
}

# 2. License Analysis
Write-Host "`n--- Analyzing Licenses ---" -ForegroundColor Yellow

if ($user.AssignedLicenses.Count -eq 0) {
    Write-Issue -Category "Licensing" -Description "No licenses assigned to user" -Severity "Critical"
} else {
    foreach ($license in $user.AssignedLicenses) {
        $skuDetails = Get-MgSubscribedSku | Where-Object SkuId -eq $license.SkuId
        Write-Host "Licensed for: $($skuDetails.SkuPartNumber)" -ForegroundColor Green
        
        # Check for disabled service plans
        if ($license.DisabledPlans.Count -gt 0) {
            Write-Issue -Category "Licensing" -Description "Some service plans are disabled for $($skuDetails.SkuPartNumber)" -Severity "Warning"
        }
    }
}

# 3. Group Membership Check
Write-Host "`n--- Checking Group Memberships ---" -ForegroundColor Yellow

$userGroups = Get-MgUserMemberOf -UserId $user.Id
$securityGroups = $userGroups | Where-Object { $_.'@odata.type' -eq '#microsoft.graph.group' }
$directoryRoles = $userGroups | Where-Object { $_.'@odata.type' -eq '#microsoft.graph.directoryRole' }

Write-Host "Member of $($securityGroups.Count) groups and $($directoryRoles.Count) directory roles" -ForegroundColor Green

# Check for common required groups
$commonGroups = @("All Users", "All Company", "Everyone")
foreach ($groupName in $commonGroups) {
    $group = Get-MgGroup -Filter "displayName eq '$groupName'" -ErrorAction SilentlyContinue
    if ($group) {
        $membership = Get-MgGroupMember -GroupId $group.Id | Where-Object Id -eq $user.Id
        if (-not $membership) {
            Write-Issue -Category "Groups" -Description "User not in common group: $groupName" -Severity "Warning"
        }
    }
}

# 4. Exchange Online Mailbox Check
Write-Host "`n--- Checking Exchange Online ---" -ForegroundColor Yellow

try {
    $mailbox = Get-Mailbox -Identity $UserPrincipalName -ErrorAction Stop
    Write-Host "Mailbox found: $($mailbox.PrimarySmtpAddress)" -ForegroundColor Green
    
    # Check mailbox statistics
    $mailboxStats = Get-MailboxStatistics -Identity $UserPrincipalName
    $quotaPercent = [math]::Round(($mailboxStats.TotalItemSize.Value.ToBytes() / $mailbox.ProhibitSendQuota.Value.ToBytes()) * 100, 2)
    
    if ($quotaPercent -gt 90) {
        Write-Issue -Category "Exchange" -Description "Mailbox is $quotaPercent% full" -Severity "Critical"
    } elseif ($quotaPercent -gt 75) {
        Write-Issue -Category "Exchange" -Description "Mailbox is $quotaPercent% full" -Severity "Warning"
    }
    
    # Check for forwarding rules
    $forwardingRules = Get-InboxRule -Mailbox $UserPrincipalName | Where-Object { $_.ForwardTo -or $_.RedirectTo }
    if ($forwardingRules) {
        Write-Issue -Category "Exchange" -Description "Found $($forwardingRules.Count) forwarding rules" -Severity "Warning"
    }
    
    # Check for recent logons
    $recentLogons = Get-MailboxStatistics -Identity $UserPrincipalName | Select-Object LastLogonTime
    if ($recentLogons.LastLogonTime -lt (Get-Date).AddDays(-30)) {
        Write-Issue -Category "Exchange" -Description "No recent mailbox activity (last logon: $($recentLogons.LastLogonTime))" -Severity "Warning"
    }
}
catch {
    Write-Issue -Category "Exchange" -Description "Mailbox not found or inaccessible: $($_.Exception.Message)" -Severity "Critical"
}

# 5. OneDrive Check  
Write-Host "`n--- Checking OneDrive ---" -ForegroundColor Yellow

try {
    $oneDriveSite = Get-MgUserDrive -UserId $user.Id -ErrorAction Stop
    Write-Host "OneDrive found: $($oneDriveSite.WebUrl)" -ForegroundColor Green
    
    # Check quota usage
    if ($oneDriveSite.Quota.Used -and $oneDriveSite.Quota.Total) {
        $driveQuotaPercent = [math]::Round(($oneDriveSite.Quota.Used / $oneDriveSite.Quota.Total) * 100, 2)
        if ($driveQuotaPercent -gt 90) {
            Write-Issue -Category "OneDrive" -Description "OneDrive is $driveQuotaPercent% full" -Severity "Critical"
        }
    }
}
catch {
    Write-Issue -Category "OneDrive" -Description "OneDrive not accessible or not provisioned" -Severity "Warning"
}

# 6. Teams Check
Write-Host "`n--- Checking Teams ---" -ForegroundColor Yellow

$teamsLicense = $user.AssignedLicenses | ForEach-Object {
    $sku = Get-MgSubscribedSku | Where-Object SkuId -eq $_.SkuId
    $sku.ServicePlans | Where-Object ServicePlanName -like "*TEAMS*" -and $_.ProvisioningStatus -eq "Success"
}

if ($teamsLicense) {
    Write-Host "Teams license found and enabled" -ForegroundColor Green
} else {
    Write-Issue -Category "Teams" -Description "Teams license not found or disabled" -Severity "Warning"
}

# 7. Conditional Access Policy Impact
Write-Host "`n--- Checking Conditional Access ---" -ForegroundColor Yellow

try {
    $signInLogs = Get-MgAuditLogSignIn -Filter "userId eq '$($user.Id)'" -Top 10
    $recentFailures = $signInLogs | Where-Object { $_.Status.ErrorCode -ne 0 }
    
    if ($recentFailures) {
        Write-Issue -Category "Authentication" -Description "Found $($recentFailures.Count) recent sign-in failures" -Severity "Warning"
    }
}
catch {
    Write-Issue -Category "Authentication" -Description "Unable to retrieve sign-in logs" -Severity "Info"
}

# Repair Actions (if requested)
if ($FixIssues) {
    Write-Host "`n=== ATTEMPTING REPAIRS ===" -ForegroundColor Cyan
    
    # Example repairs (customize based on your environment)
    foreach ($issue in $issuesFound) {
        switch ($issue.Category) {
            "Licensing" {
                if ($issue.Description -like "*No licenses assigned*") {
                    Write-Host "Would assign default license here (disabled for safety)" -ForegroundColor Yellow
                    Write-Repair -Action "License Assignment" -Result "Skipped (manual intervention required)"
                }
            }
            "Exchange" {
                if ($issue.Description -like "*forwarding rules*") {
                    Write-Host "Consider reviewing forwarding rules manually" -ForegroundColor Yellow
                    Write-Repair -Action "Forwarding Rules Review" -Result "Manual review recommended"
                }
            }
        }
    }
}

# Summary Report
Write-Host "`n=== SUMMARY REPORT ===" -ForegroundColor Cyan
Write-Host "Total Issues Found: $($issuesFound.Count)" -ForegroundColor White

$criticalIssues = $issuesFound | Where-Object Severity -eq "Critical"
$warningIssues = $issuesFound | Where-Object Severity -eq "Warning"

if ($criticalIssues) {
    Write-Host "Critical Issues: $($criticalIssues.Count)" -ForegroundColor Red
    $criticalIssues | ForEach-Object { Write-Host "  - $($_.Category): $($_.Description)" -ForegroundColor Red }
}

if ($warningIssues) {
    Write-Host "Warnings: $($warningIssues.Count)" -ForegroundColor Yellow
    $warningIssues | ForEach-Object { Write-Host "  - $($_.Category): $($_.Description)" -ForegroundColor Yellow }
}

if ($issuesFound.Count -eq 0) {
    Write-Host "No issues detected - user configuration appears healthy!" -ForegroundColor Green
}

# Export detailed report
$reportPath = ".\M365-UserReport-$(($UserPrincipalName -split '@')[0])-$(Get-Date -Format 'yyyy-MM-dd-HHmm').json"
$report = @{
    User = $user
    Issues = $issuesFound
    RepairActions = $repairActions
    GeneratedDate = Get-Date
}

$report | ConvertTo-Json -Depth 5 | Out-File -FilePath $reportPath
Write-Host "`nDetailed report saved to: $reportPath" -ForegroundColor Green

# Cleanup
Disconnect-MgGraph
Disconnect-ExchangeOnline -Confirm:$false