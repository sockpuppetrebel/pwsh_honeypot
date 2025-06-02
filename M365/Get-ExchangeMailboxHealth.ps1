#Requires -Modules ExchangeOnlineManagement

<#
.SYNOPSIS
    Comprehensive Exchange Online mailbox health and usage analysis
.DESCRIPTION
    Analyzes mailbox sizes, permissions, forwarding rules, and compliance status.
    Identifies mailboxes that may need attention for storage, security, or governance.
.PARAMETER IncludeSharedMailboxes
    Include shared mailboxes in analysis
.PARAMETER IncludeResourceMailboxes
    Include room and equipment mailboxes
.PARAMETER CheckForwarding
    Check for mail forwarding rules and external forwarding
.PARAMETER ExportToExcel
    Export detailed results to Excel format
.PARAMETER SizeThresholdGB
    Mailbox size threshold in GB for flagging large mailboxes (default: 50)
.EXAMPLE
    .\Get-ExchangeMailboxHealth.ps1 -IncludeSharedMailboxes -CheckForwarding -ExportToExcel -SizeThresholdGB 25
#>

param(
    [Parameter(Mandatory = $false)]
    [switch]$IncludeSharedMailboxes,
    
    [Parameter(Mandatory = $false)]
    [switch]$IncludeResourceMailboxes,
    
    [Parameter(Mandatory = $false)]
    [switch]$CheckForwarding,
    
    [Parameter(Mandatory = $false)]
    [switch]$ExportToExcel,
    
    [Parameter(Mandatory = $false)]
    [int]$SizeThresholdGB = 50
)

# Connect to Exchange Online
try {
    Connect-ExchangeOnline -ShowProgress $true
    Write-Host "Connected to Exchange Online" -ForegroundColor Green
}
catch {
    Write-Error "Failed to connect to Exchange Online: $($_.Exception.Message)"
    exit 1
}

Write-Host "=== EXCHANGE ONLINE MAILBOX HEALTH ANALYSIS ===" -ForegroundColor Cyan

# Build mailbox filter
$mailboxTypes = @("UserMailbox")
if ($IncludeSharedMailboxes) { $mailboxTypes += "SharedMailbox" }
if ($IncludeResourceMailboxes) { $mailboxTypes += "RoomMailbox", "EquipmentMailbox" }

Write-Host "Retrieving mailboxes..." -ForegroundColor Yellow
$mailboxes = Get-Mailbox -ResultSize Unlimited | Where-Object RecipientTypeDetails -in $mailboxTypes

if ($mailboxes.Count -eq 0) {
    Write-Host "No mailboxes found matching criteria" -ForegroundColor Yellow
    Disconnect-ExchangeOnline -Confirm:$false
    exit 0
}

Write-Host "Analyzing $($mailboxes.Count) mailboxes..." -ForegroundColor Yellow

$mailboxHealthData = @()
$mailboxCounter = 0
$sizeThresholdBytes = $SizeThresholdGB * 1GB

foreach ($mailbox in $mailboxes) {
    $mailboxCounter++
    Write-Progress -Activity "Analyzing mailboxes" -Status "$($mailbox.DisplayName) ($mailboxCounter/$($mailboxes.Count))" -PercentComplete (($mailboxCounter / $mailboxes.Count) * 100)
    
    try {
        # Get mailbox statistics
        $mailboxStats = Get-MailboxStatistics -Identity $mailbox.Identity
        
        # Get mailbox permissions
        $permissions = Get-MailboxPermission -Identity $mailbox.Identity | Where-Object { 
            $_.User -ne "NT AUTHORITY\SELF" -and 
            $_.User -notlike "S-1-*" -and
            $_.AccessRights -contains "FullAccess"
        }
        
        # Parse mailbox size
        $mailboxSizeBytes = 0
        $mailboxSizeGB = 0
        if ($mailboxStats.TotalItemSize) {
            $sizeString = $mailboxStats.TotalItemSize.ToString()
            if ($sizeString -match '([0-9,]+)\s*bytes') {
                $mailboxSizeBytes = [long]($matches[1] -replace ',', '')
                $mailboxSizeGB = [math]::Round($mailboxSizeBytes / 1GB, 2)
            }
        }
        
        # Get mailbox quota information
        $quotaGB = 0
        if ($mailbox.ProhibitSendQuota -and $mailbox.ProhibitSendQuota -ne "Unlimited") {
            $quotaString = $mailbox.ProhibitSendQuota.ToString()
            if ($quotaString -match '([0-9.]+)\s*GB') {
                $quotaGB = [decimal]$matches[1]
            } elseif ($quotaString -match '([0-9,]+)\s*MB') {
                $quotaGB = [decimal]($matches[1] -replace ',', '') / 1024
            }
        }
        
        # Calculate quota usage percentage
        $quotaUsagePercent = if ($quotaGB -gt 0) {
            [math]::Round(($mailboxSizeGB / $quotaGB) * 100, 2)
        } else { 0 }
        
        # Check last logon
        $daysSinceLastLogon = if ($mailboxStats.LastLogonTime) {
            (Get-Date) - $mailboxStats.LastLogonTime | Select-Object -ExpandProperty Days
        } else {
            $null
        }
        
        # Analyze forwarding settings
        $forwardingInfo = @{
            HasForwarding = $false
            ForwardingAddress = $null
            ForwardingSmtpAddress = $null
            DeliverToMailboxAndForward = $false
            ExternalForwarding = $false
        }
        
        if ($CheckForwarding) {
            if ($mailbox.ForwardingAddress) {
                $forwardingInfo.HasForwarding = $true
                $forwardingInfo.ForwardingAddress = $mailbox.ForwardingAddress
                $forwardingInfo.DeliverToMailboxAndForward = $mailbox.DeliverToMailboxAndForward
            }
            
            if ($mailbox.ForwardingSmtpAddress) {
                $forwardingInfo.HasForwarding = $true
                $forwardingInfo.ForwardingSmtpAddress = $mailbox.ForwardingSmtpAddress
                $forwardingInfo.DeliverToMailboxAndForward = $mailbox.DeliverToMailboxAndForward
                
                # Check if forwarding is external
                $domain = ($mailbox.ForwardingSmtpAddress -split '@')[1]
                $acceptedDomains = Get-AcceptedDomain
                $forwardingInfo.ExternalForwarding = $domain -notin $acceptedDomains.DomainName
            }
            
            # Check for inbox rules with forwarding
            try {
                $inboxRules = Get-InboxRule -Mailbox $mailbox.Identity
                $forwardingRules = $inboxRules | Where-Object { 
                    $_.ForwardTo -or $_.ForwardAsAttachmentTo -or $_.RedirectTo 
                }
                
                if ($forwardingRules) {
                    $forwardingInfo.HasForwarding = $true
                    $forwardingInfo.InboxRulesWithForwarding = $forwardingRules.Count
                }
            }
            catch {
                # Inbox rules check failed
            }
        }
        
        # Identify potential issues
        $healthIssues = @()
        
        if ($mailboxSizeBytes -gt $sizeThresholdBytes) {
            $healthIssues += "Large mailbox (>$SizeThresholdGB GB)"
        }
        
        if ($quotaUsagePercent -gt 90) {
            $healthIssues += "Near quota limit ($quotaUsagePercent%)"
        }
        
        if ($daysSinceLastLogon -and $daysSinceLastLogon -gt 90) {
            $healthIssues += "Inactive ($daysSinceLastLogon days since last logon)"
        }
        
        if ($permissions.Count -gt 5) {
            $healthIssues += "Many full access permissions ($($permissions.Count))"
        }
        
        if ($forwardingInfo.ExternalForwarding) {
            $healthIssues += "External mail forwarding enabled"
        }
        
        if ($mailbox.LitigationHoldEnabled -eq $false -and $mailbox.RecipientTypeDetails -eq "UserMailbox") {
            $healthIssues += "Litigation hold not enabled"
        }
        
        # Calculate health score
        $healthScore = 100
        $healthScore -= ($healthIssues.Count * 12)
        $healthScore = [math]::Max($healthScore, 0)
        
        $mailboxHealth = [PSCustomObject]@{
            DisplayName = $mailbox.DisplayName
            UserPrincipalName = $mailbox.UserPrincipalName
            PrimarySmtpAddress = $mailbox.PrimarySmtpAddress
            RecipientTypeDetails = $mailbox.RecipientTypeDetails
            
            # Size and Storage
            MailboxSizeGB = $mailboxSizeGB
            MailboxSizeBytes = $mailboxSizeBytes
            ItemCount = $mailboxStats.ItemCount
            DeletedItemCount = $mailboxStats.DeletedItemCount
            QuotaGB = $quotaGB
            QuotaUsagePercent = $quotaUsagePercent
            
            # Activity
            LastLogonTime = $mailboxStats.LastLogonTime
            DaysSinceLastLogon = $daysSinceLastLogon
            IsInactive = ($daysSinceLastLogon -and $daysSinceLastLogon -gt 90)
            
            # Permissions
            FullAccessPermissions = $permissions.Count
            FullAccessUsers = ($permissions | ForEach-Object { $_.User.ToString() }) -join "; "
            
            # Forwarding
            HasForwarding = $forwardingInfo.HasForwarding
            ForwardingAddress = $forwardingInfo.ForwardingAddress
            ForwardingSmtpAddress = $forwardingInfo.ForwardingSmtpAddress
            ExternalForwarding = $forwardingInfo.ExternalForwarding
            DeliverToMailboxAndForward = $forwardingInfo.DeliverToMailboxAndForward
            InboxRulesWithForwarding = $forwardingInfo.InboxRulesWithForwarding
            
            # Compliance and Security
            LitigationHoldEnabled = $mailbox.LitigationHoldEnabled
            LitigationHoldDuration = $mailbox.LitigationHoldDuration
            InPlaceHolds = ($mailbox.InPlaceHolds -join "; ")
            RetentionPolicy = $mailbox.RetentionPolicy
            
            # Archive
            ArchiveStatus = $mailbox.ArchiveStatus
            ArchiveDatabase = $mailbox.ArchiveDatabase
            
            # Additional Settings
            HiddenFromAddressListsEnabled = $mailbox.HiddenFromAddressListsEnabled
            RequireSenderAuthenticationEnabled = $mailbox.RequireSenderAuthenticationEnabled
            
            # Health Assessment
            HealthScore = $healthScore
            HealthIssues = ($healthIssues -join "; ")
            NeedsAttention = ($healthIssues.Count -gt 0)
            
            # Categories
            IsLargeMailbox = ($mailboxSizeBytes -gt $sizeThresholdBytes)
            IsNearQuota = ($quotaUsagePercent -gt 80)
            HasExcessivePermissions = ($permissions.Count -gt 5)
        }
        
        $mailboxHealthData += $mailboxHealth
    }
    catch {
        Write-Host "Error analyzing mailbox $($mailbox.DisplayName): $($_.Exception.Message)" -ForegroundColor Yellow
        
        # Add basic info even if detailed analysis fails
        $basicHealth = [PSCustomObject]@{
            DisplayName = $mailbox.DisplayName
            UserPrincipalName = $mailbox.UserPrincipalName
            PrimarySmtpAddress = $mailbox.PrimarySmtpAddress
            RecipientTypeDetails = $mailbox.RecipientTypeDetails
            HealthScore = 0
            HealthIssues = "Analysis failed: $($_.Exception.Message)"
            NeedsAttention = $true
        }
        
        $mailboxHealthData += $basicHealth
    }
}

Write-Progress -Activity "Analyzing mailboxes" -Completed

# Generate summary report
Write-Host "`n--- MAILBOX HEALTH SUMMARY ---" -ForegroundColor Cyan

$totalMailboxes = $mailboxHealthData.Count
$mailboxesNeedingAttention = ($mailboxHealthData | Where-Object NeedsAttention -eq $true).Count
$largeMailboxes = ($mailboxHealthData | Where-Object IsLargeMailbox -eq $true).Count
$inactiveMailboxes = ($mailboxHealthData | Where-Object IsInactive -eq $true).Count
$externalForwardingMailboxes = ($mailboxHealthData | Where-Object ExternalForwarding -eq $true).Count

Write-Host "Total Mailboxes: $totalMailboxes" -ForegroundColor White
Write-Host "Mailboxes Needing Attention: $mailboxesNeedingAttention" -ForegroundColor $(if ($mailboxesNeedingAttention -gt 0) { "Yellow" } else { "Green" })
Write-Host "Large Mailboxes (>$SizeThresholdGB GB): $largeMailboxes" -ForegroundColor White
Write-Host "Inactive Mailboxes (>90 days): $inactiveMailboxes" -ForegroundColor White
Write-Host "External Forwarding Enabled: $externalForwardingMailboxes" -ForegroundColor White

# Storage analysis
$totalStorageGB = ($mailboxHealthData | Where-Object MailboxSizeGB | Measure-Object MailboxSizeGB -Sum).Sum
$averageMailboxSizeGB = ($mailboxHealthData | Where-Object MailboxSizeGB | Measure-Object MailboxSizeGB -Average).Average

Write-Host "`nStorage Overview:" -ForegroundColor Yellow
Write-Host "Total Storage Used: $([math]::Round($totalStorageGB, 2)) GB" -ForegroundColor White
Write-Host "Average Mailbox Size: $([math]::Round($averageMailboxSizeGB, 2)) GB" -ForegroundColor White

# Top issues
if ($mailboxesNeedingAttention -gt 0) {
    Write-Host "`n--- MAILBOXES NEEDING ATTENTION ---" -ForegroundColor Red
    $mailboxHealthData | Where-Object NeedsAttention -eq $true | 
        Sort-Object HealthScore | 
        Select-Object -First 10 | 
        Select-Object DisplayName, RecipientTypeDetails, MailboxSizeGB, HealthScore, HealthIssues | 
        Format-Table -Wrap
}

# Large mailboxes
if ($largeMailboxes -gt 0) {
    Write-Host "`n--- LARGEST MAILBOXES ---" -ForegroundColor Yellow
    $mailboxHealthData | Where-Object IsLargeMailbox -eq $true | 
        Sort-Object MailboxSizeGB -Descending | 
        Select-Object -First 10 | 
        Select-Object DisplayName, RecipientTypeDetails, MailboxSizeGB, QuotaUsagePercent | 
        Format-Table
}

# Inactive mailboxes
if ($inactiveMailboxes -gt 0) {
    Write-Host "`n--- INACTIVE MAILBOXES ---" -ForegroundColor Yellow
    $mailboxHealthData | Where-Object IsInactive -eq $true | 
        Sort-Object DaysSinceLastLogon -Descending | 
        Select-Object -First 10 | 
        Select-Object DisplayName, RecipientTypeDetails, DaysSinceLastLogon, MailboxSizeGB | 
        Format-Table
}

# External forwarding analysis
if ($externalForwardingMailboxes -gt 0) {
    Write-Host "`n--- EXTERNAL FORWARDING DETECTED ---" -ForegroundColor Red
    $mailboxHealthData | Where-Object ExternalForwarding -eq $true | 
        Select-Object DisplayName, ForwardingSmtpAddress, DeliverToMailboxAndForward | 
        Format-Table -Wrap
}

# Permission analysis
$excessivePermissionsMailboxes = ($mailboxHealthData | Where-Object HasExcessivePermissions -eq $true).Count
if ($excessivePermissionsMailboxes -gt 0) {
    Write-Host "`n--- MAILBOXES WITH EXCESSIVE PERMISSIONS ---" -ForegroundColor Yellow
    $mailboxHealthData | Where-Object HasExcessivePermissions -eq $true | 
        Sort-Object FullAccessPermissions -Descending | 
        Select-Object -First 10 | 
        Select-Object DisplayName, RecipientTypeDetails, FullAccessPermissions | 
        Format-Table
}

# Compliance analysis
$nonComplianceMailboxes = ($mailboxHealthData | Where-Object { $_.LitigationHoldEnabled -eq $false -and $_.RecipientTypeDetails -eq "UserMailbox" }).Count
if ($nonComplianceMailboxes -gt 0) {
    Write-Host "`n--- COMPLIANCE ISSUES ---" -ForegroundColor Yellow
    Write-Host "User mailboxes without litigation hold: $nonComplianceMailboxes" -ForegroundColor White
}

# Recommendations
Write-Host "`n--- RECOMMENDATIONS ---" -ForegroundColor Cyan

$recommendations = @()

if ($mailboxesNeedingAttention -gt 0) {
    $recommendations += "Review $mailboxesNeedingAttention mailboxes with health issues"
}

if ($largeMailboxes -gt 0) {
    $recommendations += "Implement archive policies for $largeMailboxes large mailboxes"
}

if ($inactiveMailboxes -gt 0) {
    $recommendations += "Review $inactiveMailboxes inactive mailboxes for conversion to shared or removal"
}

if ($externalForwardingMailboxes -gt 0) {
    $recommendations += "Audit $externalForwardingMailboxes mailboxes with external forwarding for security risks"
}

if ($excessivePermissionsMailboxes -gt 0) {
    $recommendations += "Review and reduce permissions for $excessivePermissionsMailboxes mailboxes with excessive access"
}

if ($nonComplianceMailboxes -gt 0) {
    $recommendations += "Enable litigation hold for $nonComplianceMailboxes user mailboxes to meet compliance requirements"
}

if ($recommendations.Count -gt 0) {
    $recommendations | ForEach-Object { Write-Host "• $_" -ForegroundColor Cyan }
} else {
    Write-Host "• All mailboxes appear to be in good health!" -ForegroundColor Green
}

# Export results
if ($ExportToExcel) {
    $timestamp = Get-Date -Format "yyyy-MM-dd-HHmm"
    $outputPath = ".\Exchange-Mailbox-Health-$timestamp.xlsx"
    
    if (Get-Module -ListAvailable -Name ImportExcel) {
        # Main health data
        $mailboxHealthData | Export-Excel -Path $outputPath -WorksheetName "Mailbox Health" -AutoSize -FreezeTopRow
        
        # Summary statistics
        $summaryData = @(
            [PSCustomObject]@{ Metric = "Total Mailboxes"; Value = $totalMailboxes }
            [PSCustomObject]@{ Metric = "Needing Attention"; Value = $mailboxesNeedingAttention }
            [PSCustomObject]@{ Metric = "Large Mailboxes"; Value = $largeMailboxes }
            [PSCustomObject]@{ Metric = "Inactive Mailboxes"; Value = $inactiveMailboxes }
            [PSCustomObject]@{ Metric = "External Forwarding"; Value = $externalForwardingMailboxes }
            [PSCustomObject]@{ Metric = "Total Storage (GB)"; Value = [math]::Round($totalStorageGB, 2) }
            [PSCustomObject]@{ Metric = "Average Size (GB)"; Value = [math]::Round($averageMailboxSizeGB, 2) }
        )
        
        $summaryData | Export-Excel -Path $outputPath -WorksheetName "Summary" -AutoSize
        
        # Issues breakdown
        if ($mailboxesNeedingAttention -gt 0) {
            $mailboxHealthData | Where-Object NeedsAttention -eq $true | 
                Export-Excel -Path $outputPath -WorksheetName "Needs Attention" -AutoSize
        }
        
        Write-Host "`nDetailed report exported to: $outputPath" -ForegroundColor Green
    } else {
        Write-Host "ImportExcel module not available. Exporting to CSV..." -ForegroundColor Yellow
        $mailboxHealthData | Export-Csv -Path ".\Exchange-Mailbox-Health-$timestamp.csv" -NoTypeInformation
        Write-Host "Report exported to: .\Exchange-Mailbox-Health-$timestamp.csv" -ForegroundColor Green
    }
}

Write-Host "`n--- NEXT STEPS ---" -ForegroundColor Cyan
Write-Host "• Implement mailbox lifecycle management policies" -ForegroundColor White
Write-Host "• Set up automated archiving for large mailboxes" -ForegroundColor White
Write-Host "• Review and audit external forwarding configurations" -ForegroundColor White
Write-Host "• Enable litigation hold for compliance requirements" -ForegroundColor White

Disconnect-ExchangeOnline -Confirm:$false
Write-Host "Mailbox health analysis complete" -ForegroundColor Green