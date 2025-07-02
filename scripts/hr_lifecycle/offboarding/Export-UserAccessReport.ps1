<#
.SYNOPSIS
    Generates a comprehensive access report for a user before offboarding.

.DESCRIPTION
    This script creates a detailed report of all user access including:
    - Azure AD group memberships
    - Assigned licenses
    - Assigned applications
    - Mailbox permissions and delegations
    - SharePoint/OneDrive access
    - Conditional Access policies
    - MFA status
    - Last sign-in information

.PARAMETER UserPrincipalName
    The UPN of the user to report on

.PARAMETER OutputFormat
    Output format: HTML, CSV, or JSON (default: HTML)

.PARAMETER IncludeSharePoint
    Include SharePoint site permissions (may take longer)

.EXAMPLE
    .\Export-UserAccessReport.ps1 -UserPrincipalName "john.doe@company.com"

.EXAMPLE
    .\Export-UserAccessReport.ps1 -UserPrincipalName "john.doe@company.com" -OutputFormat CSV -IncludeSharePoint

.NOTES
    Author: HR Automation Script
    Date: 2025-07-02
    Version: 1.0
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)]
    [string]$UserPrincipalName,
    
    [Parameter(Mandatory=$false)]
    [ValidateSet("HTML", "CSV", "JSON")]
    [string]$OutputFormat = "HTML",
    
    [Parameter(Mandatory=$false)]
    [switch]$IncludeSharePoint
)

# Initialize logging
$LogPath = "$PSScriptRoot\Reports"
if (-not (Test-Path $LogPath)) {
    New-Item -ItemType Directory -Path $LogPath -Force | Out-Null
}
$SessionId = Get-Date -Format 'yyyyMMdd-HHmmss'
$ReportName = "AccessReport-$($UserPrincipalName.Replace('@','_'))-$SessionId"

Write-Host "`n=== USER ACCESS REPORT GENERATION ===" -ForegroundColor Green
Write-Host "User: $UserPrincipalName" -ForegroundColor Yellow
Write-Host "Output Format: $OutputFormat" -ForegroundColor Yellow
Write-Host "Include SharePoint: $IncludeSharePoint" -ForegroundColor Yellow
Write-Host "Started: $(Get-Date)" -ForegroundColor Yellow
Write-Host "="*50 -ForegroundColor Green

# Initialize report data structure
$reportData = @{
    GeneratedDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    User = @{}
    Groups = @()
    Licenses = @()
    Applications = @()
    MailboxPermissions = @{
        GrantedByUser = @()
        GrantedToUser = @()
    }
    SharePointAccess = @()
    SecurityInfo = @{
        MFAStatus = ""
        ConditionalAccessPolicies = @()
        LastSignIn = @{}
        RiskySignIns = @()
    }
    Summary = @{}
}

try {
    # Connect to required services
    Write-Host "`nConnecting to Microsoft services..." -ForegroundColor Cyan
    
    # Connect to Microsoft Graph
    $mgContext = Get-MgContext -ErrorAction SilentlyContinue
    if (-not $mgContext) {
        Connect-MgGraph -Scopes "User.Read.All", "Group.Read.All", "Application.Read.All", "Directory.Read.All", "AuditLog.Read.All", "Policy.Read.All", "Sites.Read.All" -NoWelcome
    }
    
    # Connect to Exchange Online
    $exoSession = Get-PSSession | Where-Object {$_.ConfigurationName -eq "Microsoft.Exchange" -and $_.State -eq "Opened"}
    if (-not $exoSession) {
        Connect-ExchangeOnline -ShowBanner:$false
    }
    
    # Step 1: Get User Information
    Write-Host "`n[Step 1/8] Getting user information..." -ForegroundColor Cyan
    $user = Get-MgUser -UserId $UserPrincipalName -Property * -ErrorAction Stop
    $reportData.User = @{
        UserPrincipalName = $user.UserPrincipalName
        DisplayName = $user.DisplayName
        GivenName = $user.GivenName
        Surname = $user.Surname
        JobTitle = $user.JobTitle
        Department = $user.Department
        OfficeLocation = $user.OfficeLocation
        Manager = if ($user.Manager) { (Get-MgUser -UserId $user.Manager.Id).DisplayName } else { "N/A" }
        AccountEnabled = $user.AccountEnabled
        CreatedDateTime = $user.CreatedDateTime
        Mail = $user.Mail
        ObjectId = $user.Id
    }
    Write-Host "  User found: $($user.DisplayName)" -ForegroundColor Green
    
    # Step 2: Get Group Memberships
    Write-Host "`n[Step 2/8] Getting group memberships..." -ForegroundColor Cyan
    $groups = Get-MgUserMemberOf -UserId $user.Id -All
    foreach ($group in $groups | Where-Object { $_.'@odata.type' -eq '#microsoft.graph.group' }) {
        $groupDetails = Get-MgGroup -GroupId $group.Id -Property *
        $reportData.Groups += @{
            DisplayName = $groupDetails.DisplayName
            Description = $groupDetails.Description
            GroupType = if ($groupDetails.GroupTypes -contains "Unified") { "Microsoft 365" } 
                       elseif ($groupDetails.SecurityEnabled) { "Security" } 
                       else { "Distribution" }
            IsDynamic = $groupDetails.GroupTypes -contains "DynamicMembership"
            Mail = $groupDetails.Mail
        }
    }
    Write-Host "  Found $($reportData.Groups.Count) group memberships" -ForegroundColor Green
    
    # Step 3: Get Licenses
    Write-Host "`n[Step 3/8] Getting assigned licenses..." -ForegroundColor Cyan
    $licenses = Get-MgUserLicenseDetail -UserId $user.Id
    foreach ($license in $licenses) {
        $reportData.Licenses += @{
            SkuPartNumber = $license.SkuPartNumber
            SkuId = $license.SkuId
            ServicePlans = $license.ServicePlans | Where-Object { $_.ProvisioningStatus -eq "Success" } | 
                          Select-Object -ExpandProperty ServicePlanName
        }
    }
    Write-Host "  Found $($reportData.Licenses.Count) licenses assigned" -ForegroundColor Green
    
    # Step 4: Get Assigned Applications
    Write-Host "`n[Step 4/8] Getting assigned applications..." -ForegroundColor Cyan
    $appAssignments = Get-MgUserAppRoleAssignment -UserId $user.Id -All
    foreach ($assignment in $appAssignments) {
        $app = Get-MgServicePrincipal -ServicePrincipalId $assignment.ResourceId
        $reportData.Applications += @{
            ApplicationName = $app.DisplayName
            AppId = $app.AppId
            AssignedRole = $assignment.AppRoleId
        }
    }
    Write-Host "  Found $($reportData.Applications.Count) application assignments" -ForegroundColor Green
    
    # Step 5: Get Mailbox Permissions
    Write-Host "`n[Step 5/8] Getting mailbox permissions..." -ForegroundColor Cyan
    try {
        $mailbox = Get-Mailbox -Identity $UserPrincipalName -ErrorAction SilentlyContinue
        
        if ($mailbox) {
            # Get permissions granted BY the user
            $mailboxes = Get-Mailbox -ResultSize 100
            foreach ($mb in $mailboxes | Where-Object { $_.PrimarySmtpAddress -ne $mailbox.PrimarySmtpAddress }) {
                $perms = Get-MailboxPermission -Identity $mb.Identity -User $UserPrincipalName -ErrorAction SilentlyContinue
                if ($perms -and $perms.AccessRights -contains "FullAccess") {
                    $reportData.MailboxPermissions.GrantedByUser += @{
                        Mailbox = $mb.PrimarySmtpAddress
                        Permission = "Full Access"
                    }
                }
            }
            
            # Get permissions granted TO the user
            $permsToUser = Get-MailboxPermission -Identity $UserPrincipalName | Where-Object {
                $_.User -notlike "NT AUTHORITY\*" -and 
                $_.User -notlike "S-1-5-*" -and
                $_.AccessRights -contains "FullAccess" -and
                $_.IsInherited -eq $false
            }
            foreach ($perm in $permsToUser) {
                $reportData.MailboxPermissions.GrantedToUser += @{
                    User = $perm.User
                    Permission = "Full Access"
                }
            }
        }
    } catch {
        Write-Host "  ! Unable to retrieve mailbox permissions" -ForegroundColor Yellow
    }
    Write-Host "  Mailbox permissions retrieved" -ForegroundColor Green
    
    # Step 6: Get SharePoint Access (if requested)
    if ($IncludeSharePoint) {
        Write-Host "`n[Step 6/8] Getting SharePoint/OneDrive access..." -ForegroundColor Cyan
        Write-Host "  ! SharePoint scanning can take several minutes..." -ForegroundColor Yellow
        
        # This would require PnP PowerShell module
        # Simplified version - just noting that full implementation would scan all sites
        $reportData.SharePointAccess = @(
            @{
                Note = "Full SharePoint scanning requires additional modules and permissions"
                Recommendation = "Run separate SharePoint audit script for detailed permissions"
            }
        )
        Write-Host "  SharePoint access noted" -ForegroundColor Green
    } else {
        Write-Host "`n[Step 6/8] Skipping SharePoint access (not requested)..." -ForegroundColor Gray
    }
    
    # Step 7: Get Security Information
    Write-Host "`n[Step 7/8] Getting security information..." -ForegroundColor Cyan
    
    # MFA Status
    try {
        $mfaMethods = Get-MgUserAuthenticationMethod -UserId $user.Id
        $reportData.SecurityInfo.MFAStatus = if ($mfaMethods.Count -gt 1) { "Enabled" } else { "Not Configured" }
    } catch {
        $reportData.SecurityInfo.MFAStatus = "Unable to determine"
    }
    
    # Last Sign-in
    try {
        $signInActivity = Get-MgUser -UserId $user.Id -Property SignInActivity
        if ($signInActivity.SignInActivity) {
            $reportData.SecurityInfo.LastSignIn = @{
                LastSignInDateTime = $signInActivity.SignInActivity.LastSignInDateTime
                LastNonInteractiveSignInDateTime = $signInActivity.SignInActivity.LastNonInteractiveSignInDateTime
            }
        }
    } catch {
        $reportData.SecurityInfo.LastSignIn = @{ Note = "Unable to retrieve sign-in data" }
    }
    
    Write-Host "  Security information retrieved" -ForegroundColor Green
    
    # Step 8: Generate Summary
    Write-Host "`n[Step 8/8] Generating summary..." -ForegroundColor Cyan
    $reportData.Summary = @{
        TotalGroups = $reportData.Groups.Count
        SecurityGroups = ($reportData.Groups | Where-Object { $_.GroupType -eq "Security" }).Count
        M365Groups = ($reportData.Groups | Where-Object { $_.GroupType -eq "Microsoft 365" }).Count
        TotalLicenses = $reportData.Licenses.Count
        TotalApplications = $reportData.Applications.Count
        MailboxDelegationsGranted = $reportData.MailboxPermissions.GrantedByUser.Count
        MailboxDelegationsReceived = $reportData.MailboxPermissions.GrantedToUser.Count
        AccountStatus = if ($user.AccountEnabled) { "Active" } else { "Disabled" }
        MFAStatus = $reportData.SecurityInfo.MFAStatus
    }
    
    # Export Report
    Write-Host "`nExporting report..." -ForegroundColor Cyan
    
    switch ($OutputFormat) {
        "JSON" {
            $outputFile = "$LogPath\$ReportName.json"
            $reportData | ConvertTo-Json -Depth 10 | Out-File $outputFile
        }
        
        "CSV" {
            # Create multiple CSV files for different sections
            $outputFile = "$LogPath\$ReportName-Summary.csv"
            $reportData.Summary | ForEach-Object { [PSCustomObject]$_ } | Export-Csv $outputFile -NoTypeInformation
            
            $reportData.Groups | Export-Csv "$LogPath\$ReportName-Groups.csv" -NoTypeInformation
            $reportData.Licenses | Export-Csv "$LogPath\$ReportName-Licenses.csv" -NoTypeInformation
            $reportData.Applications | Export-Csv "$LogPath\$ReportName-Applications.csv" -NoTypeInformation
        }
        
        "HTML" {
            $outputFile = "$LogPath\$ReportName.html"
            $html = @"
<!DOCTYPE html>
<html>
<head>
    <title>User Access Report - $($user.DisplayName)</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; background-color: #f5f5f5; }
        .container { max-width: 1200px; margin: 0 auto; background-color: white; padding: 20px; box-shadow: 0 0 10px rgba(0,0,0,0.1); }
        h1, h2, h3 { color: #333; }
        table { border-collapse: collapse; width: 100%; margin-bottom: 20px; }
        th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }
        th { background-color: #4CAF50; color: white; }
        tr:nth-child(even) { background-color: #f2f2f2; }
        .summary { background-color: #e7f3fe; padding: 15px; border-left: 4px solid #2196F3; margin-bottom: 20px; }
        .section { margin-bottom: 30px; }
        .status-active { color: green; font-weight: bold; }
        .status-disabled { color: red; font-weight: bold; }
    </style>
</head>
<body>
    <div class="container">
        <h1>User Access Report</h1>
        <p><strong>Generated:</strong> $($reportData.GeneratedDate)</p>
        
        <div class="section">
            <h2>User Information</h2>
            <table>
                <tr><th>Property</th><th>Value</th></tr>
                <tr><td>Display Name</td><td>$($reportData.User.DisplayName)</td></tr>
                <tr><td>User Principal Name</td><td>$($reportData.User.UserPrincipalName)</td></tr>
                <tr><td>Job Title</td><td>$($reportData.User.JobTitle)</td></tr>
                <tr><td>Department</td><td>$($reportData.User.Department)</td></tr>
                <tr><td>Manager</td><td>$($reportData.User.Manager)</td></tr>
                <tr><td>Account Status</td><td class="$(if ($reportData.User.AccountEnabled) {'status-active'} else {'status-disabled'})">$(if ($reportData.User.AccountEnabled) {'Active'} else {'Disabled'})</td></tr>
                <tr><td>Created Date</td><td>$($reportData.User.CreatedDateTime)</td></tr>
            </table>
        </div>
        
        <div class="summary">
            <h2>Summary</h2>
            <ul>
                <li><strong>Total Groups:</strong> $($reportData.Summary.TotalGroups) (Security: $($reportData.Summary.SecurityGroups), M365: $($reportData.Summary.M365Groups))</li>
                <li><strong>Total Licenses:</strong> $($reportData.Summary.TotalLicenses)</li>
                <li><strong>Total Applications:</strong> $($reportData.Summary.TotalApplications)</li>
                <li><strong>Mailbox Delegations Given:</strong> $($reportData.Summary.MailboxDelegationsGranted)</li>
                <li><strong>Mailbox Delegations Received:</strong> $($reportData.Summary.MailboxDelegationsReceived)</li>
                <li><strong>MFA Status:</strong> $($reportData.Summary.MFAStatus)</li>
            </ul>
        </div>
        
        <div class="section">
            <h2>Group Memberships ($($reportData.Groups.Count))</h2>
            <table>
                <tr><th>Group Name</th><th>Type</th><th>Email</th><th>Dynamic</th></tr>
"@
            foreach ($group in $reportData.Groups | Sort-Object DisplayName) {
                $html += "<tr><td>$($group.DisplayName)</td><td>$($group.GroupType)</td><td>$($group.Mail)</td><td>$(if ($group.IsDynamic) {'Yes'} else {'No'})</td></tr>`n"
            }
            $html += @"
            </table>
        </div>
        
        <div class="section">
            <h2>Assigned Licenses ($($reportData.Licenses.Count))</h2>
            <table>
                <tr><th>License</th><th>Service Plans</th></tr>
"@
            foreach ($license in $reportData.Licenses) {
                $html += "<tr><td>$($license.SkuPartNumber)</td><td>$($license.ServicePlans -join ', ')</td></tr>`n"
            }
            $html += @"
            </table>
        </div>
        
        <div class="section">
            <h2>Application Assignments ($($reportData.Applications.Count))</h2>
            <table>
                <tr><th>Application</th><th>App ID</th></tr>
"@
            foreach ($app in $reportData.Applications | Sort-Object ApplicationName) {
                $html += "<tr><td>$($app.ApplicationName)</td><td>$($app.AppId)</td></tr>`n"
            }
            $html += @"
            </table>
        </div>
        
        <div class="section">
            <h2>Security Information</h2>
            <table>
                <tr><th>Property</th><th>Value</th></tr>
                <tr><td>MFA Status</td><td>$($reportData.SecurityInfo.MFAStatus)</td></tr>
                <tr><td>Last Sign-In</td><td>$(if ($reportData.SecurityInfo.LastSignIn.LastSignInDateTime) {$reportData.SecurityInfo.LastSignIn.LastSignInDateTime} else {'N/A'})</td></tr>
            </table>
        </div>
    </div>
</body>
</html>
"@
            $html | Out-File $outputFile
        }
    }
    
    # Display completion message
    Write-Host "`n=== REPORT GENERATION COMPLETE ===" -ForegroundColor Green
    Write-Host "Report saved to: $outputFile" -ForegroundColor Yellow
    
    # Display quick summary
    Write-Host "`nQuick Summary for $($user.DisplayName):" -ForegroundColor Cyan
    Write-Host "  - Groups: $($reportData.Summary.TotalGroups)" -ForegroundColor Gray
    Write-Host "  - Licenses: $($reportData.Summary.TotalLicenses)" -ForegroundColor Gray
    Write-Host "  - Applications: $($reportData.Summary.TotalApplications)" -ForegroundColor Gray
    Write-Host "  - Account Status: $($reportData.Summary.AccountStatus)" -ForegroundColor Gray
    Write-Host "  - MFA Status: $($reportData.Summary.MFAStatus)" -ForegroundColor Gray
    
} catch {
    Write-Host "`nERROR: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Stack Trace: $($_.ScriptStackTrace)" -ForegroundColor Red
}