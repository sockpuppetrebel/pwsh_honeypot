#Requires -Modules Microsoft.Graph.Identity.SignIns, Microsoft.Graph.Identity.DirectoryManagement

<#
.SYNOPSIS
    Comprehensive Azure AD security audit and compliance check
.DESCRIPTION
    Performs security audit of Azure AD configuration including:
    - Privileged user analysis
    - MFA compliance
    - Conditional Access policy effectiveness
    - Risky sign-ins and users
    - Password policy compliance
.PARAMETER DaysToAnalyze
    Number of days of sign-in logs to analyze (default: 7)
.PARAMETER ExportReport
    Export detailed findings to JSON report
.EXAMPLE
    .\Audit-AzureADSecurity.ps1 -DaysToAnalyze 14 -ExportReport
#>

param(
    [Parameter(Mandatory = $false)]
    [int]$DaysToAnalyze = 7,
    
    [Parameter(Mandatory = $false)]
    [switch]$ExportReport
)

# Connect to Microsoft Graph
try {
    $requiredScopes = @(
        "Directory.Read.All",
        "Policy.Read.All",
        "UserAuthenticationMethod.Read.All",
        "IdentityRiskyUser.Read.All",
        "IdentityRiskEvent.Read.All",
        "AuditLog.Read.All"
    )
    
    Connect-MgGraph -Scopes $requiredScopes
    Write-Host "Connected to Microsoft Graph with security audit permissions" -ForegroundColor Green
}
catch {
    Write-Error "Failed to connect to Microsoft Graph: $($_.Exception.Message)"
    exit 1
}

Write-Host "=== AZURE AD SECURITY AUDIT ===" -ForegroundColor Cyan

$auditResults = @{
    AuditDate = Get-Date
    TenantInfo = $null
    PrivilegedUsers = @()
    MFACompliance = @()
    ConditionalAccess = @()
    RiskyUsers = @()
    PasswordPolicy = @()
    SignInAnalysis = @()
    SecurityFindings = @()
    Recommendations = @()
}

# Get tenant information
try {
    $tenantInfo = Get-MgOrganization
    $auditResults.TenantInfo = $tenantInfo | Select-Object DisplayName, Id, VerifiedDomains
    Write-Host "Auditing tenant: $($tenantInfo.DisplayName)" -ForegroundColor Yellow
}
catch {
    Write-Host "Could not retrieve tenant information" -ForegroundColor Red
}

# 1. Privileged User Analysis
Write-Host "`n--- ANALYZING PRIVILEGED USERS ---" -ForegroundColor Yellow

try {
    $privilegedRoles = @(
        "Global Administrator",
        "Privileged Role Administrator", 
        "Security Administrator",
        "User Administrator",
        "Exchange Administrator",
        "SharePoint Administrator",
        "Billing Administrator"
    )
    
    $directoryRoles = Get-MgDirectoryRole -All
    
    foreach ($roleName in $privilegedRoles) {
        $role = $directoryRoles | Where-Object DisplayName -eq $roleName
        if ($role) {
            $roleMembers = Get-MgDirectoryRoleMember -DirectoryRoleId $role.Id
            
            foreach ($member in $roleMembers) {
                if ($member.AdditionalProperties['@odata.type'] -eq '#microsoft.graph.user') {
                    $user = Get-MgUser -UserId $member.Id -Property "Id,UserPrincipalName,DisplayName,AccountEnabled,CreatedDateTime,SignInActivity"
                    
                    $privilegedUserInfo = [PSCustomObject]@{
                        UserPrincipalName = $user.UserPrincipalName
                        DisplayName = $user.DisplayName
                        Role = $roleName
                        AccountEnabled = $user.AccountEnabled
                        LastSignIn = $user.SignInActivity.LastSignInDateTime
                        CreatedDate = $user.CreatedDateTime
                        DaysSinceLastSignIn = if ($user.SignInActivity.LastSignInDateTime) {
                            (Get-Date) - (Get-Date $user.SignInActivity.LastSignInDateTime) | Select-Object -ExpandProperty Days
                        } else { 
                            "Never" 
                        }
                    }
                    
                    $auditResults.PrivilegedUsers += $privilegedUserInfo
                }
            }
        }
    }
    
    Write-Host "Found $($auditResults.PrivilegedUsers.Count) privileged user assignments" -ForegroundColor White
    
    # Flag concerning privileged users
    $stalePrivileged = $auditResults.PrivilegedUsers | Where-Object { 
        $_.DaysSinceLastSignIn -is [int] -and $_.DaysSinceLastSignIn -gt 30 
    }
    
    if ($stalePrivileged) {
        $auditResults.SecurityFindings += "Found $($stalePrivileged.Count) privileged users who haven't signed in for 30+ days"
    }
}
catch {
    Write-Host "Error analyzing privileged users: $($_.Exception.Message)" -ForegroundColor Red
}

# 2. MFA Compliance Analysis
Write-Host "`n--- ANALYZING MFA COMPLIANCE ---" -ForegroundColor Yellow

try {
    $users = Get-MgUser -All -Property "Id,UserPrincipalName,DisplayName,AccountEnabled"
    $mfaStats = @{
        TotalUsers = 0
        MFAEnabled = 0
        MFANotEnabled = 0
        MFACapable = 0
    }
    
    foreach ($user in $users | Where-Object AccountEnabled -eq $true) {
        $mfaStats.TotalUsers++
        
        try {
            $authMethods = Get-MgUserAuthenticationMethod -UserId $user.Id -ErrorAction SilentlyContinue
            
            $hasMFA = $authMethods | Where-Object { 
                $_.AdditionalProperties['@odata.type'] -in @(
                    '#microsoft.graph.microsoftAuthenticatorAuthenticationMethod',
                    '#microsoft.graph.phoneAuthenticationMethod',
                    '#microsoft.graph.fido2AuthenticationMethod'
                )
            }
            
            if ($hasMFA) {
                $mfaStats.MFAEnabled++
                $mfaStatus = "Enabled"
            } else {
                $mfaStats.MFANotEnabled++
                $mfaStatus = "Not Enabled"
            }
            
            $mfaUserInfo = [PSCustomObject]@{
                UserPrincipalName = $user.UserPrincipalName
                DisplayName = $user.DisplayName
                MFAStatus = $mfaStatus
                AuthMethodCount = $authMethods.Count
            }
            
            $auditResults.MFACompliance += $mfaUserInfo
        }
        catch {
            # User authentication methods not accessible
            $mfaStats.MFANotEnabled++
        }
    }
    
    $mfaPercentage = [math]::Round(($mfaStats.MFAEnabled / $mfaStats.TotalUsers) * 100, 2)
    Write-Host "MFA Enabled: $($mfaStats.MFAEnabled)/$($mfaStats.TotalUsers) users ($mfaPercentage%)" -ForegroundColor White
    
    if ($mfaPercentage -lt 95) {
        $auditResults.SecurityFindings += "MFA compliance is $mfaPercentage% - target should be 95%+"
    }
}
catch {
    Write-Host "Error analyzing MFA compliance: $($_.Exception.Message)" -ForegroundColor Red
}

# 3. Conditional Access Policy Analysis
Write-Host "`n--- ANALYZING CONDITIONAL ACCESS POLICIES ---" -ForegroundColor Yellow

try {
    $caPolicies = Get-MgIdentityConditionalAccessPolicy -All
    
    foreach ($policy in $caPolicies) {
        $policyInfo = [PSCustomObject]@{
            DisplayName = $policy.DisplayName
            State = $policy.State
            CreatedDateTime = $policy.CreatedDateTime
            ModifiedDateTime = $policy.ModifiedDateTime
            UserIncludeAll = $policy.Conditions.Users.IncludeUsers -contains "All"
            RequiresMFA = $policy.GrantControls.BuiltInControls -contains "mfa"
            BlocksAccess = $policy.GrantControls.BuiltInControls -contains "block"
            ApplicationCount = $policy.Conditions.Applications.IncludeApplications.Count
        }
        
        $auditResults.ConditionalAccess += $policyInfo
    }
    
    Write-Host "Found $($caPolicies.Count) Conditional Access policies" -ForegroundColor White
    
    $enabledPolicies = $caPolicies | Where-Object State -eq "enabled"
    $mfaPolicies = $enabledPolicies | Where-Object { $_.GrantControls.BuiltInControls -contains "mfa" }
    
    if ($mfaPolicies.Count -eq 0) {
        $auditResults.SecurityFindings += "No Conditional Access policies requiring MFA found"
    }
    
    $reportOnlyPolicies = $caPolicies | Where-Object State -eq "enabledForReportingButNotEnforced"
    if ($reportOnlyPolicies.Count -gt 5) {
        $auditResults.SecurityFindings += "$($reportOnlyPolicies.Count) policies in report-only mode - consider enabling"
    }
}
catch {
    Write-Host "Error analyzing Conditional Access policies: $($_.Exception.Message)" -ForegroundColor Red
}

# 4. Risky Users and Sign-ins
Write-Host "`n--- ANALYZING RISKY USERS AND SIGN-INS ---" -ForegroundColor Yellow

try {
    $riskyUsers = Get-MgIdentityRiskyUser -All
    
    foreach ($riskyUser in $riskyUsers) {
        $riskyUserInfo = [PSCustomObject]@{
            UserPrincipalName = $riskyUser.UserPrincipalName
            RiskLevel = $riskyUser.RiskLevel
            RiskState = $riskyUser.RiskState
            RiskLastUpdatedDateTime = $riskyUser.RiskLastUpdatedDateTime
            RiskDetail = $riskyUser.RiskDetail
        }
        
        $auditResults.RiskyUsers += $riskyUserInfo
    }
    
    Write-Host "Found $($riskyUsers.Count) risky users" -ForegroundColor White
    
    $highRiskUsers = $riskyUsers | Where-Object RiskLevel -eq "high"
    if ($highRiskUsers.Count -gt 0) {
        $auditResults.SecurityFindings += "Found $($highRiskUsers.Count) high-risk users requiring immediate attention"
    }
}
catch {
    Write-Host "Error analyzing risky users: $($_.Exception.Message)" -ForegroundColor Red
}

# 5. Sign-in Analysis
Write-Host "`n--- ANALYZING RECENT SIGN-INS ---" -ForegroundColor Yellow

try {
    $cutoffDate = (Get-Date).AddDays(-$DaysToAnalyze)
    $signIns = Get-MgAuditLogSignIn -Filter "createdDateTime ge $($cutoffDate.ToString('yyyy-MM-ddTHH:mm:ssZ'))" -Top 1000
    
    $signInStats = @{
        TotalSignIns = $signIns.Count
        SuccessfulSignIns = ($signIns | Where-Object { $_.Status.ErrorCode -eq 0 }).Count
        FailedSignIns = ($signIns | Where-Object { $_.Status.ErrorCode -ne 0 }).Count
        RiskySignIns = ($signIns | Where-Object { $_.RiskLevelDuringSignIn -ne "none" }).Count
        MFASignIns = ($signIns | Where-Object { $_.AuthenticationRequirement -eq "multiFactorAuthentication" }).Count
    }
    
    $auditResults.SignInAnalysis = $signInStats
    
    Write-Host "Analyzed $($signInStats.TotalSignIns) sign-ins from the last $DaysToAnalyze days" -ForegroundColor White
    Write-Host "Success rate: $([math]::Round(($signInStats.SuccessfulSignIns / $signInStats.TotalSignIns) * 100, 2))%" -ForegroundColor White
    
    if ($signInStats.RiskySignIns -gt 0) {
        $auditResults.SecurityFindings += "Found $($signInStats.RiskySignIns) risky sign-ins in the last $DaysToAnalyze days"
    }
}
catch {
    Write-Host "Error analyzing sign-ins: $($_.Exception.Message)" -ForegroundColor Red
}

# Generate Recommendations
Write-Host "`n--- GENERATING SECURITY RECOMMENDATIONS ---" -ForegroundColor Cyan

# Privileged user recommendations
if ($auditResults.PrivilegedUsers.Count -gt 10) {
    $auditResults.Recommendations += "Consider implementing Privileged Identity Management (PIM) for $($auditResults.PrivilegedUsers.Count) privileged users"
}

# MFA recommendations
$mfaCompliance = if ($auditResults.MFACompliance.Count -gt 0) {
    ($auditResults.MFACompliance | Where-Object MFAStatus -eq "Enabled").Count / $auditResults.MFACompliance.Count * 100
} else { 0 }

if ($mfaCompliance -lt 95) {
    $auditResults.Recommendations += "Increase MFA adoption from $([math]::Round($mfaCompliance, 1))% to 95%+"
}

# Conditional Access recommendations
if ($auditResults.ConditionalAccess.Count -lt 3) {
    $auditResults.Recommendations += "Implement additional Conditional Access policies for comprehensive protection"
}

# Display security findings
if ($auditResults.SecurityFindings.Count -gt 0) {
    Write-Host "`n--- SECURITY FINDINGS ---" -ForegroundColor Red
    $auditResults.SecurityFindings | ForEach-Object { Write-Host "⚠ $_" -ForegroundColor Yellow }
} else {
    Write-Host "`n--- SECURITY STATUS ---" -ForegroundColor Green
    Write-Host "✓ No critical security issues identified" -ForegroundColor Green
}

# Display recommendations
if ($auditResults.Recommendations.Count -gt 0) {
    Write-Host "`n--- RECOMMENDATIONS ---" -ForegroundColor Cyan
    $auditResults.Recommendations | ForEach-Object { Write-Host "• $_" -ForegroundColor Cyan }
}

# Export report if requested
if ($ExportReport) {
    $timestamp = Get-Date -Format "yyyy-MM-dd-HHmm"
    $reportPath = ".\AzureAD-Security-Audit-$timestamp.json"
    $auditResults | ConvertTo-Json -Depth 5 | Out-File -FilePath $reportPath
    Write-Host "`nDetailed audit report exported to: $reportPath" -ForegroundColor Green
}

Write-Host "`n--- NEXT STEPS ---" -ForegroundColor Cyan
Write-Host "• Review and remediate identified security findings" -ForegroundColor White
Write-Host "• Implement recommended security improvements" -ForegroundColor White
Write-Host "• Schedule regular security audits (monthly recommended)" -ForegroundColor White
Write-Host "• Monitor Azure AD security dashboard regularly" -ForegroundColor White

Disconnect-MgGraph
Write-Host "Security audit completed" -ForegroundColor Green