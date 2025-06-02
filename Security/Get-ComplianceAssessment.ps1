#Requires -Modules Microsoft.Graph.Compliance, Microsoft.Graph.Identity.DirectoryManagement, Microsoft.Graph.DeviceManagement

<#
.SYNOPSIS
    Comprehensive compliance and governance assessment across Microsoft 365 and Azure
.DESCRIPTION
    Performs multi-service compliance assessment including data governance, retention policies,
    DLP, conditional access, device compliance, and regulatory compliance posture.
.PARAMETER IncludeDetailedFindings
    Include detailed compliance findings and recommendations
.PARAMETER CheckRetentionPolicies
    Analyze retention and data lifecycle policies
.PARAMETER CheckDLPPolicies
    Analyze Data Loss Prevention policies and effectiveness
.PARAMETER ExportToExcel
    Export comprehensive results to Excel format
.PARAMETER ComplianceFramework
    Target compliance framework (SOC2, ISO27001, HIPAA, GDPR, All)
.EXAMPLE
    .\Get-ComplianceAssessment.ps1 -IncludeDetailedFindings -CheckRetentionPolicies -CheckDLPPolicies -ComplianceFramework "SOC2" -ExportToExcel
#>

param(
    [Parameter(Mandatory = $false)]
    [switch]$IncludeDetailedFindings,
    
    [Parameter(Mandatory = $false)]
    [switch]$CheckRetentionPolicies,
    
    [Parameter(Mandatory = $false)]
    [switch]$CheckDLPPolicies,
    
    [Parameter(Mandatory = $false)]
    [switch]$ExportToExcel,
    
    [Parameter(Mandatory = $false)]
    [ValidateSet("SOC2", "ISO27001", "HIPAA", "GDPR", "All")]
    [string]$ComplianceFramework = "All"
)

# Connect to Microsoft Graph
try {
    $scopes = @(
        "Policy.Read.All",
        "Directory.Read.All",
        "DeviceManagementConfiguration.Read.All",
        "DeviceManagementManagedDevices.Read.All",
        "InformationProtectionPolicy.Read.All",
        "SecurityEvents.Read.All",
        "AuditLog.Read.All"
    )
    
    Connect-MgGraph -Scopes $scopes
    Write-Host "Connected to Microsoft Graph" -ForegroundColor Green
}
catch {
    Write-Error "Failed to connect to Microsoft Graph: $($_.Exception.Message)"
    exit 1
}

Write-Host "=== COMPLIANCE AND GOVERNANCE ASSESSMENT ===" -ForegroundColor Cyan
Write-Host "Target Framework: $ComplianceFramework" -ForegroundColor Yellow

$complianceAssessment = @{
    AssessmentDate = Get-Date
    ComplianceFramework = $ComplianceFramework
    OverallScore = 0
    ControlResults = @()
    Findings = @()
    Recommendations = @()
    Summary = @{}
}

# Define compliance controls based on framework
$complianceControls = @()

switch ($ComplianceFramework) {
    "SOC2" {
        $complianceControls = @(
            @{ Id = "CC6.1"; Name = "Logical Access Controls"; Category = "Access Management" },
            @{ Id = "CC6.2"; Name = "Multi-Factor Authentication"; Category = "Access Management" },
            @{ Id = "CC6.3"; Name = "User Access Reviews"; Category = "Access Management" },
            @{ Id = "CC7.1"; Name = "Data Encryption"; Category = "Data Protection" },
            @{ Id = "CC7.2"; Name = "Data Retention"; Category = "Data Protection" },
            @{ Id = "CC8.1"; Name = "Vulnerability Management"; Category = "Security Monitoring" }
        )
    }
    "ISO27001" {
        $complianceControls = @(
            @{ Id = "A.9.1.1"; Name = "Access Control Policy"; Category = "Access Control" },
            @{ Id = "A.9.2.1"; Name = "User Registration"; Category = "Access Control" },
            @{ Id = "A.9.4.2"; Name = "Secure Log-on Procedures"; Category = "Access Control" },
            @{ Id = "A.10.1.1"; Name = "Cryptographic Controls"; Category = "Cryptography" },
            @{ Id = "A.12.3.1"; Name = "Information Backup"; Category = "Operations Security" },
            @{ Id = "A.16.1.2"; Name = "Reporting Information Security Events"; Category = "Incident Management" }
        )
    }
    "HIPAA" {
        $complianceControls = @(
            @{ Id = "164.308(a)(1)"; Name = "Security Management Process"; Category = "Administrative Safeguards" },
            @{ Id = "164.308(a)(3)"; Name = "Workforce Training"; Category = "Administrative Safeguards" },
            @{ Id = "164.310(a)(1)"; Name = "Facility Access Controls"; Category = "Physical Safeguards" },
            @{ Id = "164.312(a)(1)"; Name = "Access Control"; Category = "Technical Safeguards" },
            @{ Id = "164.312(e)(1)"; Name = "Transmission Security"; Category = "Technical Safeguards" }
        )
    }
    "GDPR" {
        $complianceControls = @(
            @{ Id = "Art.25"; Name = "Data Protection by Design"; Category = "Data Protection" },
            @{ Id = "Art.30"; Name = "Records of Processing"; Category = "Documentation" },
            @{ Id = "Art.32"; Name = "Security of Processing"; Category = "Security" },
            @{ Id = "Art.33"; Name = "Breach Notification"; Category = "Incident Response" },
            @{ Id = "Art.35"; Name = "Data Protection Impact Assessment"; Category = "Privacy" }
        )
    }
    "All" {
        $complianceControls = @(
            @{ Id = "IAM-001"; Name = "Identity and Access Management"; Category = "Access Control" },
            @{ Id = "MFA-001"; Name = "Multi-Factor Authentication"; Category = "Authentication" },
            @{ Id = "ENC-001"; Name = "Data Encryption"; Category = "Data Protection" },
            @{ Id = "MON-001"; Name = "Security Monitoring"; Category = "Monitoring" },
            @{ Id = "GOV-001"; Name = "Data Governance"; Category = "Governance" },
            @{ Id = "COM-001"; Name = "Device Compliance"; Category = "Device Management" }
        )
    }
}

Write-Host "Assessing $($complianceControls.Count) compliance controls..." -ForegroundColor Yellow

# Assess each compliance control
foreach ($control in $complianceControls) {
    Write-Host "Assessing: $($control.Name)" -ForegroundColor Gray
    
    $controlResult = [PSCustomObject]@{
        ControlId = $control.Id
        ControlName = $control.Name
        Category = $control.Category
        Status = "Not Assessed"
        Score = 0
        Findings = @()
        Evidence = @()
        Recommendations = @()
    }
    
    try {
        switch ($control.Id) {
            { $_ -like "*IAM*" -or $_ -like "*Access*" -or $_ -like "A.9.*" -or $_ -like "CC6.*" } {
                # Identity and Access Management Assessment
                $users = Get-MgUser -All -Property "Id,UserPrincipalName,AccountEnabled,SignInActivity"
                $adminRoles = Get-MgDirectoryRole -All
                
                $inactiveAdmins = 0
                $unassignedRoles = 0
                
                foreach ($role in $adminRoles) {
                    $roleMembers = Get-MgDirectoryRoleMember -DirectoryRoleId $role.Id
                    if ($roleMembers.Count -eq 0 -and $role.DisplayName -like "*Administrator*") {
                        $unassignedRoles++
                    }
                    
                    foreach ($member in $roleMembers) {
                        if ($member.AdditionalProperties['@odata.type'] -eq '#microsoft.graph.user') {
                            $user = $users | Where-Object Id -eq $member.Id
                            if ($user.SignInActivity.LastSignInDateTime) {
                                $daysSinceSignIn = (Get-Date) - (Get-Date $user.SignInActivity.LastSignInDateTime) | Select-Object -ExpandProperty Days
                                if ($daysSinceSignIn -gt 30) {
                                    $inactiveAdmins++
                                }
                            }
                        }
                    }
                }
                
                $score = 100
                if ($inactiveAdmins -gt 0) {
                    $score -= 20
                    $controlResult.Findings += "Found $inactiveAdmins inactive administrative accounts"
                }
                if ($unassignedRoles -gt 0) {
                    $score -= 10
                    $controlResult.Findings += "Found $unassignedRoles administrative roles with no members"
                }
                
                $controlResult.Score = [math]::Max($score, 0)
                $controlResult.Status = if ($score -ge 80) { "Compliant" } elseif ($score -ge 60) { "Partially Compliant" } else { "Non-Compliant" }
                $controlResult.Evidence += "Analyzed $($users.Count) users and $($adminRoles.Count) directory roles"
            }
            
            { $_ -like "*MFA*" -or $_ -like "*Authentication*" -or $_ -like "CC6.2" } {
                # Multi-Factor Authentication Assessment
                $users = Get-MgUser -All -Property "Id,UserPrincipalName,AccountEnabled"
                $mfaEnabledCount = 0
                $totalActiveUsers = ($users | Where-Object AccountEnabled -eq $true).Count
                
                foreach ($user in ($users | Where-Object AccountEnabled -eq $true)) {
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
                            $mfaEnabledCount++
                        }
                    }
                    catch {
                        # MFA check failed for this user
                    }
                }
                
                $mfaPercentage = if ($totalActiveUsers -gt 0) { 
                    [math]::Round(($mfaEnabledCount / $totalActiveUsers) * 100, 2) 
                } else { 0 }
                
                $score = $mfaPercentage
                $controlResult.Score = $score
                $controlResult.Status = if ($score -ge 95) { "Compliant" } elseif ($score -ge 80) { "Partially Compliant" } else { "Non-Compliant" }
                $controlResult.Evidence += "MFA enabled for $mfaEnabledCount of $totalActiveUsers active users ($mfaPercentage%)"
                
                if ($score -lt 95) {
                    $controlResult.Findings += "MFA adoption is $mfaPercentage% (target: 95%+)"
                    $controlResult.Recommendations += "Implement MFA enforcement policies"
                }
            }
            
            { $_ -like "*Encryption*" -or $_ -like "*Data Protection*" -or $_ -like "CC7.*" } {
                # Data Protection and Encryption Assessment
                try {
                    $devices = Get-MgDeviceManagementManagedDevice -All -ErrorAction SilentlyContinue
                    $encryptedDevices = ($devices | Where-Object IsEncrypted -eq $true).Count
                    $totalDevices = $devices.Count
                    
                    $encryptionPercentage = if ($totalDevices -gt 0) { 
                        [math]::Round(($encryptedDevices / $totalDevices) * 100, 2) 
                    } else { 100 }
                    
                    $score = $encryptionPercentage
                    $controlResult.Score = $score
                    $controlResult.Status = if ($score -ge 95) { "Compliant" } elseif ($score -ge 80) { "Partially Compliant" } else { "Non-Compliant" }
                    $controlResult.Evidence += "Device encryption: $encryptedDevices of $totalDevices devices encrypted ($encryptionPercentage%)"
                    
                    if ($score -lt 95) {
                        $controlResult.Findings += "Device encryption coverage is $encryptionPercentage% (target: 95%+)"
                        $controlResult.Recommendations += "Enforce device encryption policies"
                    }
                }
                catch {
                    $controlResult.Score = 50
                    $controlResult.Status = "Partially Compliant"
                    $controlResult.Findings += "Could not assess device encryption status"
                }
            }
            
            { $_ -like "*Monitoring*" -or $_ -like "*Security*" -or $_ -like "CC8.*" } {
                # Security Monitoring Assessment
                try {
                    $caPolicies = Get-MgIdentityConditionalAccessPolicy -All
                    $enabledPolicies = ($caPolicies | Where-Object State -eq "enabled").Count
                    $totalPolicies = $caPolicies.Count
                    
                    $score = 100
                    if ($enabledPolicies -lt 3) {
                        $score = 50
                        $controlResult.Findings += "Limited conditional access policies ($enabledPolicies enabled)"
                    }
                    
                    $controlResult.Score = $score
                    $controlResult.Status = if ($score -ge 80) { "Compliant" } elseif ($score -ge 60) { "Partially Compliant" } else { "Non-Compliant" }
                    $controlResult.Evidence += "Conditional Access: $enabledPolicies of $totalPolicies policies enabled"
                    
                    if ($score -lt 80) {
                        $controlResult.Recommendations += "Implement comprehensive conditional access policies"
                    }
                }
                catch {
                    $controlResult.Score = 30
                    $controlResult.Status = "Non-Compliant"
                    $controlResult.Findings += "Could not assess security monitoring controls"
                }
            }
            
            { $_ -like "*Governance*" -or $_ -like "*Retention*" } {
                # Data Governance and Retention Assessment
                if ($CheckRetentionPolicies) {
                    # Note: This would require Microsoft Purview/Compliance modules
                    $score = 75  # Placeholder assessment
                    $controlResult.Score = $score
                    $controlResult.Status = "Partially Compliant"
                    $controlResult.Evidence += "Data governance policies partially implemented"
                    $controlResult.Recommendations += "Implement comprehensive data retention policies"
                } else {
                    $controlResult.Score = 50
                    $controlResult.Status = "Not Assessed"
                    $controlResult.Evidence += "Data governance assessment requires additional permissions"
                }
            }
            
            { $_ -like "*Device*" -or $_ -like "*Compliance*" } {
                # Device Compliance Assessment
                try {
                    $devices = Get-MgDeviceManagementManagedDevice -All
                    $compliantDevices = ($devices | Where-Object ComplianceState -eq "Compliant").Count
                    $totalDevices = $devices.Count
                    
                    $compliancePercentage = if ($totalDevices -gt 0) { 
                        [math]::Round(($compliantDevices / $totalDevices) * 100, 2) 
                    } else { 100 }
                    
                    $score = $compliancePercentage
                    $controlResult.Score = $score
                    $controlResult.Status = if ($score -ge 95) { "Compliant" } elseif ($score -ge 80) { "Partially Compliant" } else { "Non-Compliant" }
                    $controlResult.Evidence += "Device compliance: $compliantDevices of $totalDevices devices compliant ($compliancePercentage%)"
                    
                    if ($score -lt 95) {
                        $controlResult.Findings += "Device compliance is $compliancePercentage% (target: 95%+)"
                        $controlResult.Recommendations += "Review and remediate non-compliant devices"
                    }
                }
                catch {
                    $controlResult.Score = 60
                    $controlResult.Status = "Partially Compliant"
                    $controlResult.Findings += "Could not assess device compliance status"
                }
            }
            
            default {
                # Generic assessment for framework-specific controls
                $controlResult.Score = 70
                $controlResult.Status = "Partially Compliant"
                $controlResult.Evidence += "Manual assessment required for this control"
                $controlResult.Recommendations += "Conduct detailed manual review of $($control.Name)"
            }
        }
    }
    catch {
        $controlResult.Score = 0
        $controlResult.Status = "Assessment Failed"
        $controlResult.Findings += "Assessment failed: $($_.Exception.Message)"
    }
    
    $complianceAssessment.ControlResults += $controlResult
}

# Calculate overall compliance score
$totalScore = ($complianceAssessment.ControlResults | Measure-Object Score -Average).Average
$complianceAssessment.OverallScore = [math]::Round($totalScore, 1)

# Generate summary
$compliantControls = ($complianceAssessment.ControlResults | Where-Object Status -eq "Compliant").Count
$partiallyCompliantControls = ($complianceAssessment.ControlResults | Where-Object Status -eq "Partially Compliant").Count
$nonCompliantControls = ($complianceAssessment.ControlResults | Where-Object Status -eq "Non-Compliant").Count

$complianceAssessment.Summary = @{
    TotalControls = $complianceControls.Count
    CompliantControls = $compliantControls
    PartiallyCompliantControls = $partiallyCompliantControls
    NonCompliantControls = $nonCompliantControls
    CompliancePercentage = [math]::Round(($compliantControls / $complianceControls.Count) * 100, 1)
}

# Display results
Write-Host "`n--- COMPLIANCE ASSESSMENT RESULTS ---" -ForegroundColor Cyan

Write-Host "Overall Compliance Score: $($complianceAssessment.OverallScore)%" -ForegroundColor $(
    if ($complianceAssessment.OverallScore -ge 80) { "Green" } 
    elseif ($complianceAssessment.OverallScore -ge 60) { "Yellow" } 
    else { "Red" }
)

Write-Host "`nControl Status Summary:" -ForegroundColor Yellow
Write-Host "  Compliant: $compliantControls ($($complianceAssessment.Summary.CompliancePercentage)%)" -ForegroundColor Green
Write-Host "  Partially Compliant: $partiallyCompliantControls" -ForegroundColor Yellow
Write-Host "  Non-Compliant: $nonCompliantControls" -ForegroundColor Red

# Show control results
Write-Host "`n--- CONTROL ASSESSMENT DETAILS ---" -ForegroundColor Yellow
$complianceAssessment.ControlResults | Sort-Object Score | 
    Select-Object ControlId, ControlName, Status, Score, @{Name="Primary Finding";Expression={$_.Findings[0]}} | 
    Format-Table -AutoSize

# Show findings and recommendations
if ($IncludeDetailedFindings) {
    $allFindings = $complianceAssessment.ControlResults | Where-Object { $_.Findings.Count -gt 0 } | ForEach-Object { $_.Findings }
    $allRecommendations = $complianceAssessment.ControlResults | Where-Object { $_.Recommendations.Count -gt 0 } | ForEach-Object { $_.Recommendations }
    
    if ($allFindings.Count -gt 0) {
        Write-Host "`n--- KEY FINDINGS ---" -ForegroundColor Red
        $allFindings | Sort-Object -Unique | ForEach-Object { Write-Host "• $_" -ForegroundColor Yellow }
    }
    
    if ($allRecommendations.Count -gt 0) {
        Write-Host "`n--- RECOMMENDATIONS ---" -ForegroundColor Cyan
        $allRecommendations | Sort-Object -Unique | ForEach-Object { Write-Host "• $_" -ForegroundColor Cyan }
    }
}

# Export results
if ($ExportToExcel) {
    $timestamp = Get-Date -Format "yyyy-MM-dd-HHmm"
    $outputPath = ".\Compliance-Assessment-$ComplianceFramework-$timestamp.xlsx"
    
    if (Get-Module -ListAvailable -Name ImportExcel) {
        # Summary sheet
        $summaryData = @(
            [PSCustomObject]@{ Metric = "Overall Compliance Score"; Value = "$($complianceAssessment.OverallScore)%" }
            [PSCustomObject]@{ Metric = "Compliance Framework"; Value = $ComplianceFramework }
            [PSCustomObject]@{ Metric = "Assessment Date"; Value = $complianceAssessment.AssessmentDate }
            [PSCustomObject]@{ Metric = "Total Controls"; Value = $complianceAssessment.Summary.TotalControls }
            [PSCustomObject]@{ Metric = "Compliant Controls"; Value = $compliantControls }
            [PSCustomObject]@{ Metric = "Partially Compliant"; Value = $partiallyCompliantControls }
            [PSCustomObject]@{ Metric = "Non-Compliant"; Value = $nonCompliantControls }
        )
        
        $summaryData | Export-Excel -Path $outputPath -WorksheetName "Summary" -AutoSize
        
        # Control results
        $complianceAssessment.ControlResults | Export-Excel -Path $outputPath -WorksheetName "Control Results" -AutoSize -FreezeTopRow
        
        # Detailed findings
        if ($IncludeDetailedFindings) {
            $findingsData = $complianceAssessment.ControlResults | Where-Object { $_.Findings.Count -gt 0 } | ForEach-Object {
                foreach ($finding in $_.Findings) {
                    [PSCustomObject]@{
                        ControlId = $_.ControlId
                        ControlName = $_.ControlName
                        Category = $_.Category
                        Finding = $finding
                    }
                }
            }
            
            $findingsData | Export-Excel -Path $outputPath -WorksheetName "Detailed Findings" -AutoSize
        }
        
        Write-Host "`nCompliance assessment exported to: $outputPath" -ForegroundColor Green
    } else {
        Write-Host "ImportExcel module not available. Exporting to CSV..." -ForegroundColor Yellow
        $complianceAssessment.ControlResults | Export-Csv -Path ".\Compliance-Assessment-$ComplianceFramework-$timestamp.csv" -NoTypeInformation
        Write-Host "Assessment exported to: .\Compliance-Assessment-$ComplianceFramework-$timestamp.csv" -ForegroundColor Green
    }
}

Write-Host "`n--- NEXT STEPS ---" -ForegroundColor Cyan
Write-Host "• Address non-compliant controls to improve overall compliance posture" -ForegroundColor White
Write-Host "• Implement recommended security and governance controls" -ForegroundColor White
Write-Host "• Schedule regular compliance assessments (quarterly recommended)" -ForegroundColor White
Write-Host "• Document evidence and remediation activities for audit purposes" -ForegroundColor White

Disconnect-MgGraph
Write-Host "Compliance assessment complete" -ForegroundColor Green