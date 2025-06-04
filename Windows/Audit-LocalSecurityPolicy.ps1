<#
.SYNOPSIS
Audits local security policy settings across Windows endpoints

.DESCRIPTION
Performs comprehensive audit of local security policies including password policy,
account lockout policy, user rights assignments, security options, and audit policies.
Compares against security baselines and generates compliance reports.

.PARAMETER ComputerName
Target computer(s) to audit. Defaults to local machine.

.PARAMETER ExportPath
Path to save audit reports. Creates timestamped folder if not specified.

.PARAMETER Baseline
Security baseline to compare against: CIS, NIST, DISA, Custom

.PARAMETER BaselinePath
Path to custom baseline file (JSON format)

.PARAMETER IncludeRecommendations
Include security recommendations in the report

.EXAMPLE
Audit-LocalSecurityPolicy
Audits local machine security policy

.EXAMPLE
Audit-LocalSecurityPolicy -ComputerName "SERVER01","WS02" -Baseline CIS -IncludeRecommendations
Audits multiple machines against CIS baseline with recommendations

.NOTES
Author: Enterprise PowerShell Collection
Requires: Administrative privileges, secedit.exe
Version: 1.0
#>

[CmdletBinding()]
param(
    [Parameter(ValueFromPipeline = $true)]
    [string[]]$ComputerName = $env:COMPUTERNAME,
    
    [Parameter()]
    [string]$ExportPath,
    
    [Parameter()]
    [ValidateSet("CIS", "NIST", "DISA", "Custom")]
    [string]$Baseline,
    
    [Parameter()]
    [string]$BaselinePath,
    
    [Parameter()]
    [switch]$IncludeRecommendations
)

begin {
    Write-Host "Starting Local Security Policy Audit..." -ForegroundColor Cyan
    
    # Setup export directory
    if (-not $ExportPath) {
        $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
        $ExportPath = "SecurityAudit_$timestamp"
    }
    
    if (-not (Test-Path $ExportPath)) {
        New-Item -Path $ExportPath -ItemType Directory -Force | Out-Null
    }
    
    # Define security baselines
    $baselines = @{
        CIS = @{
            PasswordHistorySize = 24
            MaxPasswordAge = 365
            MinPasswordAge = 1
            MinPasswordLength = 14
            ComplexityEnabled = $true
            LockoutThreshold = 10
            LockoutDuration = 15
            ResetLockoutCounterAfter = 15
        }
        NIST = @{
            PasswordHistorySize = 12
            MaxPasswordAge = 365
            MinPasswordAge = 1
            MinPasswordLength = 12
            ComplexityEnabled = $true
            LockoutThreshold = 5
            LockoutDuration = 30
            ResetLockoutCounterAfter = 30
        }
    }
    
    $results = @()
}

process {
    foreach ($computer in $ComputerName) {
        Write-Host "Auditing security policy: $computer" -ForegroundColor Yellow
        
        try {
            $computerResult = [PSCustomObject]@{
                ComputerName = $computer
                Timestamp = Get-Date
                PasswordPolicy = $null
                AccountLockoutPolicy = $null
                UserRights = $null
                SecurityOptions = $null
                AuditPolicy = $null
                ComplianceScore = 0
                Issues = @()
                Recommendations = @()
                Status = "Success"
            }
            
            # Export security policy
            Write-Progress -Activity "Security Audit: $computer" -Status "Exporting security policy" -PercentComplete 20
            
            $tempPath = "\\$computer\c$\temp\secpol_audit.inf"
            $localTempPath = "C:\temp\secpol_audit.inf"
            
            Invoke-Command -ComputerName $computer -ScriptBlock {
                param($LocalPath)
                if (-not (Test-Path "C:\temp")) { New-Item -Path "C:\temp" -ItemType Directory -Force }
                secedit.exe /export /cfg $LocalPath | Out-Null
            } -ArgumentList $localTempPath
            
            # Read and parse security policy
            if (Test-Path $tempPath) {
                $secPol = Get-Content $tempPath | Where-Object { $_ -and $_ -notmatch '^\s*$' -and $_ -notmatch '^\[' }
                
                # Parse password policy
                $passwordPolicy = @{}
                $passwordPolicy.PasswordHistorySize = [int]($secPol | Where-Object { $_ -match 'PasswordHistorySize' } | ForEach-Object { ($_ -split '=')[1].Trim() })
                $passwordPolicy.MaxPasswordAge = [int]($secPol | Where-Object { $_ -match 'MaximumPasswordAge' } | ForEach-Object { ($_ -split '=')[1].Trim() })
                $passwordPolicy.MinPasswordAge = [int]($secPol | Where-Object { $_ -match 'MinimumPasswordAge' } | ForEach-Object { ($_ -split '=')[1].Trim() })
                $passwordPolicy.MinPasswordLength = [int]($secPol | Where-Object { $_ -match 'MinimumPasswordLength' } | ForEach-Object { ($_ -split '=')[1].Trim() })
                $passwordPolicy.ComplexityEnabled = [bool][int]($secPol | Where-Object { $_ -match 'PasswordComplexity' } | ForEach-Object { ($_ -split '=')[1].Trim() })
                
                $computerResult.PasswordPolicy = $passwordPolicy
                
                # Parse account lockout policy
                $lockoutPolicy = @{}
                $lockoutPolicy.LockoutThreshold = [int]($secPol | Where-Object { $_ -match 'LockoutBadCount' } | ForEach-Object { ($_ -split '=')[1].Trim() })
                $lockoutPolicy.LockoutDuration = [int]($secPol | Where-Object { $_ -match 'LockoutDuration' } | ForEach-Object { ($_ -split '=')[1].Trim() }) / 60
                $lockoutPolicy.ResetLockoutCounterAfter = [int]($secPol | Where-Object { $_ -match 'ResetLockoutCount' } | ForEach-Object { ($_ -split '=')[1].Trim() }) / 60
                
                $computerResult.AccountLockoutPolicy = $lockoutPolicy
                
                # Get user rights assignments
                Write-Progress -Activity "Security Audit: $computer" -Status "Analyzing user rights" -PercentComplete 50
                
                $userRights = @{}
                $privilegeLines = $secPol | Where-Object { $_ -match 'Se\w+Privilege' }
                foreach ($line in $privilegeLines) {
                    $parts = $line -split '='
                    if ($parts.Length -eq 2) {
                        $privilege = $parts[0].Trim()
                        $users = $parts[1].Trim() -split ','
                        $userRights[$privilege] = $users
                    }
                }
                $computerResult.UserRights = $userRights
                
                # Cleanup temp file
                Remove-Item $tempPath -Force -ErrorAction SilentlyContinue
            }
            
            # Get audit policy
            Write-Progress -Activity "Security Audit: $computer" -Status "Checking audit policy" -PercentComplete 70
            
            $auditPolicy = Invoke-Command -ComputerName $computer -ScriptBlock {
                $auditSettings = @{}
                try {
                    $auditOutput = auditpol.exe /get /category:*
                    foreach ($line in $auditOutput) {
                        if ($line -match '^\s*(.+?)\s+(Success and Failure|Success|Failure|No Auditing)\s*$') {
                            $auditSettings[$matches[1].Trim()] = $matches[2].Trim()
                        }
                    }
                } catch {
                    $auditSettings["Error"] = $_.Exception.Message
                }
                return $auditSettings
            }
            $computerResult.AuditPolicy = $auditPolicy
            
            # Baseline comparison
            Write-Progress -Activity "Security Audit: $computer" -Status "Comparing against baseline" -PercentComplete 85
            
            if ($Baseline -and $baselines.ContainsKey($Baseline)) {
                $baselineConfig = $baselines[$Baseline]
                $score = 100
                
                # Password policy checks
                if ($passwordPolicy.PasswordHistorySize -lt $baselineConfig.PasswordHistorySize) {
                    $computerResult.Issues += "Password history size ($($passwordPolicy.PasswordHistorySize)) below baseline ($($baselineConfig.PasswordHistorySize))"
                    $score -= 10
                }
                
                if ($passwordPolicy.MinPasswordLength -lt $baselineConfig.MinPasswordLength) {
                    $computerResult.Issues += "Minimum password length ($($passwordPolicy.MinPasswordLength)) below baseline ($($baselineConfig.MinPasswordLength))"
                    $score -= 15
                }
                
                if (-not $passwordPolicy.ComplexityEnabled -and $baselineConfig.ComplexityEnabled) {
                    $computerResult.Issues += "Password complexity not enabled"
                    $score -= 20
                }
                
                # Account lockout checks
                if ($lockoutPolicy.LockoutThreshold -eq 0 -or $lockoutPolicy.LockoutThreshold -gt $baselineConfig.LockoutThreshold) {
                    $computerResult.Issues += "Account lockout threshold ($($lockoutPolicy.LockoutThreshold)) exceeds baseline ($($baselineConfig.LockoutThreshold))"
                    $score -= 15
                }
                
                $computerResult.ComplianceScore = [Math]::Max($score, 0)
            }
            
            # Generate recommendations
            if ($IncludeRecommendations) {
                if ($passwordPolicy.MinPasswordLength -lt 12) {
                    $computerResult.Recommendations += "Increase minimum password length to at least 12 characters"
                }
                if (-not $passwordPolicy.ComplexityEnabled) {
                    $computerResult.Recommendations += "Enable password complexity requirements"
                }
                if ($lockoutPolicy.LockoutThreshold -eq 0) {
                    $computerResult.Recommendations += "Enable account lockout policy to prevent brute force attacks"
                }
                if (-not $auditPolicy.ContainsKey("Logon/Logoff") -or $auditPolicy["Logon/Logoff"] -eq "No Auditing") {
                    $computerResult.Recommendations += "Enable auditing for logon/logoff events"
                }
            }
            
            $results += $computerResult
            Write-Host "  ✓ Completed audit for $computer" -ForegroundColor Green
            
        } catch {
            Write-Error "Failed to audit $computer : $_"
            $results += [PSCustomObject]@{
                ComputerName = $computer
                Status = "Error"
                Error = $_.Exception.Message
                Timestamp = Get-Date
            }
        }
        
        Write-Progress -Activity "Security Audit: $computer" -Completed
    }
}

end {
    # Export results
    Write-Host "`nExporting audit results..." -ForegroundColor Cyan
    
    $results | Export-Clixml -Path "$ExportPath\SecurityAuditResults.xml"
    $results | ConvertTo-Json -Depth 5 | Out-File "$ExportPath\SecurityAuditResults.json"
    
    # Create summary report
    $summary = $results | Where-Object Status -eq "Success" | Select-Object ComputerName, 
        @{N="MinPasswordLength";E={$_.PasswordPolicy.MinPasswordLength}},
        @{N="ComplexityEnabled";E={$_.PasswordPolicy.ComplexityEnabled}},
        @{N="LockoutThreshold";E={$_.AccountLockoutPolicy.LockoutThreshold}},
        @{N="ComplianceScore";E={$_.ComplianceScore}},
        @{N="IssueCount";E={$_.Issues.Count}}
    
    $summary | Export-Csv -Path "$ExportPath\SecurityAuditSummary.csv" -NoTypeInformation
    
    # Display summary
    Write-Host "`n" + "="*80 -ForegroundColor Cyan
    Write-Host "SECURITY POLICY AUDIT SUMMARY" -ForegroundColor Cyan
    Write-Host "="*80 -ForegroundColor Cyan
    
    $summary | Format-Table -AutoSize
    
    # Compliance statistics
    if ($Baseline) {
        $avgScore = [math]::Round(($summary | Measure-Object ComplianceScore -Average).Average, 1)
        $compliantSystems = ($summary | Where-Object ComplianceScore -ge 80).Count
        
        Write-Host "`nCompliance Statistics ($Baseline Baseline):" -ForegroundColor Yellow
        Write-Host "  Average Compliance Score: $avgScore%" -ForegroundColor White
        Write-Host "  Compliant Systems (≥80%): $compliantSystems/$($summary.Count)" -ForegroundColor $(if ($compliantSystems -eq $summary.Count) {"Green"} else {"Red"})
    }
    
    # Top issues
    $allIssues = $results | ForEach-Object { $_.Issues } | Group-Object | Sort-Object Count -Descending | Select-Object -First 5
    if ($allIssues) {
        Write-Host "`nMost Common Issues:" -ForegroundColor Yellow
        $allIssues | ForEach-Object {
            Write-Host "  • $($_.Name) ($($_.Count) systems)" -ForegroundColor Red
        }
    }
    
    # Top recommendations
    if ($IncludeRecommendations) {
        $allRecommendations = $results | ForEach-Object { $_.Recommendations } | Group-Object | Sort-Object Count -Descending | Select-Object -First 5
        if ($allRecommendations) {
            Write-Host "`nTop Recommendations:" -ForegroundColor Yellow
            $allRecommendations | ForEach-Object {
                Write-Host "  • $($_.Name) ($($_.Count) systems)" -ForegroundColor Cyan
            }
        }
    }
    
    Write-Host "`nDetailed audit results saved to: $ExportPath" -ForegroundColor Green
    Write-Host "Security audit completed for $($results.Count) systems" -ForegroundColor Cyan
}