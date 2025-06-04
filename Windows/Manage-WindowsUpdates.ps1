<#
.SYNOPSIS
Comprehensive Windows Update management across enterprise endpoints

.DESCRIPTION
Provides centralized Windows Update management including status checking,
update installation, reboot scheduling, and compliance reporting. Supports
both WSUS and Windows Update for Business environments.

.PARAMETER ComputerName
Target computer(s) for update management. Defaults to local machine.

.PARAMETER Action
Action to perform: Check, Install, Download, Reboot, Report, Configure

.PARAMETER UpdateCategories
Specific update categories to target: Security, Critical, Important, Optional, Drivers

.PARAMETER ExcludeKBs
KB numbers to exclude from installation (comma-separated)

.PARAMETER ScheduleReboot
Schedule automatic reboot after installation (hours from now)

.PARAMETER ExportPath
Path to save update reports and logs

.PARAMETER Force
Skip confirmation prompts for installation and reboots

.EXAMPLE
Manage-WindowsUpdates -Action Check
Checks update status on local machine

.EXAMPLE
Manage-WindowsUpdates -ComputerName "SERVER01","WS02" -Action Install -UpdateCategories Security,Critical -ScheduleReboot 2
Installs security and critical updates on multiple machines with 2-hour reboot delay

.EXAMPLE
Manage-WindowsUpdates -Action Report -ExportPath "C:\Reports" -ComputerName (Get-Content servers.txt)
Generates compliance report for all servers

.NOTES
Author: Enterprise PowerShell Collection
Requires: Administrative privileges, PSWindowsUpdate module (optional)
Version: 1.0
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(ValueFromPipeline = $true)]
    [string[]]$ComputerName = $env:COMPUTERNAME,
    
    [Parameter(Mandatory)]
    [ValidateSet("Check", "Install", "Download", "Reboot", "Report", "Configure")]
    [string]$Action,
    
    [Parameter()]
    [ValidateSet("Security", "Critical", "Important", "Optional", "Drivers", "All")]
    [string[]]$UpdateCategories = @("Security", "Critical"),
    
    [Parameter()]
    [string[]]$ExcludeKBs,
    
    [Parameter()]
    [int]$ScheduleReboot,
    
    [Parameter()]
    [string]$ExportPath,
    
    [Parameter()]
    [switch]$Force
)

begin {
    Write-Host "Starting Windows Update Management - Action: $Action" -ForegroundColor Cyan
    
    # Setup export directory
    if ($ExportPath -and -not (Test-Path $ExportPath)) {
        New-Item -Path $ExportPath -ItemType Directory -Force | Out-Null
    }
    
    $results = @()
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    
    # Check for PSWindowsUpdate module
    $hasPSWU = Get-Module -ListAvailable -Name PSWindowsUpdate -ErrorAction SilentlyContinue
    if (-not $hasPSWU -and $Action -in @("Install", "Download")) {
        Write-Warning "PSWindowsUpdate module not found. Using built-in methods (limited functionality)."
    }
}

process {
    foreach ($computer in $ComputerName) {
        Write-Host "Processing $computer..." -ForegroundColor Yellow
        
        try {
            $computerResult = [PSCustomObject]@{
                ComputerName = $computer
                Timestamp = Get-Date
                Action = $Action
                UpdatesAvailable = 0
                UpdatesInstalled = 0
                UpdatesPending = 0
                RebootRequired = $false
                LastUpdateDate = $null
                UpdateDetails = @()
                Status = "Unknown"
                Error = $null
            }
            
            # Test connectivity
            if ($computer -ne $env:COMPUTERNAME -and -not (Test-Connection -ComputerName $computer -Count 1 -Quiet)) {
                throw "Cannot connect to $computer"
            }
            
            switch ($Action) {
                "Check" {
                    Write-Progress -Activity "Windows Updates: $computer" -Status "Checking for updates" -PercentComplete 25
                    
                    # Check update status using WUA COM object
                    $updateSession = Invoke-Command -ComputerName $computer -ScriptBlock {
                        try {
                            $session = New-Object -ComObject Microsoft.Update.Session
                            $searcher = $session.CreateUpdateSearcher()
                            $searchResult = $searcher.Search("IsInstalled=0 and IsHidden=0")
                            
                            $updates = @()
                            foreach ($update in $searchResult.Updates) {
                                $categories = @()
                                foreach ($category in $update.Categories) {
                                    $categories += $category.Name
                                }
                                
                                $updates += [PSCustomObject]@{
                                    Title = $update.Title
                                    Description = $update.Description
                                    Size = [math]::Round($update.MaxDownloadSize / 1MB, 2)
                                    Categories = $categories -join ", "
                                    Severity = if ($update.MsrcSeverity) { $update.MsrcSeverity } else { "Unknown" }
                                    KBArticleIDs = $update.KBArticleIDs -join ", "
                                    IsDownloaded = $update.IsDownloaded
                                    RebootRequired = $update.InstallationBehavior.RebootBehavior -ne 0
                                }
                            }
                            
                            return @{
                                UpdateCount = $searchResult.Updates.Count
                                Updates = $updates
                                LastSearchSuccessDate = $searcher.GetTotalHistoryCount()
                            }
                        } catch {
                            throw "Error accessing Windows Update: $($_.Exception.Message)"
                        }
                    } -ErrorAction Stop
                    
                    $computerResult.UpdatesAvailable = $updateSession.UpdateCount
                    $computerResult.UpdateDetails = $updateSession.Updates
                    $computerResult.Status = "Checked"
                    
                    # Check reboot status
                    $rebootPending = Invoke-Command -ComputerName $computer -ScriptBlock {
                        $cbsReboot = Get-ChildItem "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending" -ErrorAction SilentlyContinue
                        $wuReboot = Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired" -ErrorAction SilentlyContinue
                        $pendingFileRename = Get-ItemProperty "HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager" -Name PendingFileRenameOperations -ErrorAction SilentlyContinue
                        
                        return [bool]($cbsReboot -or $wuReboot -or $pendingFileRename)
                    }
                    $computerResult.RebootRequired = $rebootPending
                }
                
                "Install" {
                    if (-not $Force -and -not $PSCmdlet.ShouldProcess($computer, "Install Windows Updates")) {
                        continue
                    }
                    
                    Write-Progress -Activity "Windows Updates: $computer" -Status "Installing updates" -PercentComplete 50
                    
                    if ($hasPSWU) {
                        # Use PSWindowsUpdate module
                        $installResult = Invoke-Command -ComputerName $computer -ScriptBlock {
                            param($Categories, $ExcludeKBs)
                            
                            Import-Module PSWindowsUpdate -Force
                            
                            $criteria = @()
                            if ($Categories -contains "Security") { $criteria += "BrowseOnly=0 and IsInstalled=0 and CategoryIDs contains '0FA1201D-4330-4FA8-8AE9-B877473B6441'" }
                            if ($Categories -contains "Critical") { $criteria += "BrowseOnly=0 and IsInstalled=0 and CategoryIDs contains 'E6CF1350-C01B-414D-A61F-263D14D133B4'" }
                            
                            $updates = Get-WUList -Criteria ($criteria -join " or ") | Where-Object {
                                if ($ExcludeKBs) {
                                    $_.KBArticleIDs | ForEach-Object { $_ -notin $ExcludeKBs }
                                } else { $true }
                            }
                            
                            if ($updates) {
                                $result = Install-WindowsUpdate -KBArticleID $updates.KBArticleIDs -AcceptAll -AutoReboot:$false -Confirm:$false
                                return $result
                            } else {
                                return @{ Result = "No updates to install" }
                            }
                        } -ArgumentList $UpdateCategories, $ExcludeKBs
                        
                        $computerResult.UpdatesInstalled = ($installResult | Measure-Object).Count
                        $computerResult.Status = "Installed"
                    } else {
                        # Use built-in WUA COM object
                        $installResult = Invoke-Command -ComputerName $computer -ScriptBlock {
                            param($Categories, $ExcludeKBs)
                            
                            try {
                                $session = New-Object -ComObject Microsoft.Update.Session
                                $searcher = $session.CreateUpdateSearcher()
                                $searchResult = $searcher.Search("IsInstalled=0 and IsHidden=0")
                                
                                $updatesToInstall = New-Object -ComObject Microsoft.Update.UpdateColl
                                
                                foreach ($update in $searchResult.Updates) {
                                    $shouldInclude = $false
                                    foreach ($category in $update.Categories) {
                                        if ($Categories -contains $category.Name) {
                                            $shouldInclude = $true
                                            break
                                        }
                                    }
                                    
                                    # Check exclusions
                                    if ($shouldInclude -and $ExcludeKBs) {
                                        foreach ($kb in $update.KBArticleIDs) {
                                            if ($kb -in $ExcludeKBs) {
                                                $shouldInclude = $false
                                                break
                                            }
                                        }
                                    }
                                    
                                    if ($shouldInclude) {
                                        $updatesToInstall.Add($update) | Out-Null
                                    }
                                }
                                
                                if ($updatesToInstall.Count -gt 0) {
                                    $downloader = $session.CreateUpdateDownloader()
                                    $downloader.Updates = $updatesToInstall
                                    $downloadResult = $downloader.Download()
                                    
                                    $installer = $session.CreateUpdateInstaller()
                                    $installer.Updates = $updatesToInstall
                                    $installResult = $installer.Install()
                                    
                                    return @{
                                        ResultCode = $installResult.ResultCode
                                        RebootRequired = $installResult.RebootRequired
                                        UpdatesProcessed = $updatesToInstall.Count
                                    }
                                } else {
                                    return @{ Result = "No applicable updates found" }
                                }
                            } catch {
                                throw "Installation failed: $($_.Exception.Message)"
                            }
                        } -ArgumentList $UpdateCategories, $ExcludeKBs
                        
                        $computerResult.UpdatesInstalled = $installResult.UpdatesProcessed
                        $computerResult.RebootRequired = $installResult.RebootRequired
                        $computerResult.Status = "Installed"
                    }
                    
                    # Schedule reboot if requested
                    if ($ScheduleReboot -and $computerResult.RebootRequired) {
                        $rebootTime = (Get-Date).AddHours($ScheduleReboot)
                        Invoke-Command -ComputerName $computer -ScriptBlock {
                            param($RebootTime)
                            schtasks.exe /create /tn "WindowsUpdateReboot" /tr "shutdown.exe /r /f" /sc once /st $RebootTime.ToString("HH:mm") /sd $RebootTime.ToString("MM/dd/yyyy") /f
                        } -ArgumentList $rebootTime
                        
                        Write-Host "  ✓ Reboot scheduled for $rebootTime" -ForegroundColor Green
                    }
                }
                
                "Report" {
                    Write-Progress -Activity "Windows Updates: $computer" -Status "Generating report" -PercentComplete 75
                    
                    # Get update history
                    $updateHistory = Invoke-Command -ComputerName $computer -ScriptBlock {
                        try {
                            $session = New-Object -ComObject Microsoft.Update.Session
                            $searcher = $session.CreateUpdateSearcher()
                            $historyCount = $searcher.GetTotalHistoryCount()
                            
                            if ($historyCount -gt 0) {
                                $history = $searcher.QueryHistory(0, [Math]::Min($historyCount, 50))
                                $recentUpdates = @()
                                
                                foreach ($entry in $history) {
                                    $recentUpdates += [PSCustomObject]@{
                                        Title = $entry.Title
                                        Date = $entry.Date
                                        Operation = switch ($entry.Operation) {
                                            0 { "Not Started" }
                                            1 { "Installation" }
                                            2 { "Uninstallation" }
                                            3 { "Other" }
                                        }
                                        Result = switch ($entry.ResultCode) {
                                            0 { "Not Started" }
                                            1 { "In Progress" }
                                            2 { "Succeeded" }
                                            3 { "Succeeded With Errors" }
                                            4 { "Failed" }
                                            5 { "Aborted" }
                                        }
                                    }
                                }
                                
                                return $recentUpdates | Sort-Object Date -Descending
                            }
                        } catch {
                            return @()
                        }
                    }
                    
                    $computerResult.UpdateDetails = $updateHistory
                    $computerResult.LastUpdateDate = if ($updateHistory) { ($updateHistory | Where-Object Result -eq "Succeeded" | Select-Object -First 1).Date } else { $null }
                    $computerResult.Status = "Reported"
                }
                
                "Reboot" {
                    if (-not $Force -and -not $PSCmdlet.ShouldProcess($computer, "Restart Computer")) {
                        continue
                    }
                    
                    Write-Progress -Activity "Windows Updates: $computer" -Status "Restarting computer" -PercentComplete 90
                    Restart-Computer -ComputerName $computer -Force -ErrorAction Stop
                    $computerResult.Status = "Rebooted"
                }
                
                "Configure" {
                    Write-Progress -Activity "Windows Updates: $computer" -Status "Configuring update settings" -PercentComplete 60
                    
                    Invoke-Command -ComputerName $computer -ScriptBlock {
                        # Configure automatic updates
                        $auKey = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update"
                        Set-ItemProperty -Path $auKey -Name "AUOptions" -Value 4 -Force # Download and install automatically
                        Set-ItemProperty -Path $auKey -Name "ScheduledInstallDay" -Value 0 -Force # Every day
                        Set-ItemProperty -Path $auKey -Name "ScheduledInstallTime" -Value 3 -Force # 3 AM
                        
                        # Restart Windows Update service
                        Restart-Service -Name wuauserv -Force
                    }
                    
                    $computerResult.Status = "Configured"
                }
            }
            
            $results += $computerResult
            Write-Host "  ✓ Completed $Action for $computer" -ForegroundColor Green
            
        } catch {
            Write-Error "Failed to process $computer : $_"
            $results += [PSCustomObject]@{
                ComputerName = $computer
                Action = $Action
                Status = "Error"
                Error = $_.Exception.Message
                Timestamp = Get-Date
            }
        }
        
        Write-Progress -Activity "Windows Updates: $computer" -Completed
    }
}

end {
    # Export results if path specified
    if ($ExportPath) {
        $results | Export-Csv -Path "$ExportPath\WindowsUpdates_$Action_$timestamp.csv" -NoTypeInformation
        $results | Export-Clixml -Path "$ExportPath\WindowsUpdates_$Action_$timestamp.xml"
        Write-Host "Results exported to $ExportPath" -ForegroundColor Green
    }
    
    # Display summary
    Write-Host "`n" + "="*80 -ForegroundColor Cyan
    Write-Host "WINDOWS UPDATE MANAGEMENT SUMMARY - $Action" -ForegroundColor Cyan
    Write-Host "="*80 -ForegroundColor Cyan
    
    switch ($Action) {
        "Check" {
            $summary = $results | Select-Object ComputerName, UpdatesAvailable, RebootRequired, Status | Format-Table -AutoSize
            $summary
            
            $totalUpdates = ($results | Measure-Object UpdatesAvailable -Sum).Sum
            $needReboot = ($results | Where-Object RebootRequired -eq $true).Count
            
            Write-Host "Total updates available: $totalUpdates" -ForegroundColor Yellow
            Write-Host "Systems requiring reboot: $needReboot" -ForegroundColor $(if ($needReboot -gt 0) {"Red"} else {"Green"})
        }
        
        "Install" {
            $summary = $results | Select-Object ComputerName, UpdatesInstalled, RebootRequired, Status | Format-Table -AutoSize
            $summary
            
            $totalInstalled = ($results | Measure-Object UpdatesInstalled -Sum).Sum
            Write-Host "Total updates installed: $totalInstalled" -ForegroundColor Green
        }
        
        "Report" {
            $summary = $results | Select-Object ComputerName, LastUpdateDate, @{N="RecentUpdates";E={$_.UpdateDetails.Count}}, Status | Format-Table -AutoSize
            $summary
        }
        
        default {
            $results | Select-Object ComputerName, Status | Format-Table -AutoSize
        }
    }
    
    $successCount = ($results | Where-Object Status -notlike "*Error*").Count
    $errorCount = ($results | Where-Object Status -like "*Error*").Count
    
    Write-Host "`nProcessed $($results.Count) systems: $successCount successful, $errorCount errors" -ForegroundColor $(if ($errorCount -eq 0) {"Green"} else {"Yellow"})
}