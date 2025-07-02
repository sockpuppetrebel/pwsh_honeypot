<#
.SYNOPSIS
    Converts a distribution list to a shared mailbox and grants full access to all current members.

.DESCRIPTION
    This script performs the following actions:
    1. Gets all current members of the distribution list
    2. Removes the distribution list
    3. Creates a new shared mailbox with the same email address
    4. Grants full access permissions to all former DL members
    5. Logs all actions for audit purposes

.PARAMETER DistributionListEmail
    The email address of the distribution list to convert (e.g., CSOutreach@optimizely.com)

.PARAMETER DisplayName
    The display name for the new shared mailbox (optional, defaults to DL display name)

.EXAMPLE
    .\Convert-DLToSharedMailbox.ps1 -DistributionListEmail "CSOutreach@optimizely.com"

.NOTES
    Author: Exchange Admin Script
    Date: 2025-07-02
    Version: 1.0
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)]
    [string]$DistributionListEmail,
    
    [Parameter(Mandatory=$false)]
    [string]$DisplayName = ""
)

# Start transcript for logging
$LogPath = "$PSScriptRoot\Logs"
if (-not (Test-Path $LogPath)) {
    New-Item -ItemType Directory -Path $LogPath -Force | Out-Null
}
$LogFile = "$LogPath\Convert-DL-$(Get-Date -Format 'yyyyMMdd-HHmmss').log"
Start-Transcript -Path $LogFile

Write-Host "=== Starting DL to Shared Mailbox Conversion ===" -ForegroundColor Green
Write-Host "Distribution List: $DistributionListEmail" -ForegroundColor Yellow
Write-Host "Timestamp: $(Get-Date)" -ForegroundColor Yellow

try {
    # Connect to Exchange Online (if not already connected)
    $connectionTest = Get-PSSession | Where-Object {$_.ConfigurationName -eq "Microsoft.Exchange" -and $_.State -eq "Opened"}
    if (-not $connectionTest) {
        Write-Host "Connecting to Exchange Online..." -ForegroundColor Yellow
        Connect-ExchangeOnline -ShowBanner:$false
    }
    
    # Step 1: Get the distribution list details
    Write-Host "`nStep 1: Getting distribution list information..." -ForegroundColor Cyan
    $DL = Get-DistributionGroup -Identity $DistributionListEmail -ErrorAction Stop
    
    if (-not $DL) {
        throw "Distribution list '$DistributionListEmail' not found!"
    }
    
    # Use DL display name if no custom display name provided
    if ([string]::IsNullOrWhiteSpace($DisplayName)) {
        $DisplayName = $DL.DisplayName
    }
    
    Write-Host "Found DL: $($DL.DisplayName) ($($DL.PrimarySmtpAddress))" -ForegroundColor Green
    
    # Step 2: Get all members of the distribution list
    Write-Host "`nStep 2: Getting distribution list members..." -ForegroundColor Cyan
    $DLMembers = Get-DistributionGroupMember -Identity $DistributionListEmail -ResultSize Unlimited
    
    Write-Host "Found $($DLMembers.Count) members:" -ForegroundColor Green
    $DLMembers | ForEach-Object {
        Write-Host "  - $($_.DisplayName) ($($_.PrimarySmtpAddress))" -ForegroundColor Gray
    }
    
    # Step 3: Export DL settings for reference
    Write-Host "`nStep 3: Exporting DL settings for backup..." -ForegroundColor Cyan
    $BackupFile = "$LogPath\DL-Backup-$($DL.Alias)-$(Get-Date -Format 'yyyyMMdd-HHmmss').json"
    $DLBackup = @{
        DisplayName = $DL.DisplayName
        Alias = $DL.Alias
        PrimarySmtpAddress = $DL.PrimarySmtpAddress
        EmailAddresses = $DL.EmailAddresses
        Members = $DLMembers | Select-Object DisplayName, PrimarySmtpAddress, RecipientType
        ManagedBy = $DL.ManagedBy
        AcceptMessagesOnlyFrom = $DL.AcceptMessagesOnlyFrom
        AcceptMessagesOnlyFromDLMembers = $DL.AcceptMessagesOnlyFromDLMembers
        RequireSenderAuthenticationEnabled = $DL.RequireSenderAuthenticationEnabled
    }
    $DLBackup | ConvertTo-Json -Depth 10 | Out-File $BackupFile
    Write-Host "DL backup saved to: $BackupFile" -ForegroundColor Green
    
    # Step 4: Remove the distribution list
    Write-Host "`nStep 4: Removing distribution list..." -ForegroundColor Cyan
    Write-Host "Are you sure you want to remove the DL '$($DL.DisplayName)'? (Y/N): " -NoNewline -ForegroundColor Yellow
    $confirmation = Read-Host
    
    if ($confirmation -eq 'Y' -or $confirmation -eq 'y') {
        Remove-DistributionGroup -Identity $DistributionListEmail -Confirm:$false
        Write-Host "Distribution list removed successfully." -ForegroundColor Green
        
        # Wait a moment for replication
        Write-Host "Waiting 30 seconds for AD replication..." -ForegroundColor Yellow
        Start-Sleep -Seconds 30
    }
    else {
        Write-Host "Operation cancelled by user." -ForegroundColor Red
        Stop-Transcript
        return
    }
    
    # Step 5: Create the shared mailbox
    Write-Host "`nStep 5: Creating shared mailbox..." -ForegroundColor Cyan
    $SharedMailbox = New-Mailbox -Shared -Name $DL.Name -DisplayName $DisplayName -Alias $DL.Alias -PrimarySmtpAddress $DL.PrimarySmtpAddress
    
    if ($SharedMailbox) {
        Write-Host "Shared mailbox created successfully:" -ForegroundColor Green
        Write-Host "  Display Name: $($SharedMailbox.DisplayName)" -ForegroundColor Gray
        Write-Host "  Email: $($SharedMailbox.PrimarySmtpAddress)" -ForegroundColor Gray
    }
    
    # Wait for mailbox to be fully provisioned
    Write-Host "Waiting 60 seconds for mailbox provisioning..." -ForegroundColor Yellow
    Start-Sleep -Seconds 60
    
    # Step 6: Grant full access to all former DL members
    Write-Host "`nStep 6: Granting full access permissions to former DL members..." -ForegroundColor Cyan
    $successCount = 0
    $failCount = 0
    
    foreach ($member in $DLMembers) {
        try {
            # Only grant access to user mailboxes (not other groups or contacts)
            if ($member.RecipientType -like "*Mailbox*") {
                Add-MailboxPermission -Identity $SharedMailbox.PrimarySmtpAddress -User $member.PrimarySmtpAddress -AccessRights FullAccess -InheritanceType All -AutoMapping $true -ErrorAction Stop
                Add-RecipientPermission -Identity $SharedMailbox.PrimarySmtpAddress -Trustee $member.PrimarySmtpAddress -AccessRights SendAs -Confirm:$false -ErrorAction Stop
                Write-Host "  ✓ Granted access to: $($member.DisplayName)" -ForegroundColor Green
                $successCount++
            }
            else {
                Write-Host "  ! Skipped non-mailbox member: $($member.DisplayName) (Type: $($member.RecipientType))" -ForegroundColor Yellow
            }
        }
        catch {
            Write-Host "  ✗ Failed to grant access to: $($member.DisplayName) - Error: $_" -ForegroundColor Red
            $failCount++
        }
    }
    
    # Step 7: Copy any additional email addresses
    Write-Host "`nStep 7: Adding additional email addresses..." -ForegroundColor Cyan
    $additionalAddresses = $DL.EmailAddresses | Where-Object { $_ -like "smtp:*" -and $_ -ne "SMTP:$($DL.PrimarySmtpAddress)" }
    
    if ($additionalAddresses.Count -gt 0) {
        Set-Mailbox -Identity $SharedMailbox.PrimarySmtpAddress -EmailAddresses @{Add=$additionalAddresses}
        Write-Host "Added $($additionalAddresses.Count) additional email addresses." -ForegroundColor Green
    }
    
    # Summary
    Write-Host "`n=== Conversion Summary ===" -ForegroundColor Green
    Write-Host "Original DL: $($DL.DisplayName) ($($DL.PrimarySmtpAddress))" -ForegroundColor Yellow
    Write-Host "New Shared Mailbox: $($SharedMailbox.DisplayName) ($($SharedMailbox.PrimarySmtpAddress))" -ForegroundColor Yellow
    Write-Host "Permissions granted: $successCount successful, $failCount failed" -ForegroundColor Yellow
    Write-Host "Log file: $LogFile" -ForegroundColor Yellow
    Write-Host "Backup file: $BackupFile" -ForegroundColor Yellow
    
    Write-Host "`nConversion completed successfully!" -ForegroundColor Green
    
}
catch {
    Write-Host "`nERROR: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Stack Trace: $($_.ScriptStackTrace)" -ForegroundColor Red
    
    # Attempt to restore if something went wrong
    if ($DL -and -not (Get-Mailbox -Identity $DistributionListEmail -ErrorAction SilentlyContinue)) {
        Write-Host "`nAttempting to restore distribution list..." -ForegroundColor Yellow
        # This would need more complex restore logic based on the backup
        Write-Host "Please review the backup file at: $BackupFile" -ForegroundColor Yellow
    }
}
finally {
    Stop-Transcript
    
    # Disconnect from Exchange Online if we connected in this script
    if (-not $connectionTest) {
        Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
    }
}