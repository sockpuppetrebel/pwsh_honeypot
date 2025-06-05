#Requires -Modules ExchangeOnlineManagement

<#
.SYNOPSIS
    Sets up mail forwarding for multiple mailboxes to a specified recipient
    
.DESCRIPTION
    This script configures mail forwarding for a list of mailboxes to a target recipient.
    It validates each mailbox exists, sets forwarding, and provides a summary report.
    
.PARAMETER TargetRecipient
    The email address where all mail should be forwarded to
    
.PARAMETER SourceMailboxes
    Array of source mailbox email addresses to configure forwarding for
    
.PARAMETER DeliverToMailboxAndForward
    If specified, mail will be delivered to both the original mailbox AND forwarded.
    If not specified (default), mail will ONLY be forwarded (not kept in original mailbox).
    
.EXAMPLE
    .\Set-MailboxForwarding-Bulk.ps1 -TargetRecipient "first.last@optimizely.com" -SourceMailboxes @("source1@company.com", "source2@company.com")
    
.EXAMPLE
    .\Set-MailboxForwarding-Bulk.ps1 -TargetRecipient "first.last@optimizely.com" -SourceMailboxes @("source@company.com") -DeliverToMailboxAndForward
    
.NOTES
    Author: Jason Slater
    Version: 1.0
    Date: 2025-06-05
    Requires: Exchange Online Management with mailbox admin permissions
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory = $true)]
    [ValidatePattern('^[^@]+@[^@]+\.[^@]+$')]
    [string]$TargetRecipient,
    
    [Parameter(Mandatory = $true)]
    [string[]]$SourceMailboxes,
    
    [Parameter(Mandatory = $false)]
    [switch]$DeliverToMailboxAndForward
)

# Function to write colored output
function Write-ColorOutput {
    param(
        [string]$Message,
        [string]$ForegroundColor = "White"
    )
    Write-Host $Message -ForegroundColor $ForegroundColor
}

# Connect to Exchange Online if not already connected
$exoSession = Get-PSSession | Where-Object {$_.ConfigurationName -eq "Microsoft.Exchange" -and $_.State -eq "Opened"}
if (-not $exoSession) {
    Write-ColorOutput "Connecting to Exchange Online..." -ForegroundColor Yellow
    try {
        Connect-ExchangeOnline -ShowBanner:$false
        Write-ColorOutput "✓ Connected to Exchange Online" -ForegroundColor Green
    }
    catch {
        Write-Error "Failed to connect to Exchange Online: $_"
        exit 1
    }
}

Write-ColorOutput "`n=== BULK MAILBOX FORWARDING SETUP ===" -ForegroundColor Cyan
Write-ColorOutput "Target recipient: $TargetRecipient" -ForegroundColor Yellow
Write-ColorOutput "Number of source mailboxes: $($SourceMailboxes.Count)" -ForegroundColor Yellow
Write-ColorOutput "Delivery mode: $(if ($DeliverToMailboxAndForward) { 'Forward AND keep in original mailbox' } else { 'Forward ONLY (original mailbox will be empty)' })" -ForegroundColor Yellow

# Validate target recipient exists
Write-ColorOutput "`nValidating target recipient..." -ForegroundColor Yellow
try {
    $targetUser = Get-Mailbox -Identity $TargetRecipient -ErrorAction Stop
    Write-ColorOutput "✓ Target recipient found: $($targetUser.DisplayName)" -ForegroundColor Green
}
catch {
    Write-Error "Target recipient not found: $TargetRecipient"
    exit 1
}

# Validate source mailboxes and check current forwarding status
Write-ColorOutput "`nValidating source mailboxes..." -ForegroundColor Yellow
$validMailboxes = @()
$invalidMailboxes = @()
$alreadyForwarding = @()
$correctlyConfigured = @()

foreach ($mailbox in $SourceMailboxes) {
    try {
        $mbx = Get-Mailbox -Identity $mailbox -ErrorAction Stop
        
        # Check current forwarding status
        if ($mbx.ForwardingAddress -or $mbx.ForwardingSmtpAddress) {
            $currentForwarding = if ($mbx.ForwardingAddress) { $mbx.ForwardingAddress } else { $mbx.ForwardingSmtpAddress }
            
            # Check if already forwarding to our target
            if ($currentForwarding -eq $TargetRecipient) {
                $correctlyConfigured += [PSCustomObject]@{
                    Mailbox = $mailbox
                    CurrentForwarding = $currentForwarding
                    DeliverToMailboxAndForward = $mbx.DeliverToMailboxAndForward
                    MailboxObject = $mbx
                }
                Write-ColorOutput "  ✓ $mailbox (already forwarding to target: $currentForwarding)" -ForegroundColor Green
            } else {
                $alreadyForwarding += [PSCustomObject]@{
                    Mailbox = $mailbox
                    CurrentForwarding = $currentForwarding
                    DeliverToMailboxAndForward = $mbx.DeliverToMailboxAndForward
                    MailboxObject = $mbx
                }
                Write-ColorOutput "  ⚠ $mailbox (forwarding to different address: $currentForwarding)" -ForegroundColor Yellow
            }
        } else {
            $validMailboxes += $mbx
            Write-ColorOutput "  ✓ $mailbox (no current forwarding)" -ForegroundColor Green
        }
    }
    catch {
        $invalidMailboxes += $mailbox
        Write-ColorOutput "  ✗ $mailbox (not found)" -ForegroundColor Red
    }
}

# Summary of validation
Write-ColorOutput "`n=== VALIDATION SUMMARY ===" -ForegroundColor Cyan
Write-ColorOutput "Ready to configure: $($validMailboxes.Count)" -ForegroundColor Green
Write-ColorOutput "Already correctly configured: $($correctlyConfigured.Count)" -ForegroundColor Green
Write-ColorOutput "Forwarding conflicts (need manual review): $($alreadyForwarding.Count)" -ForegroundColor Yellow
Write-ColorOutput "Invalid mailboxes: $($invalidMailboxes.Count)" -ForegroundColor Red

if ($invalidMailboxes.Count -gt 0) {
    Write-ColorOutput "`n=== INVALID MAILBOXES ===" -ForegroundColor Red
    Write-ColorOutput "These mailboxes were not found:" -ForegroundColor Red
    $invalidMailboxes | ForEach-Object { Write-ColorOutput "  ✗ $_" -ForegroundColor Red }
}

if ($correctlyConfigured.Count -gt 0) {
    Write-ColorOutput "`n=== ALREADY CORRECTLY CONFIGURED ===" -ForegroundColor Green
    Write-ColorOutput "These mailboxes are already forwarding to ${TargetRecipient}:" -ForegroundColor Green
    $correctlyConfigured | ForEach-Object { 
        $deliveryMode = if ($_.DeliverToMailboxAndForward) { "Forward AND Keep" } else { "Forward ONLY" }
        Write-ColorOutput "  ✓ $($_.Mailbox) → $($_.CurrentForwarding) ($deliveryMode)" -ForegroundColor Green 
    }
}

if ($alreadyForwarding.Count -gt 0) {
    Write-ColorOutput "`n=== FORWARDING CONFLICTS ===" -ForegroundColor Yellow
    Write-ColorOutput "These mailboxes are forwarding to different addresses:" -ForegroundColor Yellow
    Write-ColorOutput "(These will be SKIPPED to preserve existing settings)" -ForegroundColor Red
    $alreadyForwarding | ForEach-Object { 
        $deliveryMode = if ($_.DeliverToMailboxAndForward) { "Forward AND Keep" } else { "Forward ONLY" }
        Write-ColorOutput "  ⚠ $($_.Mailbox) → $($_.CurrentForwarding) ($deliveryMode)" -ForegroundColor Yellow 
    }
}

# Only configure mailboxes without forwarding conflicts
$mailboxesToConfigure = $validMailboxes

if ($alreadyForwarding.Count -gt 0) {
    Write-ColorOutput "`n⚠ WARNING: Mailboxes with forwarding conflicts will be SKIPPED" -ForegroundColor Yellow
    Write-ColorOutput "This script will NOT modify existing forwarding to preserve current settings." -ForegroundColor Yellow
    Write-ColorOutput "If you need to change these, please review them manually." -ForegroundColor Yellow
}

if ($mailboxesToConfigure.Count -eq 0) {
    Write-ColorOutput "`nNo mailboxes to configure. Exiting." -ForegroundColor Yellow
    exit 0
}

# Final confirmation
Write-ColorOutput "`n=== CONFIGURATION PLAN ===" -ForegroundColor Cyan
if ($mailboxesToConfigure.Count -gt 0) {
    Write-ColorOutput "Will configure forwarding for $($mailboxesToConfigure.Count) mailboxes:" -ForegroundColor White
    $mailboxesToConfigure | ForEach-Object { Write-ColorOutput "  → $($_.PrimarySmtpAddress)" -ForegroundColor White }
    Write-ColorOutput "`nTarget: $TargetRecipient" -ForegroundColor White
    Write-ColorOutput "Mode: $(if ($DeliverToMailboxAndForward) { 'Forward AND keep' } else { 'Forward ONLY' })" -ForegroundColor White
} else {
    Write-ColorOutput "No mailboxes need configuration." -ForegroundColor Yellow
}

if (-not $WhatIfPreference) {
    Write-ColorOutput "`nProceed with configuration? (Y/N): " -ForegroundColor Yellow -NoNewline
    $confirm = Read-Host
    if ($confirm -ne 'Y' -and $confirm -ne 'y') {
        Write-ColorOutput "Configuration cancelled." -ForegroundColor Yellow
        exit 0
    }
}

# Configure forwarding
Write-ColorOutput "`nConfiguring mail forwarding..." -ForegroundColor Yellow
$successCount = 0
$errorCount = 0
$errors = @()

foreach ($mailbox in $mailboxesToConfigure) {
    try {
        if ($PSCmdlet.ShouldProcess($mailbox.PrimarySmtpAddress, "Set mail forwarding to $TargetRecipient")) {
            Set-Mailbox -Identity $mailbox.Identity -ForwardingSmtpAddress $TargetRecipient -DeliverToMailboxAndForward:$DeliverToMailboxAndForward -ErrorAction Stop
            $successCount++
            Write-ColorOutput "  ✓ $($mailbox.PrimarySmtpAddress)" -ForegroundColor Green
        }
    }
    catch {
        $errorCount++
        $errorMessage = "Failed to configure $($mailbox.PrimarySmtpAddress): $_"
        $errors += $errorMessage
        Write-ColorOutput "  ✗ $($mailbox.PrimarySmtpAddress) - $_" -ForegroundColor Red
    }
}

# Final summary
Write-ColorOutput "`n=== CONFIGURATION COMPLETE ===" -ForegroundColor Cyan
Write-ColorOutput "Successfully configured: $successCount mailboxes" -ForegroundColor Green
Write-ColorOutput "Already correctly configured: $($correctlyConfigured.Count) mailboxes" -ForegroundColor Green
Write-ColorOutput "Skipped (forwarding conflicts): $($alreadyForwarding.Count) mailboxes" -ForegroundColor Yellow
Write-ColorOutput "Invalid mailboxes: $($invalidMailboxes.Count) mailboxes" -ForegroundColor Red

if ($errorCount -gt 0) {
    Write-ColorOutput "Failed to configure: $errorCount mailboxes" -ForegroundColor Red
    Write-ColorOutput "`nErrors encountered:" -ForegroundColor Red
    $errors | ForEach-Object { Write-ColorOutput "  $_" -ForegroundColor Red }
}

if ($alreadyForwarding.Count -gt 0) {
    Write-ColorOutput "`n=== MANUAL REVIEW REQUIRED ===" -ForegroundColor Yellow
    Write-ColorOutput "The following mailboxes have existing forwarding that was NOT modified:" -ForegroundColor Yellow
    $alreadyForwarding | ForEach-Object { Write-ColorOutput "  $($_.Mailbox) → $($_.CurrentForwarding)" -ForegroundColor Yellow }
    Write-ColorOutput "Please review these manually if changes are needed." -ForegroundColor Yellow
}

# Show verification command
Write-ColorOutput "`nTo verify the configuration, run:" -ForegroundColor Cyan
Write-ColorOutput "Get-Mailbox -Identity '<mailbox>' | Select-Object DisplayName,ForwardingSmtpAddress,DeliverToMailboxAndForward" -ForegroundColor Gray

if ($successCount -gt 0) {
    Write-ColorOutput "`n✓ Mail forwarding has been configured successfully!" -ForegroundColor Green
    Write-ColorOutput "All mail to the configured mailboxes will now be $(if ($DeliverToMailboxAndForward) { 'delivered to both the original mailbox AND forwarded to' } else { 'forwarded to' }) $TargetRecipient" -ForegroundColor White
}