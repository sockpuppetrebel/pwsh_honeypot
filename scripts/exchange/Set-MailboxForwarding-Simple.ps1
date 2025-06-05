#Requires -Modules ExchangeOnlineManagement

<#
.SYNOPSIS
    Simple mailbox forwarding setup with all authentication upfront
    
.DESCRIPTION
    Sets up mail forwarding for specified mailboxes without interactive prompts.
    All authentication happens immediately at script start.
    
.PARAMETER TargetRecipient
    The email address where all mail should be forwarded to
    
.PARAMETER SourceMailboxes
    Array of source mailbox email addresses to configure forwarding for
    
.PARAMETER DeliverToMailboxAndForward
    If specified, mail will be delivered to both the original mailbox AND forwarded.
    If not specified (default), mail will ONLY be forwarded.
    
.EXAMPLE
    $mailboxes = @("source1@company.com", "source2@company.com")
    .\Set-MailboxForwarding-Simple.ps1 -TargetRecipient "first.last@optimizely.com" -SourceMailboxes $mailboxes
    
.NOTES
    Author: Jason Slater
    Version: 1.0
    Date: 2025-06-05
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

# AUTHENTICATE IMMEDIATELY - before any user interaction
Write-ColorOutput "=== AUTHENTICATING TO EXCHANGE ONLINE ===" -ForegroundColor Cyan
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
} else {
    Write-ColorOutput "✓ Already connected to Exchange Online" -ForegroundColor Green
}

Write-ColorOutput "`n=== MAILBOX FORWARDING CONFIGURATION ===" -ForegroundColor Cyan
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

# Process each mailbox
Write-ColorOutput "`nProcessing mailboxes..." -ForegroundColor Yellow
$results = @()

foreach ($mailbox in $SourceMailboxes) {
    $result = [PSCustomObject]@{
        Mailbox = $mailbox
        Status = "Unknown"
        Action = "None"
        Details = ""
    }
    
    try {
        # Check if mailbox exists
        $mbx = Get-Mailbox -Identity $mailbox -ErrorAction Stop
        
        # Check current forwarding status
        if ($mbx.ForwardingAddress -or $mbx.ForwardingSmtpAddress) {
            $currentForwarding = if ($mbx.ForwardingAddress) { $mbx.ForwardingAddress } else { $mbx.ForwardingSmtpAddress }
            
            if ($currentForwarding -eq $TargetRecipient) {
                $result.Status = "AlreadyConfigured"
                $result.Action = "Skipped"
                $result.Details = "Already forwarding to target"
                Write-ColorOutput "  ✓ $mailbox (already configured)" -ForegroundColor Green
            } else {
                $result.Status = "Conflict"
                $result.Action = "Skipped"
                $result.Details = "Forwarding to: $currentForwarding"
                Write-ColorOutput "  ⚠ $mailbox (forwarding conflict - skipped)" -ForegroundColor Yellow
            }
        } else {
            # Configure forwarding
            try {
                if ($PSCmdlet.ShouldProcess($mailbox, "Set mail forwarding to $TargetRecipient")) {
                    Set-Mailbox -Identity $mbx.Identity -ForwardingSmtpAddress $TargetRecipient -DeliverToMailboxAndForward:$DeliverToMailboxAndForward -ErrorAction Stop
                    $result.Status = "Success"
                    $result.Action = "Configured"
                    $result.Details = "Forwarding configured successfully"
                    Write-ColorOutput "  ✓ $mailbox (configured)" -ForegroundColor Green
                }
            }
            catch {
                $result.Status = "Error"
                $result.Action = "Failed"
                $result.Details = $_.Exception.Message
                Write-ColorOutput "  ✗ $mailbox (failed: $_)" -ForegroundColor Red
            }
        }
    }
    catch {
        $result.Status = "NotFound"
        $result.Action = "Skipped"
        $result.Details = "Mailbox not found"
        Write-ColorOutput "  ✗ $mailbox (not found)" -ForegroundColor Red
    }
    
    $results += $result
}

# Summary report
Write-ColorOutput "`n=== SUMMARY REPORT ===" -ForegroundColor Cyan
$configured = $results | Where-Object { $_.Status -eq "Success" }
$alreadyConfigured = $results | Where-Object { $_.Status -eq "AlreadyConfigured" }
$conflicts = $results | Where-Object { $_.Status -eq "Conflict" }
$notFound = $results | Where-Object { $_.Status -eq "NotFound" }
$errors = $results | Where-Object { $_.Status -eq "Error" }

Write-ColorOutput "Newly configured: $($configured.Count)" -ForegroundColor Green
Write-ColorOutput "Already configured: $($alreadyConfigured.Count)" -ForegroundColor Green
Write-ColorOutput "Forwarding conflicts: $($conflicts.Count)" -ForegroundColor Yellow
Write-ColorOutput "Not found: $($notFound.Count)" -ForegroundColor Red
Write-ColorOutput "Errors: $($errors.Count)" -ForegroundColor Red

if ($conflicts.Count -gt 0) {
    Write-ColorOutput "`nForwarding conflicts (manual review needed):" -ForegroundColor Yellow
    $conflicts | ForEach-Object { Write-ColorOutput "  $($_.Mailbox) → $($_.Details -replace 'Forwarding to: ', '')" -ForegroundColor Yellow }
}

if ($errors.Count -gt 0) {
    Write-ColorOutput "`nErrors encountered:" -ForegroundColor Red
    $errors | ForEach-Object { Write-ColorOutput "  $($_.Mailbox): $($_.Details)" -ForegroundColor Red }
}

if ($notFound.Count -gt 0) {
    Write-ColorOutput "`nMailboxes not found:" -ForegroundColor Red
    $notFound | ForEach-Object { Write-ColorOutput "  $($_.Mailbox)" -ForegroundColor Red }
}

Write-ColorOutput "`n✓ Forwarding configuration complete!" -ForegroundColor Green