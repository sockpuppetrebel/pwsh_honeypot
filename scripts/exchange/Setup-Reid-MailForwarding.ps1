#Requires -Modules ExchangeOnlineManagement

<#
.SYNOPSIS
    Sets up mail forwarding for specific legacy mailboxes to first.last@optimizely.com
    
.DESCRIPTION
    Configures mail forwarding for the specified list of legacy AWS and Ektron mailboxes
    to first.last@optimizely.com. Preserves any existing forwarding settings.
    
.NOTES
    Author: Jason Slater
    Version: 1.0
    Date: 2025-06-05
#>

# Import the bulk forwarding script
$scriptPath = Join-Path $PSScriptRoot "Set-MailboxForwarding-Bulk.ps1"

# Target recipient
$targetRecipient = "first.last@optimizely.com"

# Source mailboxes to configure
$sourceMailboxes = @(
    "cit-aws@ektron.com",
    "ektronmanagedservices-aws@ektron.com",
    "ektaz1@ektron.com",
    "EktronSupport-AWS@ektron.com",
    "aws.idiotest@episerver.net",
    "managedservices_rnd@ektron.com",
    "aws.welcome-gehc@episerver.net",
    "aws.welcome-trayio@episerver.net",
    "aws.welcome-trayio-hubspot@episerver.net",
    "aws.welcome-webproofing2@episerver.net"
)

Write-Host "=== SETTING UP MAIL FORWARDING FOR USER ===" -ForegroundColor Cyan
Write-Host "Target recipient: $targetRecipient" -ForegroundColor Yellow
Write-Host "Number of mailboxes: $($sourceMailboxes.Count)" -ForegroundColor Yellow
Write-Host ""

# Run the bulk forwarding script
& $scriptPath -TargetRecipient $targetRecipient -SourceMailboxes $sourceMailboxes