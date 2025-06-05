#Requires -Modules ExchangeOnlineManagement

<#
.SYNOPSIS
    Example script using the simple mailbox forwarding tool
    
.DESCRIPTION
    Demonstrates how to use Set-MailboxForwarding-Simple.ps1 for legacy mailbox forwarding.
    All authentication happens upfront to avoid interrupting workflows.
    
.NOTES
    Author: Jason Slater
    Version: 1.0
    Date: 2025-06-05
#>

# Use the simple forwarding script (no interactive prompts)
$scriptPath = Join-Path $PSScriptRoot "Set-MailboxForwarding-Simple.ps1"

# Example configuration
$targetRecipient = "first.last@optimizely.com"
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

# Run the simple forwarding script
& $scriptPath -TargetRecipient $targetRecipient -SourceMailboxes $sourceMailboxes