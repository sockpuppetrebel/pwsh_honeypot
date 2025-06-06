#Requires -Modules Az.Storage

<#
.SYNOPSIS
    Quick sync script for cloud-resume learning timeline
    
.DESCRIPTION
    Wrapper script that syncs the learning timeline from cloud-resume to Azure Blob Storage
    with pre-configured paths and settings.
    
.PARAMETER UpdateMainIndex
    If specified, also updates the main index.json file with learning timeline reference
    
.EXAMPLE
    .\Update-CloudResumeLearning.ps1
    
.EXAMPLE
    .\Update-CloudResumeLearning.ps1 -UpdateMainIndex
    
.NOTES
    Author: Jason Slater
    Version: 1.0
    Date: 2025-06-05
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [switch]$UpdateMainIndex
)

# Configuration
$config = @{
    HtmlPath = "$HOME/Projects/cloud-resume/site/index.html"
    ScriptPath = Join-Path $PSScriptRoot "Sync-LearningTimeline-ToBlob.ps1"
    StorageAccount = "jslaterdevjournal"
    Container = "claude-development-journal"
}

Write-Host "=== CLOUD RESUME LEARNING SYNC ===" -ForegroundColor Cyan
Write-Host "This will sync your learning timeline to Azure Blob Storage" -ForegroundColor Yellow

# Verify paths exist
if (-not (Test-Path $config.HtmlPath)) {
    Write-Error "Cloud resume index.html not found at: $($config.HtmlPath)"
    exit 1
}

if (-not (Test-Path $config.ScriptPath)) {
    Write-Error "Sync script not found at: $($config.ScriptPath)"
    exit 1
}

# Show what will be synced
Write-Host "`nSource: $($config.HtmlPath)" -ForegroundColor Gray
Write-Host "Target: https://$($config.StorageAccount).blob.core.windows.net/$($config.Container)/learning-timeline/timeline.json" -ForegroundColor Gray

# Run the sync
$params = @{
    HtmlPath = $config.HtmlPath
    StorageAccountName = $config.StorageAccount
    ContainerName = $config.Container
}

if ($UpdateMainIndex) {
    $params.UpdateIndex = $true
    Write-Host "`nWill also update main index.json" -ForegroundColor Yellow
}

Write-Host "`nStarting sync..." -ForegroundColor Yellow
& $config.ScriptPath @params

# Show quick access links
Write-Host "`n=== QUICK ACCESS LINKS ===" -ForegroundColor Cyan
Write-Host "Timeline JSON: " -NoNewline
Write-Host "https://$($config.StorageAccount).blob.core.windows.net/$($config.Container)/learning-timeline/timeline.json" -ForegroundColor Cyan
Write-Host "Main Index: " -NoNewline  
Write-Host "https://$($config.StorageAccount).blob.core.windows.net/$($config.Container)/index.json" -ForegroundColor Cyan

Write-Host "`nTip: Add this to your regular deployment process!" -ForegroundColor Gray