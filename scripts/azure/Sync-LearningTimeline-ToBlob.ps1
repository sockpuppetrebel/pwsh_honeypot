#Requires -Modules Az.Storage

<#
.SYNOPSIS
    Syncs learning timeline items from slater.cloud to Azure Blob Storage
    
.DESCRIPTION
    This script extracts learning items from the cloud-resume site's index.html
    and syncs them to Azure Blob Storage as structured JSON data. It maintains
    the existing index.json structure while adding learning timeline data.
    
.PARAMETER HtmlPath
    Path to the index.html file containing learning items
    
.PARAMETER StorageAccountName
    Azure Storage Account name (default: jslaterdevjournal)
    
.PARAMETER ContainerName
    Container name in Azure Blob Storage (default: claude-development-journal)
    
.PARAMETER UpdateIndex
    If specified, also updates the main index.json file
    
.EXAMPLE
    .\Sync-LearningTimeline-ToBlob.ps1 -HtmlPath "~/Projects/cloud-resume/site/index.html"
    
.NOTES
    Author: Jason Slater
    Version: 1.0
    Date: 2025-06-05
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory = $true)]
    [ValidateScript({Test-Path $_})]
    [string]$HtmlPath,
    
    [Parameter(Mandatory = $false)]
    [string]$StorageAccountName = "jslaterdevjournal",
    
    [Parameter(Mandatory = $false)]
    [string]$ContainerName = "claude-development-journal",
    
    [Parameter(Mandatory = $false)]
    [switch]$UpdateIndex
)

# Function to parse learning items from HTML
function Extract-LearningItems {
    param([string]$HtmlContent)
    
    Write-Verbose "Extracting learning items from HTML..."
    
    $learningItems = @()
    
    # Regex to match learning items
    $pattern = '(?s)<div class="learning-item"[^>]*>.*?<h3>(.*?)</h3>.*?<strong>Challenge:</strong>\s*(.*?)</p>.*?<strong>Solution:</strong>\s*(.*?)</p>.*?<div class="learning-timestamp">Added:\s*(.*?)</div>'
    
    $matches = [regex]::Matches($HtmlContent, $pattern)
    
    foreach ($match in $matches) {
        $item = [PSCustomObject]@{
            Title = $match.Groups[1].Value.Trim()
            Challenge = $match.Groups[2].Value.Trim()
            Solution = $match.Groups[3].Value.Trim()
            DateAdded = $match.Groups[4].Value.Trim()
            Id = [guid]::NewGuid().ToString()
        }
        
        # Clean up HTML entities
        $item.Title = $item.Title -replace '&amp;', '&' -replace '&lt;', '<' -replace '&gt;', '>'
        $item.Challenge = $item.Challenge -replace '&amp;', '&' -replace '&lt;', '<' -replace '&gt;', '>'
        $item.Solution = $item.Solution -replace '&amp;', '&' -replace '&lt;', '<' -replace '&gt;', '>'
        
        $learningItems += $item
    }
    
    Write-Verbose "Found $($learningItems.Count) learning items"
    return $learningItems
}

# Function to create learning timeline JSON structure
function Create-LearningTimelineJson {
    param([array]$LearningItems)
    
    $timeline = [PSCustomObject]@{
        LastUpdated = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
        TotalItems = $LearningItems.Count
        Items = $LearningItems | Sort-Object DateAdded -Descending
        Metadata = @{
            Source = "slater.cloud learning section"
            SyncMethod = "PowerShell Azure Integration"
            Version = "1.0"
        }
    }
    
    return $timeline | ConvertTo-Json -Depth 10
}

# Main execution
try {
    Write-Host "=== LEARNING TIMELINE SYNC TO AZURE BLOB ===" -ForegroundColor Cyan
    Write-Host "Storage Account: $StorageAccountName" -ForegroundColor Yellow
    Write-Host "Container: $ContainerName" -ForegroundColor Yellow
    
    # Read HTML file
    Write-Host "`nReading HTML file..." -ForegroundColor Yellow
    $htmlContent = Get-Content -Path $HtmlPath -Raw
    
    # Extract learning items
    $learningItems = Extract-LearningItems -HtmlContent $htmlContent
    Write-Host "Extracted $($learningItems.Count) learning items" -ForegroundColor Green
    
    # Display items
    Write-Host "`nLearning items found:" -ForegroundColor Cyan
    $learningItems | ForEach-Object {
        Write-Host "  - $($_.Title)" -ForegroundColor White
        Write-Host "    Date: $($_.DateAdded)" -ForegroundColor Gray
    }
    
    # Create JSON
    $jsonContent = Create-LearningTimelineJson -LearningItems $learningItems
    
    # Connect to Azure
    Write-Host "`nConnecting to Azure..." -ForegroundColor Yellow
    $context = Get-AzContext
    if (-not $context) {
        Write-Host "Not logged in to Azure. Running Connect-AzAccount..." -ForegroundColor Yellow
        Connect-AzAccount
    }
    
    # Get storage context
    Write-Host "Getting storage context..." -ForegroundColor Yellow
    $storageAccount = Get-AzStorageAccount | Where-Object { $_.StorageAccountName -eq $StorageAccountName }
    
    if (-not $storageAccount) {
        throw "Storage account '$StorageAccountName' not found"
    }
    
    $storageContext = $storageAccount.Context
    
    # Upload learning timeline JSON
    $blobName = "learning-timeline/timeline.json"
    
    if ($PSCmdlet.ShouldProcess($blobName, "Upload learning timeline to blob")) {
        Write-Host "`nUploading learning timeline to blob: $blobName" -ForegroundColor Yellow
        
        $tempFile = New-TemporaryFile
        Set-Content -Path $tempFile.FullName -Value $jsonContent
        
        $upload = Set-AzStorageBlobContent `
            -File $tempFile.FullName `
            -Container $ContainerName `
            -Blob $blobName `
            -Context $storageContext `
            -Properties @{"ContentType" = "application/json"} `
            -Force
        
        Remove-Item $tempFile.FullName -Force
        
        Write-Host "Upload successful!" -ForegroundColor Green
        Write-Host "URL: https://$StorageAccountName.blob.core.windows.net/$ContainerName/$blobName" -ForegroundColor Cyan
    }
    
    # Update main index.json if requested
    if ($UpdateIndex) {
        Write-Host "`nUpdating main index.json..." -ForegroundColor Yellow
        
        # Download current index.json
        $indexBlob = Get-AzStorageBlob -Container $ContainerName -Blob "index.json" -Context $storageContext
        $indexContent = $indexBlob.ICloudBlob.DownloadText()
        $index = $indexContent | ConvertFrom-Json
        
        # Add learning timeline reference
        if (-not $index.LearningTimeline) {
            $index | Add-Member -MemberType NoteProperty -Name "LearningTimeline" -Value @{
                Url = "https://$StorageAccountName.blob.core.windows.net/$ContainerName/$blobName"
                ItemCount = $learningItems.Count
                LastSync = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
            }
        } else {
            $index.LearningTimeline.ItemCount = $learningItems.Count
            $index.LearningTimeline.LastSync = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
        }
        
        # Upload updated index
        if ($PSCmdlet.ShouldProcess("index.json", "Update main index with learning timeline reference")) {
            $updatedIndex = $index | ConvertTo-Json -Depth 10
            $tempFile = New-TemporaryFile
            Set-Content -Path $tempFile.FullName -Value $updatedIndex
            
            Set-AzStorageBlobContent `
                -File $tempFile.FullName `
                -Container $ContainerName `
                -Blob "index.json" `
                -Context $storageContext `
                -Properties @{"ContentType" = "application/json"} `
                -Force
            
            Remove-Item $tempFile.FullName -Force
            Write-Host "Index.json updated successfully!" -ForegroundColor Green
        }
    }
    
    Write-Host "`n=== SYNC COMPLETE ===" -ForegroundColor Cyan
    Write-Host "Learning timeline data is now available in Azure Blob Storage" -ForegroundColor Green
    
    # Create local backup
    $backupPath = Join-Path (Split-Path $HtmlPath) "learning-timeline-backup.json"
    $jsonContent | Out-File -FilePath $backupPath -Encoding UTF8
    Write-Host "Local backup saved to: $backupPath" -ForegroundColor Gray
    
} catch {
    Write-Error "Sync failed: $_"
    exit 1
}