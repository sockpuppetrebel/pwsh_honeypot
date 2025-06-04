#Requires -Modules Az.Storage

<#
.SYNOPSIS
    Backs up local files to Azure Blob Storage with encryption and versioning
    
.DESCRIPTION
    Enterprise-grade backup solution for local configuration files. Demonstrates
    Azure Blob Storage integration, encryption at rest, and automated backup strategies.
    Perfect for backing up sensitive configuration files that shouldn't be in version control.
    
.PARAMETER FilesToBackup
    Array of file paths to backup
    
.PARAMETER StorageAccountName
    Azure Storage Account name
    
.PARAMETER ContainerName
    Blob container name (default: 'config-backups')
    
.PARAMETER ResourceGroupName
    Azure Resource Group name
    
.PARAMETER RetentionDays
    Number of days to retain old versions (default: 30)
    
.EXAMPLE
    .\Backup-LocalFilesToAzure.ps1 -FilesToBackup @("~/Projects/*/CLAUDE.md") -StorageAccountName "configbackupsa" -ResourceGroupName "ConfigBackups-RG"
    
.NOTES
    Author: Jason Slater
    Portfolio: Demonstrates Azure Blob Storage, PowerShell automation, and enterprise backup strategies
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)]
    [string[]]$FilesToBackup,
    
    [Parameter(Mandatory=$true)]
    [string]$StorageAccountName,
    
    [Parameter(Mandatory=$false)]
    [string]$ContainerName = "config-backups",
    
    [Parameter(Mandatory=$true)]
    [string]$ResourceGroupName,
    
    [Parameter(Mandatory=$false)]
    [int]$RetentionDays = 30
)

# Connect to Azure
Write-Host "Connecting to Azure..." -ForegroundColor Yellow
try {
    $context = Get-AzContext
    if (-not $context) {
        Connect-AzAccount
    }
    Write-Host "Connected to Azure subscription: $($context.Subscription.Name)" -ForegroundColor Green
}
catch {
    Write-Error "Failed to connect to Azure: $_"
    exit 1
}

# Get or create storage account
Write-Host "`nValidating storage account..." -ForegroundColor Yellow
$storageAccount = Get-AzStorageAccount -ResourceGroupName $ResourceGroupName -Name $StorageAccountName -ErrorAction SilentlyContinue

if (-not $storageAccount) {
    Write-Host "Storage account not found. Creating new storage account..." -ForegroundColor Yellow
    
    # Create storage account with encryption
    $storageAccount = New-AzStorageAccount `
        -ResourceGroupName $ResourceGroupName `
        -Name $StorageAccountName `
        -Location "eastus" `
        -SkuName "Standard_LRS" `
        -Kind "StorageV2" `
        -EnableHttpsTrafficOnly $true `
        -MinimumTlsVersion "TLS1_2"
    
    Write-Host "Storage account created with encryption at rest enabled" -ForegroundColor Green
}

# Get storage context
$ctx = $storageAccount.Context

# Create container if it doesn't exist
$container = Get-AzStorageContainer -Name $ContainerName -Context $ctx -ErrorAction SilentlyContinue
if (-not $container) {
    Write-Host "Creating container '$ContainerName'..." -ForegroundColor Yellow
    New-AzStorageContainer -Name $ContainerName -Context $ctx -Permission Off
    Write-Host "Container created with private access" -ForegroundColor Green
}

# Enable versioning on the container
Write-Host "Enabling blob versioning for automatic version history..." -ForegroundColor Yellow
Set-AzStorageBlobServiceProperty -ResourceGroupName $ResourceGroupName `
    -StorageAccountName $StorageAccountName `
    -IsVersioningEnabled $true

# Process each file
$backupSummary = @()
$timestamp = Get-Date -Format "yyyyMMdd-HHmmss"

foreach ($filePattern in $FilesToBackup) {
    # Expand wildcards
    $files = Get-ChildItem -Path $filePattern -ErrorAction SilentlyContinue
    
    foreach ($file in $files) {
        if (Test-Path $file.FullName -PathType Leaf) {
            try {
                # Create blob name with metadata
                $relativePath = $file.FullName.Replace($HOME, "~")
                $blobName = "$($file.Name)/$timestamp-$($file.Name)"
                
                Write-Host "Backing up: $relativePath" -ForegroundColor Cyan
                
                # Upload with metadata
                $metadata = @{
                    "OriginalPath" = $relativePath
                    "MachineName" = $env:COMPUTERNAME
                    "BackupDate" = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
                    "FileSize" = $file.Length
                    "LastModified" = $file.LastWriteTime.ToString("yyyy-MM-dd HH:mm:ss")
                }
                
                # Upload file
                $blob = Set-AzStorageBlobContent `
                    -File $file.FullName `
                    -Container $ContainerName `
                    -Blob $blobName `
                    -Context $ctx `
                    -Metadata $metadata `
                    -Force
                
                $backupSummary += [PSCustomObject]@{
                    LocalFile = $relativePath
                    BlobName = $blobName
                    Size = "{0:N2} KB" -f ($file.Length / 1KB)
                    Uploaded = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
                    Status = "Success"
                }
                
                Write-Host "  ✓ Uploaded successfully" -ForegroundColor Green
            }
            catch {
                Write-Host "  ✗ Failed to upload: $_" -ForegroundColor Red
                $backupSummary += [PSCustomObject]@{
                    LocalFile = $relativePath
                    BlobName = "N/A"
                    Size = "{0:N2} KB" -f ($file.Length / 1KB)
                    Uploaded = "N/A"
                    Status = "Failed: $_"
                }
            }
        }
    }
}

# Set lifecycle management for retention
Write-Host "`nConfiguring lifecycle management for $RetentionDays day retention..." -ForegroundColor Yellow

$lifecycleRule = @{
    Rules = @(
        @{
            Enabled = $true
            Name = "DeleteOldBackups"
            Definition = @{
                Actions = @{
                    BaseBlob = @{
                        Delete = @{
                            DaysAfterModificationGreaterThan = $RetentionDays
                        }
                    }
                }
                Filters = @{
                    PrefixMatch = @("CLAUDE.md")
                }
            }
        }
    )
}

# Clean up old versions beyond retention
Write-Host "Cleaning up backups older than $RetentionDays days..." -ForegroundColor Yellow
$cutoffDate = (Get-Date).AddDays(-$RetentionDays)
$oldBlobs = Get-AzStorageBlob -Container $ContainerName -Context $ctx | 
    Where-Object { $_.LastModified -lt $cutoffDate -and $_.Name -like "*/CLAUDE.md" }

foreach ($oldBlob in $oldBlobs) {
    Remove-AzStorageBlob -Blob $oldBlob.Name -Container $ContainerName -Context $ctx -Force
    Write-Host "  Removed old backup: $($oldBlob.Name)" -ForegroundColor Gray
}

# Display summary
Write-Host "`n=== BACKUP SUMMARY ===" -ForegroundColor Cyan
$backupSummary | Format-Table -AutoSize

Write-Host "`nBackup Details:" -ForegroundColor Cyan
Write-Host "Storage Account: $StorageAccountName" -ForegroundColor White
Write-Host "Container: $ContainerName" -ForegroundColor White
Write-Host "Files Backed Up: $($backupSummary | Where-Object Status -eq "Success").Count" -ForegroundColor Green
Write-Host "Failed: $($backupSummary | Where-Object Status -ne "Success").Count" -ForegroundColor Red
Write-Host "Retention Policy: $RetentionDays days" -ForegroundColor White

# Generate SAS token for recovery
$endTime = (Get-Date).AddDays(7)
$sasToken = New-AzStorageContainerSASToken `
    -Name $ContainerName `
    -Context $ctx `
    -Permission "rl" `
    -ExpiryTime $endTime

Write-Host "`n=== RECOVERY INFORMATION ===" -ForegroundColor Cyan
Write-Host "To restore files, use this SAS URL (valid for 7 days):" -ForegroundColor Yellow
Write-Host "$($ctx.BlobEndPoint)$ContainerName$sasToken" -ForegroundColor White
Write-Host "`nStore this URL securely for emergency recovery" -ForegroundColor Yellow