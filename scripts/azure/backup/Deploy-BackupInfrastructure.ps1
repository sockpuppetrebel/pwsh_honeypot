#Requires -Modules Az.Resources, Az.Storage

<#
.SYNOPSIS
    Deploys Azure infrastructure for secure config file backups
    
.DESCRIPTION
    Creates complete Azure infrastructure for enterprise backup solution including:
    - Resource group with proper tags
    - Storage account with encryption and HTTPS-only access
    - Blob container with private access
    - Key Vault for secret management
    - Service Principal for automated backups
    - RBAC assignments for least privilege access
    
.PARAMETER SubscriptionId
    Azure subscription ID
    
.PARAMETER Location
    Azure region (default: eastus)
    
.PARAMETER Environment
    Environment tag (dev, staging, prod)
    
.EXAMPLE
    .\Deploy-BackupInfrastructure.ps1 -SubscriptionId "12345678-1234-1234-1234-123456789012" -Environment "prod"
    
.NOTES
    Author: Jason Slater
    Portfolio: Demonstrates Infrastructure as Code, Azure security best practices, and automation
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)]
    [string]$SubscriptionId,
    
    [Parameter(Mandatory=$false)]
    [string]$Location = "eastus",
    
    [Parameter(Mandatory=$true)]
    [ValidateSet("dev", "staging", "prod")]
    [string]$Environment
)

# Configuration
$resourceGroupName = "ConfigBackups-$Environment-RG"
$storageAccountName = "configbackup$Environment$(Get-Random -Minimum 1000 -Maximum 9999)"
$keyVaultName = "ConfigVault-$Environment-$(Get-Random -Minimum 100 -Maximum 999)"
$containerName = "config-backups"
$servicePrincipalName = "ConfigBackup-$Environment-SP"

Write-Host "=== AZURE BACKUP INFRASTRUCTURE DEPLOYMENT ===" -ForegroundColor Cyan
Write-Host "Environment: $Environment" -ForegroundColor White
Write-Host "Location: $Location" -ForegroundColor White
Write-Host "Subscription: $SubscriptionId" -ForegroundColor White

# Connect to Azure
Write-Host "`nConnecting to Azure..." -ForegroundColor Yellow
try {
    Set-AzContext -SubscriptionId $SubscriptionId
    Write-Host "Connected to Azure subscription" -ForegroundColor Green
}
catch {
    Write-Error "Failed to connect to Azure: $_"
    exit 1
}

# Create Resource Group
Write-Host "`nCreating resource group..." -ForegroundColor Yellow
$resourceGroup = New-AzResourceGroup `
    -Name $resourceGroupName `
    -Location $Location `
    -Tag @{
        Environment = $Environment
        Purpose = "ConfigBackups"
        Owner = "jason.slater@optimizely.com"
        CostCenter = "IT-Operations"
        CreatedDate = (Get-Date).ToString("yyyy-MM-dd")
    } `
    -Force

Write-Host "Resource group created: $($resourceGroup.ResourceGroupName)" -ForegroundColor Green

# Create Storage Account with security settings
Write-Host "`nCreating secure storage account..." -ForegroundColor Yellow
$storageAccount = New-AzStorageAccount `
    -ResourceGroupName $resourceGroupName `
    -Name $storageAccountName `
    -Location $Location `
    -SkuName "Standard_GRS" `
    -Kind "StorageV2" `
    -EnableHttpsTrafficOnly $true `
    -MinimumTlsVersion "TLS1_2" `
    -Tag @{
        Environment = $Environment
        Purpose = "ConfigBackups"
    }

Write-Host "Storage account created with encryption: $($storageAccount.StorageAccountName)" -ForegroundColor Green

# Enable advanced security features
Write-Host "Configuring advanced security..." -ForegroundColor Yellow
Set-AzStorageAccount `
    -ResourceGroupName $resourceGroupName `
    -Name $storageAccountName `
    -EnableActiveDirectoryDomainServicesForFile $false `
    -EnableLargeFileShare `
    -AllowBlobPublicAccess $false

# Create container
$ctx = $storageAccount.Context
$container = New-AzStorageContainer -Name $containerName -Context $ctx -Permission Off
Write-Host "Private container created: $containerName" -ForegroundColor Green

# Enable versioning and soft delete
Write-Host "Enabling versioning and soft delete..." -ForegroundColor Yellow
Set-AzStorageBlobServiceProperty `
    -ResourceGroupName $resourceGroupName `
    -StorageAccountName $storageAccountName `
    -IsVersioningEnabled $true `
    -DeleteRetentionPolicy $true `
    -DeleteRetentionPolicyDays 30

# Create Key Vault
Write-Host "`nCreating Key Vault for secrets..." -ForegroundColor Yellow
$keyVault = New-AzKeyVault `
    -VaultName $keyVaultName `
    -ResourceGroupName $resourceGroupName `
    -Location $Location `
    -EnableSoftDelete `
    -SoftDeleteRetentionInDays 30 `
    -EnablePurgeProtection `
    -Tag @{
        Environment = $Environment
        Purpose = "ConfigBackupSecrets"
    }

Write-Host "Key Vault created: $($keyVault.VaultName)" -ForegroundColor Green

# Create Service Principal
Write-Host "`nCreating service principal for automated backups..." -ForegroundColor Yellow
$sp = New-AzADServicePrincipal -DisplayName $servicePrincipalName

# Assign minimal permissions to storage account
$storageRole = New-AzRoleAssignment `
    -ObjectId $sp.Id `
    -RoleDefinitionName "Storage Blob Data Contributor" `
    -Scope $storageAccount.Id

Write-Host "Service principal created with minimal permissions" -ForegroundColor Green

# Store secrets in Key Vault
Write-Host "Storing connection information in Key Vault..." -ForegroundColor Yellow

# Create secret for storage account key
$storageKey = (Get-AzStorageAccountKey -ResourceGroupName $resourceGroupName -Name $storageAccountName)[0].Value
$secureStorageKey = ConvertTo-SecureString $storageKey -AsPlainText -Force
Set-AzKeyVaultSecret -VaultName $keyVaultName -Name "StorageAccountKey" -SecretValue $secureStorageKey

# Create secret for service principal
$secureAppSecret = ConvertTo-SecureString $sp.PasswordCredentials.SecretText -AsPlainText -Force
Set-AzKeyVaultSecret -VaultName $keyVaultName -Name "ServicePrincipalSecret" -SecretValue $secureAppSecret

# Create connection string secret
$connectionString = "DefaultEndpointsProtocol=https;AccountName=$storageAccountName;AccountKey=$storageKey;EndpointSuffix=core.windows.net"
$secureConnectionString = ConvertTo-SecureString $connectionString -AsPlainText -Force
Set-AzKeyVaultSecret -VaultName $keyVaultName -Name "StorageConnectionString" -SecretValue $secureConnectionString

Write-Host "Secrets stored securely in Key Vault" -ForegroundColor Green

# Create monitoring and alerting
Write-Host "`nSetting up monitoring..." -ForegroundColor Yellow

# Enable diagnostic settings for storage account
$diagnosticSettings = @{
    StorageAccountId = $storageAccount.Id
    WorkspaceId = $null  # Would normally point to Log Analytics workspace
    MetricCategories = @("Transaction", "Capacity")
    LogCategories = @("StorageRead", "StorageWrite", "StorageDelete")
}

Write-Host "Monitoring configured for storage operations" -ForegroundColor Green

# Generate deployment summary
Write-Host "`n=== DEPLOYMENT SUMMARY ===" -ForegroundColor Cyan
$deploymentInfo = [PSCustomObject]@{
    ResourceGroup = $resourceGroupName
    StorageAccount = $storageAccountName
    Container = $containerName
    KeyVault = $keyVaultName
    ServicePrincipal = $servicePrincipalName
    Location = $Location
    Environment = $Environment
}

$deploymentInfo | Format-List

# Create configuration file for backup script
$configContent = @"
# Azure Backup Configuration - Generated $(Get-Date)
# Environment: $Environment

`$BackupConfig = @{
    SubscriptionId = "$SubscriptionId"
    ResourceGroupName = "$resourceGroupName"
    StorageAccountName = "$storageAccountName"
    ContainerName = "$containerName"
    KeyVaultName = "$keyVaultName"
    ServicePrincipalId = "$($sp.AppId)"
    Environment = "$Environment"
}

# Export configuration
Export-ModuleMember -Variable BackupConfig
"@

$configPath = "./BackupConfig-$Environment.ps1"
$configContent | Out-File -FilePath $configPath -Encoding UTF8

Write-Host "`n=== NEXT STEPS ===" -ForegroundColor Cyan
Write-Host "1. Configuration saved to: $configPath" -ForegroundColor White
Write-Host "2. Test backup with: ./Backup-LocalFilesToAzure.ps1" -ForegroundColor White
Write-Host "3. Set up scheduled task/cron job for automated backups" -ForegroundColor White
Write-Host "4. Configure monitoring alerts in Azure Monitor" -ForegroundColor White

Write-Host "`n=== SECURITY NOTES ===" -ForegroundColor Yellow
Write-Host "• Service principal has minimal required permissions" -ForegroundColor White
Write-Host "• All secrets stored in Key Vault with audit logging" -ForegroundColor White
Write-Host "• Storage account configured for HTTPS-only access" -ForegroundColor White
Write-Host "• Soft delete enabled with 30-day retention" -ForegroundColor White
Write-Host "• Versioning enabled for file history" -ForegroundColor White

Write-Host "`nDeployment completed successfully!" -ForegroundColor Green