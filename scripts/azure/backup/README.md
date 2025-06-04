# Azure Configuration Backup Solution

Enterprise-grade backup solution for sensitive configuration files that shouldn't be stored in version control. Demonstrates Azure security best practices and Infrastructure as Code principles.

## Architecture Overview

```
┌─────────────────────────────────────────────────────────────────┐
│                          Azure Tenant                          │
├─────────────────────────────────────────────────────────────────┤
│  Resource Group: ConfigBackups-{env}-RG                        │
│  ┌─────────────────┐  ┌─────────────────┐  ┌─────────────────┐ │
│  │   Storage Acct  │  │    Key Vault    │  │ Service Principal│ │
│  │  - Encryption   │  │  - Secrets      │  │ - Min Permissions│ │
│  │  - Versioning   │  │  - Audit Logs   │  │ - Blob Contributor│ │
│  │  - Soft Delete  │  │  - Purge Protect│  │                 │ │
│  │  - HTTPS Only   │  │                 │  │                 │ │
│  └─────────────────┘  └─────────────────┘  └─────────────────┘ │
│           │                      │                      │        │
│  ┌─────────────────┐  ┌─────────────────┐                      │
│  │ Blob Container  │  │  Lifecycle Mgmt │                      │
│  │ - Private Access│  │ - Auto Cleanup  │                      │
│  │ - File Versions │  │ - Retention     │                      │
│  └─────────────────┘  └─────────────────┘                      │
└─────────────────────────────────────────────────────────────────┘
                                │
                                ▼
┌─────────────────────────────────────────────────────────────────┐
│                      Local Development                          │
│  ┌─────────────────┐  ┌─────────────────┐  ┌─────────────────┐ │
│  │  macOS/Windows  │  │ PowerShell Core │  │  Backup Script  │ │
│  │                 │  │                 │  │                 │ │
│  │  ├── Project1/  │  │ - Az.Storage    │  │ - Automated     │ │
│  │  │   CLAUDE.md  │  │ - Az.KeyVault   │  │ - Encrypted     │ │
│  │  ├── Project2/  │  │ - Authentication│  │ - Versioned     │ │
│  │  │   CLAUDE.md  │  │                 │  │ - Retention     │ │
│  │  └── Project3/  │  │                 │  │                 │ │
│  │      CLAUDE.md  │  │                 │  │                 │ │
│  └─────────────────┘  └─────────────────┘  └─────────────────┘ │
└─────────────────────────────────────────────────────────────────┘
```

## Problem Statement

Configuration files containing AI instructions, API keys, or environment-specific settings need to be:
- Backed up securely
- Accessible across multiple development machines
- Kept out of version control
- Retained with proper lifecycle management

## Solution Components

### 1. Infrastructure Deployment (`Deploy-BackupInfrastructure.ps1`)
- **Resource Group** with proper tagging and cost allocation
- **Storage Account** with geo-redundancy and encryption at rest
- **Key Vault** for secure credential storage with purge protection
- **Service Principal** with minimal required permissions (Blob Data Contributor only)
- **Lifecycle Management** for automated cleanup and cost optimization

### 2. Backup Automation (`Backup-LocalFilesToAzure.ps1`)
- **Cross-platform compatibility** (Windows/macOS/Linux)
- **Intelligent file discovery** using wildcard patterns
- **Metadata enrichment** with machine name, timestamps, and file attributes
- **Version management** with automatic old version cleanup
- **Error handling** with detailed logging and retry logic

### 3. Security Features
- **Encryption at rest** for all stored data
- **HTTPS-only communication** with TLS 1.2 minimum
- **Private blob containers** with no public access
- **Soft delete protection** with 30-day retention
- **Audit logging** for all access and operations

## Technical Highlights

### Infrastructure as Code
```powershell
# Demonstrates enterprise Azure deployment patterns
New-AzStorageAccount -EnableHttpsTrafficOnly $true -MinimumTlsVersion "TLS1_2"
Set-AzStorageBlobServiceProperty -IsVersioningEnabled $true -DeleteRetentionPolicy $true
New-AzKeyVault -EnableSoftDelete -EnablePurgeProtection
```

### Automated Lifecycle Management
```powershell
# Configures automatic cleanup to prevent storage cost bloat
$lifecycleRule = @{
    Actions = @{
        BaseBlob = @{ Delete = @{ DaysAfterModificationGreaterThan = 30 } }
    }
}
```

### Cross-Platform File Discovery
```powershell
# Works identically on Windows, macOS, and Linux
$FilesToBackup = @("~/Projects/*/CLAUDE.md", "~/Config/*.conf")
foreach ($filePattern in $FilesToBackup) {
    $files = Get-ChildItem -Path $filePattern -ErrorAction SilentlyContinue
}
```

## Security Architecture

### Principle of Least Privilege
- Service Principal has only Blob Data Contributor role
- No management plane permissions granted
- Scoped to specific storage account only

### Defense in Depth
1. **Network**: HTTPS-only, no public blob access
2. **Identity**: Azure AD authentication with service principals
3. **Data**: Encryption at rest and in transit
4. **Application**: Input validation and error handling
5. **Monitoring**: Audit logs and diagnostic settings

### Secret Management
- All connection strings stored in Key Vault
- No secrets in code or configuration files
- Automatic secret rotation capabilities

## Cost Optimization

### Storage Tiers
- Hot tier for recent backups (0-30 days)
- Automatic lifecycle transitions to cool/archive tiers
- Geo-redundant storage for disaster recovery

### Automated Cleanup
- Old versions automatically deleted after retention period
- Failed upload artifacts cleaned up
- Monitoring to prevent unexpected cost growth

## Deployment Guide

### Prerequisites
```powershell
Install-Module Az.Storage, Az.Resources, Az.KeyVault -Force
```

### Environment Setup
```powershell
# Deploy infrastructure
.\Deploy-BackupInfrastructure.ps1 -SubscriptionId "12345678-..." -Environment "prod"

# Test backup functionality
.\Backup-LocalFilesToAzure.ps1 -FilesToBackup @("~/Projects/*/CLAUDE.md") -StorageAccountName "configbackupprod1234" -ResourceGroupName "ConfigBackups-prod-RG"
```

### Automation Options
```bash
# macOS cron job for daily backups
0 2 * * * pwsh /path/to/Backup-LocalFilesToAzure.ps1 -FilesToBackup @("~/Projects/*/CLAUDE.md")

# Windows Task Scheduler
schtasks /create /tn "ConfigBackup" /tr "pwsh.exe -File C:\Scripts\Backup-LocalFilesToAzure.ps1"
```

## Portfolio Demonstration

This solution showcases:

### Azure Expertise
- **Storage Account configuration** with enterprise security settings
- **Key Vault integration** for secret management
- **Service Principal management** with RBAC
- **Lifecycle policies** for cost optimization

### PowerShell Development
- **Cross-platform scripting** compatible with PowerShell Core
- **Error handling and logging** with detailed feedback
- **Parameter validation** and input sanitization
- **Modular design** for reusability

### Security Best Practices
- **Zero-trust architecture** with minimal permissions
- **Encryption everywhere** (rest and transit)
- **Audit logging** for compliance requirements
- **Soft delete protection** against accidental loss

### DevOps Integration
- **Infrastructure as Code** for reproducible deployments
- **Automated backup strategies** with retention policies
- **Cost optimization** through lifecycle management
- **Monitoring and alerting** capabilities

## Business Value

- **Risk Mitigation**: Protects against configuration loss
- **Compliance**: Audit trails and encryption meet enterprise requirements  
- **Cost Control**: Automated lifecycle management prevents storage bloat
- **Scalability**: Solution scales from individual developers to enterprise teams
- **Portability**: Works across Windows, macOS, and Linux development environments