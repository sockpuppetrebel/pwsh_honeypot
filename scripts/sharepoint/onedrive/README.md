# OneDrive Management Scripts

This directory contains scripts specifically for managing OneDrive sites and permissions.

## Scripts

### OneDrive Permission Scanning
- `onedrive_comprehensive_scan.ps1` - **[COMPREHENSIVE]** Full OneDrive permission audit including subsites
- `onedrive_fast_comprehensive.ps1` - **[FAST]** Optimized OneDrive scanning without subsite discovery

## Key Features

### Comprehensive Scanning
The OneDrive scripts check ALL possible permission types:
- **SharePoint Groups** (Owner, Member, Visitor)
- **Site Collection Administrators** 
- **Direct user permissions**
- **Document Library permissions**
- **Folder-level permissions** (shared folders)
- **Cached user profiles** in User Information Lists

### Safety & Confirmation
- **Always prompts** before removing permissions
- **Multiple removal options**: All sites, select specific sites, or exit
- **Detailed progress reporting** during operations
- **Comprehensive error handling**

## Usage Examples

### Fast OneDrive Scan (Recommended)
```powershell
# Quick scan of all OneDrive sites
.\onedrive_fast_comprehensive.ps1 -UserEmail "user@domain.com"

# Scan until 10 sites with permissions found
.\onedrive_fast_comprehensive.ps1 -UserEmail "user@domain.com" -MaxSitesToFind 10
```

### Comprehensive OneDrive Scan (Thorough)
```powershell
# Full comprehensive scan including subsites (slower)
.\onedrive_comprehensive_scan.ps1 -UserEmail "user@domain.com"
```

## Typical Use Cases

### Rehired Employee Cleanup
Perfect for finding old permissions when employees are rehired:
- Scans all OneDrive sites for any permission remnants
- Finds cached profiles that might block new access
- Identifies shared folder permissions that persist

### Permission Verification
Ideal for verifying complete permission removal:
- Confirms zero permissions exist across all OneDrive
- Provides definitive proof of clean state
- Generates audit trail for compliance

## Performance

- **Parallel processing** across 15+ threads for speed
- **Smart batching** for optimal API utilization
- **Progress indicators** showing real-time status
- **Optimized for large tenants** (3000+ OneDrive sites)

Typical performance: **3000+ sites scanned in under 1 minute**