# PowerShell Scripts Collection

A comprehensive, well-organized collection of PowerShell scripts for managing Microsoft 365 services, with both Windows-specific and cross-platform versions. Designed for IT professionals who work across different operating systems and need reliable automation tools.

## Platform Support

**All scripts have been refactored for cross-platform compatibility on macOS unless explicitly stated otherwise.**

- **Cross-Platform Scripts** (`/cross-platform/`) - Work on Windows, macOS, and Linux with PowerShell Core 7.x+
- **Windows-Only Scripts** (`/scripts/`) - Require Windows PowerShell 5.1 or Windows-specific modules (PnP, SPO)

This dual approach addresses the real-world challenge of bouncing between local macOS development and Windows VMs while maintaining full functionality across environments.

## Directory Structure

```
pwsh-honeypot/
├── Config.ps1                           # Centralized configuration
├── cross-platform/                      # 🚀 Cross-platform scripts (Windows/macOS/Linux)
│   ├── exchange/                       # Exchange Online management
│   ├── mggraph/                        # Microsoft Graph operations
│   │   └── user_management/            # User lookup and administration
│   ├── m365/                          # Microsoft 365 operations
│   │   └── groups/                     # Group management scripts
│   └── README.md                       # Cross-platform documentation
├── scripts/                            # 🪟 Windows-only scripts (PowerShell 5.1)
│   ├── exchange/                       # Exchange Online management
│   ├── mggraph/                        # Microsoft Graph API scripts
│   │   ├── group_management/           # Azure AD group operations
│   │   ├── user_management/            # User administration
│   │   └── utilities/                  # Mixed utilities (Python/PowerShell)
│   └── sharepoint/                     # SharePoint Online management (PnP/SPO)
│       ├── permissions/                # Permission management (removal, modification)
│       ├── auditing/                   # Permission auditing and reporting  
│       ├── onedrive/                   # OneDrive-specific operations
│       └── utilities/                  # Authentication, testing, diagnostics
├── certificates/                        # Authentication certificates  
│   ├── azure/                          # Azure app certificates
│   └── graph/                          # Microsoft Graph certificates
├── output/                             # Script outputs and reports
│   ├── sharepoint/audit_reports/       # SharePoint audit reports
│   ├── sharepoint/permission_reports/  # Permission scan results
│   ├── mggraph/                        # Microsoft Graph outputs
│   └── exchange/                       # Exchange outputs
└── archive/                            # Deprecated scripts
```

## Key Features

### SharePoint Permission Management
- **Bulletproof Permission Removal** - Safe, parallel processing with confirmation prompts
- **Comprehensive Auditing** - Deep permission analysis across all sites and OneDrive
- **Identity Resolution** - Find old accounts, cached profiles, and permission conflicts
- **High Performance** - Parallel processing for 7000+ sites in minutes

### User Administration  
- **Group Management** - Azure AD group operations with detailed reporting
- **UPN Discovery** - Find and resolve User Principal Name issues
- **License Analysis** - Comprehensive license and permission analysis

### Authentication & Security
- **Certificate-based Authentication** - Secure app-only authentication
- **Connection Testing** - Comprehensive authentication diagnostics
- **Certificate Management** - Generate and manage authentication certificates

## Quick Start Guide

### 1. Configuration
The `Config.ps1` file contains centralized configuration for all scripts:
```powershell
# Load configuration
. .\Config.ps1

# Certificate paths and output directories are automatically configured
```

### 2. Common Use Cases

#### Permission Cleanup for Rehired Employee
```powershell
# Comprehensive scan for any permission remnants
.\scripts\sharepoint\utilities\find_all_user_identities.ps1 -UserEmail "user@domain.com"

# Fast parallel permission removal with safety prompts
.\scripts\sharepoint\permissions\remove_permissions_parallel_pnp.ps1 -UserEmail "user@domain.com"

# OneDrive-specific cleanup
.\scripts\sharepoint\onedrive\onedrive_fast_comprehensive.ps1 -UserEmail "user@domain.com"
```

#### Permission Auditing
```powershell
# Quick permission audit
.\scripts\sharepoint\auditing\fast_scan_pnp.ps1 -UserEmail "user@domain.com"

# Comprehensive permission report
.\scripts\sharepoint\auditing\smart_search_pnp.ps1 -UserEmail "user@domain.com"
```

#### User Administration
```powershell
# Export user's group memberships
.\scripts\sharepoint\utilities\export_user_aad_groups.ps1 -UserEmail "user@domain.com"

# Add user to security group
.\scripts\mggraph\group_management\add_members_sg.ps1 -GroupName "Finance-Team" -UserEmail "user@domain.com"
```

## Safety Features

All scripts include comprehensive safety measures:
- **Confirmation prompts** before making changes
- **Detailed logging** with timestamps
- **Error handling** and rollback capabilities
- **Progress reporting** for long operations
- **CSV export** for audit trails
- **Parallel processing** with throttling controls

## Prerequisites

### PowerShell Modules
```powershell
Install-Module -Name PnP.PowerShell -Force
Install-Module -Name Microsoft.Graph -Force
```

### Azure App Registration
- App registration with appropriate Microsoft Graph and SharePoint permissions
- Certificate-based authentication configured
- Certificates stored in `/certificates/azure/` directory

### Permissions Required
- **SharePoint**: Sites.FullControl.All, User.Read.All
- **Microsoft Graph**: Group.ReadWrite.All, User.ReadWrite.All, Directory.Read.All

## Documentation

Each script category has detailed documentation:
- [SharePoint Permissions](scripts/sharepoint/permissions/README.md) - Permission management scripts
- [SharePoint Auditing](scripts/sharepoint/auditing/README.md) - Audit and reporting scripts
- [OneDrive Management](scripts/sharepoint/onedrive/README.md) - OneDrive-specific operations
- [SharePoint Utilities](scripts/sharepoint/utilities/README.md) - Authentication and diagnostics
- [Microsoft Graph Groups](scripts/mggraph/group_management/README.md) - Group management
- [Microsoft Graph Users](scripts/mggraph/user_management/README.md) - User administration

## Troubleshooting

### Authentication Issues
```powershell
# Test PnP connection
.\scripts\sharepoint\utilities\test_pnp_connection.ps1

# Test Microsoft Graph authentication
.\scripts\sharepoint\utilities\test_auth.ps1

# Diagnose certificate issues
.\scripts\sharepoint\utilities\diagnose_cert.ps1
```

### Permission Issues
```powershell
# Find all identity variations (great for rehired employees)
.\scripts\sharepoint\utilities\find_all_user_identities.ps1 -UserEmail "user@domain.com"

# Search by Object ID patterns
.\scripts\sharepoint\utilities\find_user_by_objectid.ps1 -UserEmail "user@domain.com"
```

## Output Management

All scripts automatically organize outputs:
- **Audit reports** → `/output/sharepoint/audit_reports/`
- **Permission reports** → `/output/sharepoint/permission_reports/`
- **Microsoft Graph outputs** → `/output/mggraph/`
- **Exchange outputs** → `/output/exchange/`

Files are timestamped for easy tracking and include comprehensive metadata.

## Performance

Optimized for enterprise environments:
- **Parallel processing** across multiple threads
- **Batch operations** for API efficiency
- **Smart throttling** to prevent API limits
- **Progress indicators** for long operations

**Typical performance**: 7000+ SharePoint sites scanned in under 5 minutes

## License

These scripts are provided as-is for educational and administrative purposes. Always test in a development environment before running in production.
