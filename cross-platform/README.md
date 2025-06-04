# Cross-Platform PowerShell Scripts

This directory contains PowerShell scripts that are compatible with PowerShell Core and can run on Windows, macOS, and Linux.

## Why Cross-Platform Scripts?

Working between local macOS development and Windows VMs, I needed scripts that work reliably on both platforms without modification. This cross-platform approach provides:

- **Consistency**: Same script behavior across all platforms
- **Flexibility**: Develop on macOS, deploy on Windows (or vice versa)
- **Modern Compatibility**: Built for PowerShell Core 7.x+
- **Simplified Workflow**: No need to maintain separate script versions

## Directory Structure

```
cross-platform/
├── exchange/           # Exchange Online management (cross-platform)
├── mggraph/           # Microsoft Graph operations
│   └── user_management/   # User lookup and management scripts
├── m365/              # Microsoft 365 operations
│   └── groups/        # Group management scripts
└── README.md          # This file
```

## Platform Requirements

- **PowerShell Core 7.0+** (Windows, macOS, Linux)
- **Compatible Modules**:
  - `Microsoft.Graph.*` (all Graph modules)
  - `ExchangeOnlineManagement`
  - Platform-native modules only

## Scripts Included

### Exchange Online (`exchange/`)
- **Set-DistributionListMembership.ps1** - Replace distribution list membership with exact list
- **Add-UserToMultipleGroups.ps1** - Add users to multiple distribution groups and shared mailboxes

### Microsoft Graph User Management (`mggraph/user_management/`)
- **Get-UPNByDisplayName.ps1** - Resolve display names to User Principal Names
- **Quick-UPNLookup.ps1** - Interactive UPN lookup tool
- **getUPN_batch_list.ps1** - Batch UPN resolution

### Microsoft 365 Groups (`m365/groups/`)
- **Copy-M365GroupMembers.ps1** - Copy group membership between M365 groups

## Windows-Only Scripts

For scripts that require Windows PowerShell 5.1 or Windows-only modules (PnP, SPO), see the main `/scripts/` directory. These include:

- SharePoint PnP operations (`/scripts/sharepoint/`)
- OneDrive management (`/scripts/sharepoint/onedrive/`)
- Windows-specific utilities

## Usage Examples

### macOS Development
```bash
# Navigate to cross-platform scripts
cd ~/Projects/pwsh_honeypot/cross-platform

# Run Exchange script
pwsh ./exchange/Set-DistributionListMembership.ps1

# Run Graph user lookup
pwsh ./mggraph/user_management/Quick-UPNLookup.ps1
```

### Windows Production
```powershell
# Same scripts work identically
cd C:\Scripts\pwsh_honeypot\cross-platform

# Run the same commands
.\exchange\Set-DistributionListMembership.ps1
.\mggraph\user_management\Quick-UPNLookup.ps1
```

## Key Features

- **Interactive Input**: All scripts support pasting unformatted lists (names, emails)
- **Progress Reporting**: Visual progress bars and colored output
- **Error Handling**: Graceful handling of failures with detailed feedback
- **Validation**: Pre-flight checks and confirmation prompts
- **Export Options**: CSV/Excel export capabilities where applicable

## Development Notes

- Scripts use `Write-Host` with color support for cross-platform compatibility
- File paths use `Join-Path` and environment variables for platform neutrality
- Module loading includes cross-platform connection methods
- All interactive prompts work consistently across platforms

## Migration from Windows-Only

When migrating scripts to cross-platform:

1. **Check module compatibility** - Ensure no PnP/SPO dependencies
2. **Update paths** - Use `$env:TEMP` instead of `C:\Temp`
3. **Test interactivity** - Verify prompts work on target platforms
4. **Add platform header** - Include cross-platform compatibility note

This approach maintains the full functionality of Windows-specific scripts while providing modern, cross-platform alternatives for everyday tasks.