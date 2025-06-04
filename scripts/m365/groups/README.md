# M365 Group Management Scripts

This directory contains PowerShell scripts for managing Microsoft 365 groups, including migration, member management, and configuration fixes.

## Scripts

### Copy-M365GroupMembers.ps1

Copies all members and owners from one M365 group to another. This is useful for:
- Fixing incorrectly configured groups
- Migrating groups between domains
- Creating backup groups before making changes
- Staging group replacements

#### Prerequisites

- Microsoft Graph PowerShell SDK
- Permissions: `Group.ReadWrite.All`, `Directory.ReadWrite.All`
- Global Admin or Groups Administrator role

#### Usage Examples

**Basic usage - create new group and copy members:**
```powershell
.\Copy-M365GroupMembers.ps1 -SourceGroupEmail "oldgroup@domain.com" -NewGroupDisplayName "New Group Name" -NewGroupMailNickname "newgroup"
```

**Stage and scale approach for domain migration:**
```powershell
# Step 1: Create staging group with temporary alias
.\Copy-M365GroupMembers.ps1 -SourceGroupEmail "erg-mentallyforward@episerver99.onmicrosoft.com" -NewGroupDisplayName "ERG Mentally Forward (New)" -NewGroupMailNickname "erg-mentallyforward-new"

# Step 2: Test the new group (email, Outlook access, etc.)

# Step 3: Once verified, swap aliases using Graph API
```

**Copy to existing group:**
```powershell
.\Copy-M365GroupMembers.ps1 -SourceGroupEmail "source@domain.com" -NewGroupMailNickname "existing-group" -SkipGroupCreation
```

**Test run with WhatIf:**
```powershell
.\Copy-M365GroupMembers.ps1 -SourceGroupEmail "source@domain.com" -NewGroupDisplayName "Test Group" -NewGroupMailNickname "testgroup" -WhatIf
```

#### Parameters

- **SourceGroupEmail** (Required): Email address of the source group
- **NewGroupDisplayName** (Required): Display name for the new group
- **NewGroupMailNickname** (Required): Mail alias for the new group
- **GroupVisibility**: Private (default) or Public
- **SkipGroupCreation**: Use existing group instead of creating new
- **WhatIf**: Preview changes without making them

#### Features

- Automatic duplicate detection (won't add existing members/owners)
- Progress reporting with colored output
- Error handling for individual member failures
- Detailed summary of operations
- Support for both creating new groups and updating existing ones

#### Common Scenarios

**Domain Migration (Stage & Scale)**
1. Create new group with temporary alias
2. Copy all members and owners
3. Validate functionality (email, Teams, SharePoint)
4. Remove or rename old group
5. Update new group alias to match original

**Fix Incorrectly Configured Groups**
- Copy members to properly configured group
- Maintain ownership structure
- No service disruption during migration

#### Notes

- Owners are automatically added as members, so they're processed first
- The script waits 10 seconds after group creation for provisioning
- All operations are logged with verbose output available
- Failed member additions don't stop the script, allowing partial completion

## Authentication

These scripts support multiple authentication methods:

1. **Interactive (default)**: Opens browser for authentication
2. **Device Code**: For headless/remote scenarios
3. **Certificate**: For automation (requires app registration)

Example with certificate auth:
```powershell
Connect-MgGraph -ClientId "your-app-id" -TenantId "your-tenant-id" -CertificateThumbprint "cert-thumbprint"
.\Copy-M365GroupMembers.ps1 ...
```

## Troubleshooting

**"Group already exists" error**
- Use `-SkipGroupCreation` flag to add members to existing group
- Or choose a different mail nickname

**"Access denied" errors**
- Ensure you have Groups Administrator or Global Admin role
- Check Graph API permissions if using app-only auth

**Members not showing in Outlook immediately**
- Exchange Online can take 15-30 minutes to sync
- Use `Get-UnifiedGroupLinks` in Exchange Online PowerShell to verify

## Related Documentation

- [Microsoft Graph Groups API](https://docs.microsoft.com/en-us/graph/api/resources/group)
- [Manage Microsoft 365 Groups with PowerShell](https://docs.microsoft.com/en-us/microsoft-365/enterprise/manage-microsoft-365-groups-with-powershell)