# SharePoint Permission Management Scripts

This directory contains scripts for managing SharePoint permissions - adding, removing, and modifying user access to SharePoint sites.

## Scripts

### Permission Removal Scripts
- `remove_permissions_parallel_pnp.ps1` - **[BULLETPROOF]** Fast parallel permission removal with safety prompts
- `mass_remove_sharepoint_permissions.ps1` - Mass removal of permissions across multiple sites
- `remove_permissions_pnp.ps1` - Standard permission removal script
- `remove_all_sharepoint_permissions.ps1` - Remove all permissions for a user
- `remove_sharepoint_permissions_pnp.ps1` - SharePoint-specific permission removal
- `remove_user_permissions.ps1` - Basic user permission removal
- `remove_user_permissions_comprehensive.ps1` - Comprehensive user permission removal
- `remove_user_permissions_optimized.ps1` - Optimized version for large tenants
- `remove_permissions_with_confirmation.ps1` - Permission removal with confirmation prompts
- `remove_via_groups_fast.ps1` - Fast removal via group membership
- `remove_known_sites.ps1` - Remove permissions from known site lists

### Audit and Cleanup Scripts
- `audit_and_remove_with_confirmation.ps1` - Audit permissions and remove with user confirmation

## Usage Examples

### Safe Permission Removal (Recommended)
```powershell
# Scan for permissions and prompt for removal approval
.\remove_permissions_parallel_pnp.ps1 -UserEmail "user@domain.com"

# Include subsites in scan (slower but comprehensive)
.\remove_permissions_parallel_pnp.ps1 -UserEmail "user@domain.com" -IncludeSubsites
```

### Mass Removal
```powershell
# Remove from multiple users
.\mass_remove_sharepoint_permissions.ps1 -UserList @("user1@domain.com", "user2@domain.com")
```

## Safety Features

All permission removal scripts include:
- **Confirmation prompts** before making changes
- **Comprehensive logging** of all actions
- **Error handling** and rollback capabilities  
- **Progress reporting** for long-running operations
- **CSV export** of findings for audit trails

## Authentication

These scripts use certificate-based authentication configured in the main `Config.ps1` file.