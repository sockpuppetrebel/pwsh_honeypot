# Microsoft Graph User Management Scripts

This directory contains scripts for managing users through the Microsoft Graph API.

## Scripts

### UPN Discovery and Management
- `find_missing_upns.ps1` - Find and resolve missing User Principal Names
- `getUPN_batch_list.ps1` - Batch processing for UPN lookups and validation

## Usage Examples

### UPN Discovery
```powershell
# Find missing UPNs for a list of users
.\find_missing_upns.ps1 -InputFile "users.csv"

# Batch UPN lookup and validation
.\getUPN_batch_list.ps1 -UserList @("user1", "user2", "user3")
```

## Authentication

These scripts use Microsoft Graph API with certificate-based authentication configured in the main `Config.ps1` file.

## Output

Scripts generate CSV reports with:
- User information and UPN status
- Validation results
- Error details for failed lookups
- Timestamp information for audit trails

Reports are saved to `/output/mggraph/` directory.