# Microsoft Graph Group Management Scripts

This directory contains scripts for managing Azure AD groups through the Microsoft Graph API.

## Scripts

### Group Membership Management
- `add_members_sg.ps1` - Add members to security groups with comprehensive reporting

## Usage Examples

### Adding Group Members
```powershell
# Add single user to group
.\add_members_sg.ps1 -GroupName "Finance-Team" -UserEmail "user@domain.com"

# Add multiple users from CSV
.\add_members_sg.ps1 -GroupName "Finance-Team" -InputFile "users.csv"
```

## Features

- **Comprehensive error handling** with detailed reporting
- **CSV output** with operation results
- **Batch processing** for multiple users
- **Validation** of group and user existence
- **Progress reporting** for large operations

## Authentication

Uses Microsoft Graph API with certificate-based authentication configured in the main `Config.ps1` file.

## Output

All scripts generate detailed CSV reports including:
- Operation success/failure status
- User and group information
- Error details and recommendations
- Timestamp information

Reports are saved to `/output/mggraph/group_reports/` directory.