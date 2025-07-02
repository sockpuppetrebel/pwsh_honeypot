# HR Lifecycle Scripts

This directory contains PowerShell scripts for automating user onboarding and offboarding processes in Microsoft 365.

## Directory Structure

```
hr_lifecycle/
├── onboarding/
│   └── New-UserOnboarding.ps1
└── offboarding/
    ├── Remove-AllAADGroupMemberships.ps1
    ├── Complete-UserOffboarding.ps1
    ├── Remove-MailboxDelegations.ps1
    └── Export-UserAccessReport.ps1
```

## Scripts Overview

### Onboarding Scripts

#### New-UserOnboarding.ps1
Automates the complete new user setup process:
- Creates Azure AD user account
- Generates secure temporary password
- Assigns licenses
- Creates mailbox
- Adds to appropriate groups
- Sets up manager relationship
- Configures mailbox settings
- Sends welcome email with credentials

**Usage:**
```powershell
.\New-UserOnboarding.ps1 -FirstName "John" -LastName "Doe" -JobTitle "Software Engineer" -Department "IT" -Manager "manager@company.com" -PersonalEmail "john.doe@personal.com"
```

### Offboarding Scripts

#### Remove-AllAADGroupMemberships.ps1
Removes user from all Azure AD groups during offboarding:
- Backs up current group memberships
- Removes from all groups (excluding dynamic groups)
- Option to exclude critical groups
- Full audit logging

**Usage:**
```powershell
.\Remove-AllAADGroupMemberships.ps1 -UserPrincipalName "john.doe@company.com"
```

#### Complete-UserOffboarding.ps1
Comprehensive offboarding checklist with automated actions:
- Disables user account
- Resets password
- Blocks sign-in and revokes sessions
- Sets out-of-office message
- Forwards email to manager
- Removes from all groups
- Removes licenses
- Converts mailbox to shared
- Hides from GAL
- Complete audit trail

**Usage:**
```powershell
.\Complete-UserOffboarding.ps1 -UserPrincipalName "john.doe@company.com" -ManagerEmail "manager@company.com" -TicketNumber "HR-2024-001"
```

#### Remove-MailboxDelegations.ps1
Removes all mailbox permissions and delegations:
- Removes Full Access permissions
- Removes Send As permissions
- Removes Send on Behalf permissions
- Removes calendar delegations
- Option to remove permissions in both directions
- Exports current permissions before removal

**Usage:**
```powershell
.\Remove-MailboxDelegations.ps1 -UserPrincipalName "john.doe@company.com"

# Export only (no removal)
.\Remove-MailboxDelegations.ps1 -UserPrincipalName "john.doe@company.com" -ExportOnly
```

#### Export-UserAccessReport.ps1
Generates comprehensive access report before offboarding:
- User profile information
- Group memberships
- Assigned licenses
- Application assignments
- Mailbox permissions
- Security status (MFA, last sign-in)
- Multiple output formats (HTML, CSV, JSON)

**Usage:**
```powershell
.\Export-UserAccessReport.ps1 -UserPrincipalName "john.doe@company.com"

# With SharePoint access and CSV output
.\Export-UserAccessReport.ps1 -UserPrincipalName "john.doe@company.com" -OutputFormat CSV -IncludeSharePoint
```

## Prerequisites

1. **PowerShell Modules Required:**
   - Microsoft.Graph
   - ExchangeOnlineManagement

2. **Permissions Required:**
   - User.ReadWrite.All
   - Group.ReadWrite.All
   - Directory.ReadWrite.All
   - Mail.Send
   - Exchange Administrator role (for mailbox operations)

3. **Installation:**
   ```powershell
   Install-Module Microsoft.Graph -Scope CurrentUser
   Install-Module ExchangeOnlineManagement -Scope CurrentUser
   ```

## Common Parameters

Most scripts support these common parameters:
- `-WhatIf`: Preview changes without making them
- `-Verbose`: Detailed output
- `-TicketNumber`: HR ticket reference for audit trail

## Logging

All scripts create detailed logs in their respective `Logs` subdirectories:
- Transcript logs for full session recording
- JSON exports for structured data
- HTML reports for easy viewing
- Backup files before making changes

## Best Practices

1. **Always run with -WhatIf first** to preview changes
2. **Review the generated reports** before proceeding
3. **Keep ticket numbers** for audit trail
4. **Test in a non-production environment** first
5. **Ensure proper backups** exist before offboarding

## Security Considerations

- Temporary passwords are generated securely
- All actions are logged for audit purposes
- Sensitive data is not displayed in logs
- Scripts require appropriate admin permissions
- MFA should be enforced for admin accounts

## Support

For issues or enhancements, please contact the IT team or submit a ticket.