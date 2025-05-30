# SharePoint Utilities

This directory contains utility scripts for SharePoint authentication, testing, and user management.

## Scripts

### Authentication & Testing
- `test_pnp_connection.ps1` - Test PnP PowerShell connection and authentication
- `test_auth.ps1` - Test Microsoft Graph authentication
- `test_specific_sites.ps1` - Test connectivity to specific SharePoint sites

### Certificate Management
- `generate_cert.ps1` - Generate new certificates for app authentication
- `recreate_pem_certs.ps1` - Recreate PEM certificates from existing certificates
- `diagnose_cert.ps1` - Diagnose certificate issues and validation

### User Identity Management
- `find_all_user_identities.ps1` - **[COMPREHENSIVE]** Find all identity variations across SharePoint
- `find_user_by_objectid.ps1` - Find user permissions by Azure AD Object ID patterns
- `export_user_aad_groups.ps1` - Export user's Azure AD group memberships to CSV

### Diagnostics
- `diagnose_permissions.ps1` - Diagnose permission-related issues

## Key Utility Features

### Identity Management
Perfect for troubleshooting rehired employees or complex identity issues:
```powershell
# Find all possible identity variations
.\find_all_user_identities.ps1 -UserEmail "user@domain.com" -FirstName "John" -LastName "Doe"

# Search by Object ID patterns (finds old cached identities)
.\find_user_by_objectid.ps1 -UserEmail "user@domain.com"

# Export user's group memberships
.\export_user_aad_groups.ps1 -UserEmail "user@domain.com"
```

### Authentication Testing
```powershell
# Test PnP connection
.\test_pnp_connection.ps1

# Test Microsoft Graph authentication  
.\test_auth.ps1

# Test specific sites
.\test_specific_sites.ps1 -SiteUrls @("site1.sharepoint.com", "site2.sharepoint.com")
```

### Certificate Management
```powershell
# Generate new certificate
.\generate_cert.ps1 -CertName "MyAppAuth"

# Recreate certificates from existing
.\recreate_pem_certs.ps1
```

## Troubleshooting

These utilities are essential for:
- **Authentication issues** - Test connections and certificates
- **Identity conflicts** - Find old cached user profiles
- **Permission mysteries** - Comprehensive identity searching
- **Certificate problems** - Generate and validate certificates
- **Rehired employee issues** - Find conflicting identity remnants

## Authentication

All utilities use the centralized authentication configuration from the main `Config.ps1` file.