# SharePoint Auditing Scripts

This directory contains scripts for auditing and reporting on SharePoint permissions and site access.

## Scripts

### Permission Auditing
- `check_permissions_pnp.ps1` - Check user permissions across SharePoint sites
- `search_user_permissions_pnp.ps1` - Search for specific user permissions
- `smart_search_pnp.ps1` - Intelligent permission search with filters
- `fast_scan_pnp.ps1` - Fast scanning for permission audits

### Site Auditing  
- `check_sites_v2.ps1` - Enhanced site permission checking
- `check_specific_sites.ps1` - Check permissions on specific site lists

## Usage Examples

### Basic Permission Audit
```powershell
# Check permissions for a user across all sites
.\check_permissions_pnp.ps1 -UserEmail "user@domain.com"

# Fast scan mode for quick overview
.\fast_scan_pnp.ps1 -UserEmail "user@domain.com"
```

### Advanced Searching
```powershell
# Smart search with filtering
.\smart_search_pnp.ps1 -UserEmail "user@domain.com" -SiteFilter "Finance"

# Search specific sites only
.\check_specific_sites.ps1 -UserEmail "user@domain.com" -SiteList @("site1", "site2")
```

## Output

All auditing scripts generate:
- **CSV reports** with detailed permission information
- **Console output** with real-time progress
- **Summary statistics** of findings
- **Timestamped logs** for audit trails

Reports are automatically saved to `/output/sharepoint/audit_reports/`

## Performance

- **Parallel processing** for faster execution on large tenants
- **Progress indicators** for long-running operations  
- **Throttling controls** to prevent API limits
- **Batch processing** for optimal performance