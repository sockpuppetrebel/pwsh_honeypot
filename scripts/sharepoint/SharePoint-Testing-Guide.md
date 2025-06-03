# SharePoint User Permission Fix - Testing Guide

## Overview
This guide provides a safe testing approach for fixing SharePoint user ID mismatches in an enterprise environment.

## Scripts Provided

1. **Fix-SharePointUserMismatch-SAFE.ps1** - Main script with safety features
2. **Test-SharePointUserCheck.ps1** - Permission verification script
3. **Fix-SharePointUserMismatch.ps1** - Production script (use after testing)

## Recommended Testing Process

### Phase 1: Initial Assessment
```powershell
# 1. Check current user permissions (non-destructive)
.\Test-SharePointUserCheck.ps1 -AdminUrl "https://episerver99-admin.sharepoint.com/" -UserToCheck "kaila.trapani@optimizely.com"

# This will:
# - Verify user exists in tenant
# - Check sample of sites for permissions
# - Export results to CSV
# - Show summary of access
```

### Phase 2: WhatIf Simulation
```powershell
# 2. Run full simulation without making changes
.\Fix-SharePointUserMismatch-SAFE.ps1 -AdminUrl "https://episerver99-admin.sharepoint.com/" -UserToRemove "kaila.trapani@optimizely.com" -WhatIf

# This will:
# - Connect to SharePoint
# - Enumerate ALL sites
# - Show what WOULD be removed
# - Generate detailed report
# - Make NO actual changes
```

### Phase 3: Limited Test Run
```powershell
# 3. Test on specific sites first
$testSites = @(
    "https://episerver99.sharepoint.com/sites/TestSite1",
    "https://episerver99-my.sharepoint.com/personal/test_user_optimizely_com"
)

.\Fix-SharePointUserMismatch-SAFE.ps1 -AdminUrl "https://episerver99-admin.sharepoint.com/" -UserToRemove "kaila.trapani@optimizely.com" -TestSiteUrls $testSites -Force -WhatIf:$false

# Or limit to first 5 sites
.\Fix-SharePointUserMismatch-SAFE.ps1 -AdminUrl "https://episerver99-admin.sharepoint.com/" -UserToRemove "kaila.trapani@optimizely.com" -LimitSites 5 -Force -WhatIf:$false
```

### Phase 4: Verify Changes
```powershell
# 4. Check if removal worked on test sites
.\Test-SharePointUserCheck.ps1 -AdminUrl "https://episerver99-admin.sharepoint.com/" -UserToCheck "kaila.trapani@optimizely.com" -SiteUrls $testSites
```

### Phase 5: Full Production Run
```powershell
# 5. Only after successful testing
.\Fix-SharePointUserMismatch-SAFE.ps1 -AdminUrl "https://episerver99-admin.sharepoint.com/" -UserToRemove "kaila.trapani@optimizely.com" -Force -WhatIf:$false

# Will prompt for final confirmation
```

## Safety Features

1. **Default WhatIf Mode** - Script defaults to simulation mode
2. **Force Parameter Required** - Must explicitly use -Force for live changes
3. **Pre-flight Checks** - Validates user exists before proceeding
4. **Detailed Logging** - All actions logged with timestamps
5. **Progress Tracking** - Visual progress during execution
6. **Error Handling** - Distinguishes between real errors and expected conditions

## Important Considerations

### Before Running:
- **Backup** - No direct backup needed as we're only removing permissions, not data
- **Communication** - Inform affected user (Kaila) that temporary access loss may occur
- **Timing** - Run during low-usage hours if possible
- **Test First** - Always run WhatIf mode first

### After Running:
- **Sync Time** - SharePoint can take up to 24 hours to fully sync changes
- **Cache** - User should clear browser cache/cookies
- **Verification** - Use Test-SharePointUserCheck.ps1 to verify
- **Re-sharing** - Sites may need to be re-shared with correct permissions

## Rollback Process

If issues occur after running the script:

1. **No Direct Rollback** - Removing users doesn't delete data
2. **Re-add User** - Use SharePoint admin center to re-add user to specific sites
3. **Restore Permissions** - Use site settings to restore specific permissions
4. **Alternative** - Use PowerShell to add user back:
   ```powershell
   Add-SPOUser -Site "https://site-url" -LoginName "kaila.trapani@optimizely.com" -Group "Members"
   ```

## Monitoring

Check these after running:
1. Review log files for any errors
2. Verify user can't access removed sites
3. Confirm user CAN access newly shared sites
4. Monitor for any user complaints about access

## Common Issues

1. **"User still has access"**
   - Clear browser cache
   - Wait for sync (up to 24 hours)
   - Check for direct sharing links

2. **"Script runs too fast"**
   - This means user wasn't found on most sites (normal)
   - Check logs for actual removals

3. **"Access denied errors"**
   - Ensure you have SharePoint Admin rights
   - Some sites may have unique permissions

## Support

- Keep all log files for troubleshooting
- Document which sites were affected
- Have SharePoint admin center access ready
- Microsoft Support can help with tenant-level issues