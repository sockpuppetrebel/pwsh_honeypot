# Cross-Platform Validation Checklist

Pre-deployment validation for ensuring scripts work reliably across environments.

## Environment Setup

**Prerequisites:**
- [ ] PowerShell Core 7.0+ installed
- [ ] Required modules available: `Microsoft.Graph.Users`, `ExchangeOnlineManagement`
- [ ] Network connectivity to Microsoft 365 services
- [ ] Appropriate permissions for tenant operations

## Script Categories

### Exchange Online Scripts

**Set-DistributionListMembership.ps1**
- [ ] Connects to Exchange Online successfully
- [ ] Prompts for distribution list correctly
- [ ] Parses member lists (comma, semicolon, newline separated)
- [ ] Shows current vs desired membership comparison
- [ ] WhatIf mode works without making changes
- [ ] Confirmation prompts function properly
- [ ] Error handling works for invalid groups/users

**Add-UserToMultipleGroups.ps1**
- [ ] Validates user exists before processing
- [ ] Handles different recipient types (DL, shared mailbox, user mailbox)
- [ ] Shows clear output for each addition type
- [ ] Skips existing memberships appropriately
- [ ] Provides detailed summary of operations

### Microsoft Graph User Management

**Quick-UPNLookup.ps1**
- [ ] Prompts for name input correctly
- [ ] Parses various input formats (names, commas, newlines)
- [ ] Connects to Microsoft Graph successfully
- [ ] Resolves display names to UPNs accurately
- [ ] Handles multiple matches appropriately
- [ ] Shows not found results clearly
- [ ] Offers file save option

**Get-UPNByDisplayName.ps1**
- [ ] Supports all parameter sets (single, multiple, file, raw text)
- [ ] RawText parameter parses input correctly
- [ ] Exact vs fuzzy matching works as expected
- [ ] CSV export functions properly
- [ ] Progress reporting displays correctly

### M365 Group Management

**Copy-M365GroupMembers.ps1**
- [ ] Validates source group exists
- [ ] Creates new group or uses existing as specified
- [ ] Copies owners before members
- [ ] Handles duplicate detection correctly
- [ ] Provides detailed operation summary
- [ ] WhatIf mode shows planned operations

## Platform-Specific Checks

### macOS Validation
- [ ] File paths use forward slashes appropriately
- [ ] Temp directory resolves correctly
- [ ] Color output displays properly in terminal
- [ ] Interactive prompts accept input correctly
- [ ] Module loading works without errors

### Windows Validation  
- [ ] Execution policy allows script execution
- [ ] File paths handle backslashes correctly
- [ ] PowerShell ISE compatibility (if applicable)
- [ ] Windows PowerShell vs PowerShell Core behavior consistent

## Security and Safety

**Authentication:**
- [ ] Device code authentication works
- [ ] Interactive authentication prompts correctly
- [ ] Certificate authentication (where applicable) functions
- [ ] Connection verification succeeds

**Data Protection:**
- [ ] No PII in script outputs or logs
- [ ] Error messages don't expose sensitive information
- [ ] Temporary files cleaned up properly

**Operation Safety:**
- [ ] WhatIf mode available for destructive operations
- [ ] Confirmation prompts appear for critical changes
- [ ] Error handling prevents partial operations
- [ ] Rollback information provided where applicable

## Performance and Reliability

**Execution:**
- [ ] Scripts complete within reasonable time
- [ ] Progress indicators function during long operations
- [ ] Memory usage remains reasonable
- [ ] Network timeouts handled gracefully

**Error Handling:**
- [ ] Network connectivity issues handled
- [ ] API throttling respected
- [ ] Invalid input handled gracefully
- [ ] Partial failures reported clearly

## Output Validation

**Display:**
- [ ] Color coding works correctly
- [ ] Progress bars display properly
- [ ] Summary information is accurate
- [ ] Error messages are clear and actionable

**Files:**
- [ ] CSV exports contain correct data
- [ ] Log files generated with timestamps
- [ ] Output directories created as needed
- [ ] File encoding consistent across platforms

## Final Verification

**Integration Testing:**
- [ ] Scripts work with real Microsoft 365 tenant
- [ ] Operations complete successfully end-to-end
- [ ] Results match expectations
- [ ] No unintended side effects observed

**Documentation:**
- [ ] Script help text accurate
- [ ] Examples in documentation work
- [ ] Parameter descriptions match functionality
- [ ] Cross-platform notes reflect actual behavior

## Sign-off

Tested by: ________________  
Date: ________________  
Environment: ________________  
PowerShell Version: ________________  
Notes: ________________