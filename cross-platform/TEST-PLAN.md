# Cross-Platform Testing Plan

Quick validation tests for verifying scripts work correctly on both Windows and macOS.

## Pre-Test Setup

**macOS:**
```bash
pwsh --version  # Verify PowerShell Core 7.x
```

**Windows:**
```powershell
$PSVersionTable.PSVersion  # Should show 7.x for cross-platform, 5.1 for Windows-only
```

## Core Functionality Tests

### 1. UPN Lookup Scripts
Test display name to UPN resolution.

**Quick test:**
```bash
# macOS
pwsh ./mggraph/user_management/Quick-UPNLookup.ps1

# Windows  
.\mggraph\user_management\Quick-UPNLookup.ps1
```

**Input:** `John Smith, Jane Doe`  
**Expected:** Should parse names and attempt Graph connection

### 2. Distribution List Management
Test membership replacement functionality.

**Safe test:**
```bash
# macOS
pwsh ./exchange/Set-DistributionListMembership.ps1 -WhatIf

# Windows
.\exchange\Set-DistributionListMembership.ps1 -WhatIf
```

**Expected:** Should prompt for DL name and show what would change without making changes

### 3. Group Membership Copy
Test M365 group member copying.

**Test command:**
```bash
# Both platforms
./m365/groups/Copy-M365GroupMembers.ps1 -WhatIf -SourceGroupEmail "test@domain.com" -NewGroupDisplayName "Test Group" -NewGroupMailNickname "testgroup"
```

**Expected:** Should validate source group and show planned copy operations

## Input/Output Tests

### Interactive Input
1. Run Quick-UPNLookup.ps1
2. Paste: `John Smith, Jane Doe; Bob Johnson`
3. Verify parses 3 names correctly

### File Operations
1. Check temp file creation works
2. Verify CSV export functions
3. Test log file generation

### Error Handling
1. Test with invalid group names
2. Test with non-existent users
3. Verify graceful failure messages

## Platform-Specific Validation

### Authentication
- **Graph connection** - Device code flow should work on both
- **Exchange connection** - Interactive auth should prompt correctly
- **Certificate auth** - Should handle platform differences

### File Paths
- **Temp directories** - Should use appropriate OS temp locations  
- **Output files** - Should create in correct locations
- **Path separators** - Should handle forward/back slashes correctly

## Performance Check

Run compatibility test:
```bash
pwsh ./scripts/utilities/Test-MacOSCompatibility.ps1
```

Should pass all basic PowerShell feature tests.

## Common Issues to Watch

**macOS specific:**
- Module installation permissions
- Keychain certificate access
- Case-sensitive file paths

**Windows specific:**  
- Execution policy blocks
- Module version conflicts
- Certificate store access

## Quick Validation Checklist

- [ ] Scripts start without syntax errors
- [ ] Interactive prompts display correctly
- [ ] Color output renders properly
- [ ] Progress bars work
- [ ] File operations complete
- [ ] Error messages are clear
- [ ] WhatIf mode functions correctly
- [ ] Authentication prompts appear
- [ ] CSV export works
- [ ] Log files generate

## Test Data

Use these safe test inputs:

**Distribution Lists:** `test-group@optimizely.com`  
**User Names:** `John Smith, Jane Doe, Test User`  
**UPNs:** `first.last@optimizely.com, test.user@optimizely.com`

Always use `-WhatIf` flag for destructive operations during testing.