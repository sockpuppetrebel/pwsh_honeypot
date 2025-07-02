# Getting Started with pwsh_honeypot

This guide will help you get up and running with the pwsh_honeypot repository and PowerShell scripting.

## Prerequisites

### 1. Install Git
- **Windows**: Download from [git-scm.com](https://git-scm.com/download/win)
- **macOS**: Install via Homebrew: `brew install git` or download from git-scm.com
- **Linux**: Use your package manager: `sudo apt install git` or `sudo yum install git`

### 2. Install PowerShell
- **Windows**: PowerShell 5.1 is built-in, but install PowerShell 7+ for best experience
- **macOS/Linux**: Install PowerShell 7+ from [Microsoft docs](https://docs.microsoft.com/en-us/powershell/scripting/install/installing-powershell)

```bash
# macOS via Homebrew
brew install --cask powershell

# Ubuntu/Debian
sudo apt update && sudo apt install -y powershell

# CentOS/RHEL
sudo yum install -y powershell
```

### 3. Install Visual Studio Code (Recommended)
Download from [code.visualstudio.com](https://code.visualstudio.com/)

Install the PowerShell extension:
1. Open VS Code
2. Go to Extensions (Ctrl+Shift+X)
3. Search for "PowerShell" 
4. Install the official Microsoft PowerShell extension

## Repository Setup

### 1. Clone the Repository
```bash
# Navigate to your preferred projects directory
cd ~/Projects  # or wherever you keep projects

# Clone the repository
git clone https://github.com/YOUR_USERNAME/pwsh_honeypot.git

# Navigate into the repository
cd pwsh_honeypot
```

### 2. Explore the Repository Structure
```bash
# List the main directories
ls -la

# Key directories to understand:
tree scripts/  # Shows the script organization
```

## Repository Structure Overview

```
pwsh_honeypot/
├── CLAUDE.md              # Development guidelines and standards
├── GETTING_STARTED.md     # This file
├── README.md              # Project overview
├── scripts/               # Main script directory
│   ├── azure/            # Azure-specific scripts
│   ├── exchange/         # Exchange Online scripts
│   ├── hr_lifecycle/     # User onboarding/offboarding
│   │   ├── onboarding/   # New user scripts
│   │   └── offboarding/  # User termination scripts
│   ├── m365/             # Microsoft 365 scripts
│   ├── mggraph/          # Microsoft Graph API scripts
│   ├── sharepoint/       # SharePoint scripts
│   └── utilities/        # General utility scripts
└── Azure/                # Azure deployment scripts
```

## Understanding the Codebase

### 1. Read the Development Guidelines
**Important**: Read `CLAUDE.md` first - it contains all coding standards and project conventions.

Key points:
- No emojis in code or documentation
- Professional, clean coding style
- Proper PowerShell conventions
- Security best practices

### 2. Explore Key Script Categories

#### Exchange Scripts (`scripts/exchange/`)
- Email management and automation
- Distribution list operations
- Mailbox configuration

#### HR Lifecycle (`scripts/hr_lifecycle/`)
- **Onboarding**: New user creation and setup
- **Offboarding**: User termination and cleanup
- Comprehensive automation for user lifecycle

#### Microsoft Graph (`scripts/mggraph/`)
- Azure AD user management
- Group operations
- Modern authentication patterns

#### SharePoint (`scripts/sharepoint/`)
- Site permissions and management
- OneDrive administration
- User access auditing

### 3. Study Script Patterns

Look at these example scripts to understand the coding patterns:

```powershell
# 1. Basic script structure
scripts/exchange/Add-DistributionGroupMember-Quick.ps1

# 2. Complex automation
scripts/hr_lifecycle/offboarding/Complete-UserOffboarding.ps1

# 3. User input handling
scripts/mggraph/user_management/Quick-UPNLookup.ps1
```

## PowerShell Fundamentals

### 1. Basic Commands to Learn
```powershell
# Get help for any command
Get-Help Get-Process -Examples

# List available commands
Get-Command *User*

# Explore object properties
Get-Process | Get-Member

# Filter and select data
Get-Process | Where-Object {$_.CPU -gt 100} | Select-Object Name, CPU

# Work with modules
Get-Module -ListAvailable
Import-Module Microsoft.Graph.Users
```

### 2. Essential PowerShell Concepts

#### Variables and Data Types
```powershell
# Variables
$name = "John Doe"
$numbers = @(1, 2, 3, 4, 5)
$hash = @{Name="John"; Age=30}

# Arrays and hashtables
$users = @()
$users += "user1@company.com"
$userInfo = @{Email="user@company.com"; Department="IT"}
```

#### Error Handling
```powershell
try {
    Get-Mailbox -Identity "nonexistent@company.com" -ErrorAction Stop
}
catch {
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
}
finally {
    # Cleanup code
}
```

#### Functions and Parameters
```powershell
function Get-UserInfo {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$UserPrincipalName,
        
        [Parameter(Mandatory=$false)]
        [switch]$IncludeGroups
    )
    
    # Function logic here
}
```

### 3. Microsoft 365 PowerShell Modules

#### Essential Modules to Install
```powershell
# Microsoft Graph (modern, cross-platform)
Install-Module Microsoft.Graph -Scope CurrentUser

# Exchange Online
Install-Module ExchangeOnlineManagement -Scope CurrentUser

# SharePoint (Windows only)
Install-Module Microsoft.Online.SharePoint.PowerShell -Scope CurrentUser

# Azure AD (being deprecated, use Graph instead)
Install-Module AzureAD -Scope CurrentUser
```

#### Common Connection Patterns
```powershell
# Microsoft Graph
Connect-MgGraph -Scopes "User.Read.All", "Group.ReadWrite.All"

# Exchange Online
Connect-ExchangeOnline

# SharePoint Online
Connect-SPOService -Url "https://tenant-admin.sharepoint.com"
```

## Development Workflow

### 1. Create a New Script

#### Choose the Right Location
- Exchange operations → `scripts/exchange/`
- User management → `scripts/mggraph/user_management/`
- Group management → `scripts/mggraph/group_management/`
- HR processes → `scripts/hr_lifecycle/`

#### Follow the Template Structure
```powershell
<#
.SYNOPSIS
    Brief description of what the script does

.DESCRIPTION
    Detailed description including:
    - Main functionality
    - Prerequisites
    - What it modifies

.PARAMETER ParameterName
    Description of the parameter

.EXAMPLE
    .\Script-Name.ps1 -Parameter "value"
    
    Description of what this example does

.NOTES
    Author: Your Name
    Date: YYYY-MM-DD
    Version: 1.0
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory=$true)]
    [string]$RequiredParameter,
    
    [Parameter(Mandatory=$false)]
    [switch]$OptionalSwitch
)

# Script logic here
```

### 2. Testing Your Scripts

#### Use WhatIf for Safe Testing
```powershell
# Test mode - shows what would happen without making changes
.\Your-Script.ps1 -WhatIf

# Verbose output for debugging
.\Your-Script.ps1 -Verbose
```

#### Test with Limited Scope
```powershell
# Test with a single user first
.\Remove-UserFromGroups.ps1 -UserPrincipalName "test@company.com" -WhatIf

# Then expand to larger operations
```

### 3. Version Control Workflow

#### Making Changes
```bash
# Check current status
git status

# Stage your changes
git add scripts/your-new-script.ps1

# Commit with descriptive message
git commit -m "Add script for automated mailbox delegation cleanup

- Removes Full Access and Send As permissions
- Includes backup and restore functionality
- Supports WhatIf for safe testing"

# Push to GitHub
git push origin main
```

## Common Scripting Patterns

### 1. User Input Handling
```powershell
# Simple input
$userEmail = Read-Host "Enter user email"

# Multi-line input (for lists)
Write-Host "Enter email addresses (press Enter twice when done):"
$emails = @()
do {
    $line = Read-Host
    if ($line.Trim() -ne "") {
        $emails += $line.Trim()
    }
} while ($line.Trim() -ne "")
```

### 2. Bulk Operations
```powershell
# Process multiple users
$users = @("user1@company.com", "user2@company.com")
foreach ($user in $users) {
    try {
        # Perform operation
        Write-Host "Processing $user..." -ForegroundColor Green
    }
    catch {
        Write-Host "Failed to process $user: $_" -ForegroundColor Red
    }
}
```

### 3. Progress Tracking
```powershell
$users = Get-Content "userlist.txt"
$total = $users.Count
$current = 0

foreach ($user in $users) {
    $current++
    $percent = [math]::Round(($current / $total) * 100)
    Write-Progress -Activity "Processing Users" -Status "$current of $total" -PercentComplete $percent
    
    # Do work here
}
```

## Security Best Practices

### 1. Never Hardcode Credentials
```powershell
# DON'T do this
$password = "MyPassword123"

# DO this instead
$credential = Get-Credential
# Or use certificate-based authentication
```

### 2. Validate Input
```powershell
param(
    [Parameter(Mandatory=$true)]
    [ValidatePattern('^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$')]
    [string]$Email
)
```

### 3. Use Least Privilege
```powershell
# Only request the permissions you need
Connect-MgGraph -Scopes "User.Read.All"  # Not "Directory.ReadWrite.All"
```

## Useful Resources

### Documentation
- [PowerShell Documentation](https://docs.microsoft.com/en-us/powershell/)
- [Microsoft Graph PowerShell](https://docs.microsoft.com/en-us/powershell/microsoftgraph/)
- [Exchange Online PowerShell](https://docs.microsoft.com/en-us/powershell/exchange/)

### Learning Resources
- [PowerShell in a Month of Lunches](https://www.manning.com/books/learn-powershell-in-a-month-of-lunches) (book)
- [Microsoft Learn PowerShell Path](https://docs.microsoft.com/en-us/learn/paths/powershell/)

### Tools
- [PowerShell ISE](https://docs.microsoft.com/en-us/powershell/scripting/windows-powershell/ise/introducing-the-windows-powershell-ise) (Windows)
- [Visual Studio Code](https://code.visualstudio.com/) (Cross-platform)
- [PowerShell extension for VS Code](https://marketplace.visualstudio.com/items?itemName=ms-vscode.PowerShell)

## Getting Help

### In the Repository
1. Check existing scripts for similar functionality
2. Read `CLAUDE.md` for coding standards
3. Look at script headers for usage examples

### PowerShell Help
```powershell
# Get help for any command
Get-Help Connect-MgGraph -Examples
Get-Help about_Functions
Get-Help about_ErrorHandling

# Update help files
Update-Help
```

### Online Communities
- [PowerShell subreddit](https://www.reddit.com/r/PowerShell/)
- [PowerShell Discord](https://discord.gg/PowerShell)
- [Stack Overflow PowerShell tag](https://stackoverflow.com/questions/tagged/powershell)

## Next Steps

1. **Set up your development environment** (Git, PowerShell, VS Code)
2. **Clone the repository** and explore the structure
3. **Read `CLAUDE.md`** to understand project standards
4. **Install required PowerShell modules** for your use cases
5. **Study existing scripts** that are similar to what you want to build
6. **Start with simple modifications** to existing scripts
7. **Create your first original script** following the established patterns
8. **Test thoroughly** with `-WhatIf` before running in production

Remember: Always test scripts in a safe environment before running them against production data!