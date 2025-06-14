#Requires -Version 7.0
#Requires -Modules Microsoft.Graph.Authentication, Microsoft.Graph.Users, Microsoft.Graph.Groups

<#
.SYNOPSIS
    Creates test users in Azure AD for IT-Bot testing
.DESCRIPTION
    This script creates a set of test users with different roles and permissions
    for testing password reset and other IT-Bot functionalities.
.PARAMETER TenantId
    The Azure AD tenant ID
.PARAMETER NumberOfUsers
    Number of test users to create (default: 5)
.PARAMETER TestDomain
    Domain suffix for test users (default: your tenant domain)
.EXAMPLE
    .\create-test-users.ps1 -TenantId "your-tenant-id"
#>

param(
    [Parameter(Mandatory = $true)]
    [string]$TenantId,
    
    [Parameter(Mandatory = $false)]
    [int]$NumberOfUsers = 5,
    
    [Parameter(Mandatory = $false)]
    [string]$TestDomain = $null
)

# Test user templates
$TestUserTemplates = @(
    @{
        DisplayName = "Test User Standard"
        GivenName = "Test"
        Surname = "Standard"
        JobTitle = "Test Standard User"
        Department = "Testing"
        UsageLocation = "US"
    },
    @{
        DisplayName = "Test User Manager"
        GivenName = "Test"
        Surname = "Manager"
        JobTitle = "Test Manager"
        Department = "Testing"
        UsageLocation = "US"
    },
    @{
        DisplayName = "Test User Admin"
        GivenName = "Test"
        Surname = "Admin"
        JobTitle = "Test Administrator"
        Department = "IT"
        UsageLocation = "US"
    },
    @{
        DisplayName = "Test User Helpdesk"
        GivenName = "Test"
        Surname = "Helpdesk"
        JobTitle = "Test Helpdesk Agent"
        Department = "IT Support"
        UsageLocation = "US"
    },
    @{
        DisplayName = "Test User VPN"
        GivenName = "Test"
        Surname = "VPN"
        JobTitle = "Test VPN User"
        Department = "Remote Access"
        UsageLocation = "US"
    }
)

function Write-Log {
    param($Message, $Level = "INFO")
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Write-Host "[$timestamp] [$Level] $Message"
}

function New-RandomPassword {
    param([int]$Length = 12)
    
    $chars = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789!@#$%^&*"
    $password = ""
    for ($i = 0; $i -lt $Length; $i++) {
        $password += $chars[(Get-Random -Maximum $chars.Length)]
    }
    return $password
}

function Connect-ToMicrosoftGraph {
    param([string]$TenantId)
    
    try {
        Write-Log "Connecting to Microsoft Graph..."
        Connect-MgGraph -TenantId $TenantId -Scopes "User.ReadWrite.All", "Group.ReadWrite.All", "Directory.ReadWrite.All"
        
        $context = Get-MgContext
        Write-Log "Connected successfully to tenant: $($context.TenantId)"
        return $true
    }
    catch {
        Write-Log "Failed to connect to Microsoft Graph: $($_.Exception.Message)" "ERROR"
        return $false
    }
}

function Get-TenantDomain {
    try {
        $domains = Get-MgDomain | Where-Object { $_.IsDefault -eq $true }
        return $domains[0].Id
    }
    catch {
        Write-Log "Failed to get tenant domain: $($_.Exception.Message)" "ERROR"
        return $null
    }
}

function New-TestUser {
    param(
        [hashtable]$UserTemplate,
        [string]$Domain,
        [int]$Index
    )
    
    try {
        $userPrincipalName = "testuser$Index@$Domain"
        $mailNickname = "testuser$Index"
        $password = New-RandomPassword
        
        $passwordProfile = @{
            ForceChangePasswordNextSignIn = $false
            Password = $password
        }
        
        $userParams = @{
            DisplayName = "$($UserTemplate.DisplayName) $Index"
            GivenName = $UserTemplate.GivenName
            Surname = "$($UserTemplate.Surname)$Index"
            UserPrincipalName = $userPrincipalName
            MailNickname = $mailNickname
            JobTitle = $UserTemplate.JobTitle
            Department = $UserTemplate.Department
            UsageLocation = $UserTemplate.UsageLocation
            PasswordProfile = $passwordProfile
            AccountEnabled = $true
        }
        
        Write-Log "Creating user: $userPrincipalName"
        $user = New-MgUser @userParams
        
        Write-Log "Created user successfully: $($user.DisplayName) ($($user.UserPrincipalName))"
        
        return @{
            User = $user
            TempPassword = $password
        }
    }
    catch {
        Write-Log "Failed to create user: $($_.Exception.Message)" "ERROR"
        return $null
    }
}

function New-TestGroup {
    param([string]$GroupName, [string]$Description)
    
    try {
        $groupParams = @{
            DisplayName = $GroupName
            Description = $Description
            GroupTypes = @()
            MailEnabled = $false
            MailNickname = $GroupName.Replace(" ", "").ToLower()
            SecurityEnabled = $true
        }
        
        Write-Log "Creating group: $GroupName"
        $group = New-MgGroup @groupParams
        
        Write-Log "Created group successfully: $($group.DisplayName)"
        return $group
    }
    catch {
        Write-Log "Failed to create group: $($_.Exception.Message)" "ERROR"
        return $null
    }
}

# Main execution
Write-Log "Starting IT-Bot test user creation script"

# Connect to Microsoft Graph
if (-not (Connect-ToMicrosoftGraph -TenantId $TenantId)) {
    Write-Log "Cannot proceed without Microsoft Graph connection" "ERROR"
    exit 1
}

# Get tenant domain if not provided
if (-not $TestDomain) {
    $TestDomain = Get-TenantDomain
    if (-not $TestDomain) {
        Write-Log "Could not determine tenant domain" "ERROR"
        exit 1
    }
}

Write-Log "Using domain: $TestDomain"

# Create test groups
$testGroups = @(
    @{ Name = "IT-Bot Test Users"; Description = "Group for IT-Bot testing users" },
    @{ Name = "IT-Bot VPN Users"; Description = "Group for VPN access testing" },
    @{ Name = "IT-Bot Helpdesk"; Description = "Group for helpdesk role testing" }
)

$createdGroups = @{}
foreach ($groupTemplate in $testGroups) {
    $group = New-TestGroup -GroupName $groupTemplate.Name -Description $groupTemplate.Description
    if ($group) {
        $createdGroups[$groupTemplate.Name] = $group
    }
}

# Create test users
$createdUsers = @()
for ($i = 1; $i -le $NumberOfUsers; $i++) {
    $templateIndex = ($i - 1) % $TestUserTemplates.Count
    $template = $TestUserTemplates[$templateIndex]
    
    $userResult = New-TestUser -UserTemplate $template -Domain $TestDomain -Index $i
    if ($userResult) {
        $createdUsers += $userResult
        
        # Add user to test group
        if ($createdGroups["IT-Bot Test Users"]) {
            try {
                New-MgGroupMember -GroupId $createdGroups["IT-Bot Test Users"].Id -DirectoryObjectId $userResult.User.Id
                Write-Log "Added user to IT-Bot Test Users group"
            }
            catch {
                Write-Log "Failed to add user to group: $($_.Exception.Message)" "WARNING"
            }
        }
        
        # Add specific users to special groups
        if ($template.Department -eq "IT Support" -and $createdGroups["IT-Bot Helpdesk"]) {
            try {
                New-MgGroupMember -GroupId $createdGroups["IT-Bot Helpdesk"].Id -DirectoryObjectId $userResult.User.Id
                Write-Log "Added user to IT-Bot Helpdesk group"
            }
            catch {
                Write-Log "Failed to add user to helpdesk group: $($_.Exception.Message)" "WARNING"
            }
        }
        
        if ($template.Department -eq "Remote Access" -and $createdGroups["IT-Bot VPN Users"]) {
            try {
                New-MgGroupMember -GroupId $createdGroups["IT-Bot VPN Users"].Id -DirectoryObjectId $userResult.User.Id
                Write-Log "Added user to IT-Bot VPN Users group"
            }
            catch {
                Write-Log "Failed to add user to VPN group: $($_.Exception.Message)" "WARNING"
            }
        }
    }
    
    Start-Sleep -Seconds 2
}

# Generate summary report
$reportPath = "test-users-report-$(Get-Date -Format 'yyyyMMdd-HHmmss').csv"
$createdUsers | ForEach-Object {
    [PSCustomObject]@{
        DisplayName = $_.User.DisplayName
        UserPrincipalName = $_.User.UserPrincipalName
        JobTitle = $_.User.JobTitle
        Department = $_.User.Department
        TempPassword = $_.TempPassword
        ObjectId = $_.User.Id
    }
} | Export-Csv -Path $reportPath -NoTypeInformation

Write-Log "Test user creation completed!"
Write-Log "Created $($createdUsers.Count) users"
Write-Log "Report saved to: $reportPath"
Write-Log ""
Write-Log "IMPORTANT: Save the temporary passwords from the report file!"
Write-Log "Users will need to change their passwords on first login."

# Disconnect from Microsoft Graph
Disconnect-MgGraph
Write-Log "Disconnected from Microsoft Graph"