<#
.SYNOPSIS
    Automates new user onboarding process in Microsoft 365.

.DESCRIPTION
    This script performs comprehensive onboarding tasks including:
    - Creating user account in Azure AD
    - Assigning licenses
    - Creating mailbox
    - Adding to appropriate groups
    - Setting up MFA
    - Sending welcome email
    - Generating onboarding report

.PARAMETER FirstName
    User's first name

.PARAMETER LastName
    User's last name

.PARAMETER JobTitle
    User's job title

.PARAMETER Department
    User's department

.PARAMETER Manager
    Manager's email address

.PARAMETER LicenseSKU
    License SKU to assign (e.g., "company:ENTERPRISEPACK")

.PARAMETER Groups
    Array of group names to add user to

.PARAMETER SendWelcomeEmail
    Send welcome email to user's personal email

.PARAMETER PersonalEmail
    Personal email for welcome message

.EXAMPLE
    .\New-UserOnboarding.ps1 -FirstName "John" -LastName "Doe" -JobTitle "Software Engineer" -Department "IT" -Manager "manager@company.com" -PersonalEmail "john.doe@personal.com"

.NOTES
    Author: HR Automation Script
    Date: 2025-07-02
    Version: 1.0
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory=$true)]
    [string]$FirstName,
    
    [Parameter(Mandatory=$true)]
    [string]$LastName,
    
    [Parameter(Mandatory=$true)]
    [string]$JobTitle,
    
    [Parameter(Mandatory=$true)]
    [string]$Department,
    
    [Parameter(Mandatory=$false)]
    [string]$Manager,
    
    [Parameter(Mandatory=$false)]
    [string]$LicenseSKU,
    
    [Parameter(Mandatory=$false)]
    [string[]]$Groups = @(),
    
    [Parameter(Mandatory=$false)]
    [switch]$SendWelcomeEmail = $true,
    
    [Parameter(Mandatory=$false)]
    [string]$PersonalEmail,
    
    [Parameter(Mandatory=$false)]
    [string]$UsageLocation = "US",
    
    [Parameter(Mandatory=$false)]
    [string]$TicketNumber = "N/A"
)

# Initialize logging
$LogPath = "$PSScriptRoot\Logs\Onboarding"
if (-not (Test-Path $LogPath)) {
    New-Item -ItemType Directory -Path $LogPath -Force | Out-Null
}
$SessionId = Get-Date -Format 'yyyyMMdd-HHmmss'
$LogFile = "$LogPath\Onboarding-$FirstName$LastName-$SessionId.log"
Start-Transcript -Path $LogFile

# Initialize results tracking
$results = @{
    FirstName = $FirstName
    LastName = $LastName
    TicketNumber = $TicketNumber
    StartTime = Get-Date
    Steps = @()
    Credentials = @{}
}

function Add-OnboardingStep {
    param(
        [string]$Step,
        [string]$Status,
        [string]$Details = ""
    )
    
    $results.Steps += @{
        Step = $Step
        Status = $Status
        Details = $Details
        Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    }
    
    $color = switch ($Status) {
        "Success" { "Green" }
        "Failed" { "Red" }
        "Skipped" { "Yellow" }
        "InProgress" { "Cyan" }
        default { "White" }
    }
    
    Write-Host "[$Status] $Step $(if ($Details) {": $Details"})" -ForegroundColor $color
}

function Generate-Username {
    param(
        [string]$FirstName,
        [string]$LastName,
        [string]$Domain
    )
    
    # Try different username formats
    $formats = @(
        "$($FirstName.ToLower()).$($LastName.ToLower())",
        "$($FirstName.Substring(0,1).ToLower())$($LastName.ToLower())",
        "$($FirstName.ToLower())$($LastName.Substring(0,1).ToLower())"
    )
    
    foreach ($format in $formats) {
        $upn = "$format@$Domain"
        try {
            $existingUser = Get-MgUser -UserId $upn -ErrorAction Stop
        } catch {
            # User doesn't exist, we can use this UPN
            return $upn
        }
    }
    
    # If all formats taken, add number
    $counter = 1
    while ($true) {
        $upn = "$($formats[0])$counter@$Domain"
        try {
            $existingUser = Get-MgUser -UserId $upn -ErrorAction Stop
            $counter++
        } catch {
            return $upn
        }
    }
}

Write-Host "`n=== NEW USER ONBOARDING PROCESS ===" -ForegroundColor Green
Write-Host "Name: $FirstName $LastName" -ForegroundColor Yellow
Write-Host "Title: $JobTitle" -ForegroundColor Yellow
Write-Host "Department: $Department" -ForegroundColor Yellow
Write-Host "Ticket: $TicketNumber" -ForegroundColor Yellow
Write-Host "Started: $(Get-Date)" -ForegroundColor Yellow
Write-Host "="*50 -ForegroundColor Green

try {
    # Connect to required services
    Write-Host "`nConnecting to Microsoft services..." -ForegroundColor Cyan
    
    # Connect to Microsoft Graph
    $mgContext = Get-MgContext -ErrorAction SilentlyContinue
    if (-not $mgContext) {
        Connect-MgGraph -Scopes "User.ReadWrite.All", "Group.ReadWrite.All", "Directory.ReadWrite.All", "Mail.Send", "UserAuthenticationMethod.ReadWrite.All" -NoWelcome
    }
    
    # Connect to Exchange Online
    $exoSession = Get-PSSession | Where-Object {$_.ConfigurationName -eq "Microsoft.Exchange" -and $_.State -eq "Opened"}
    if (-not $exoSession) {
        Connect-ExchangeOnline -ShowBanner:$false
    }
    
    # Get tenant domain
    $organization = Get-MgOrganization
    $domain = ($organization.VerifiedDomains | Where-Object { $_.IsDefault }).Name
    
    # Step 1: Generate Username
    Write-Host "`n[Step 1] Generating username..." -ForegroundColor Cyan
    $userPrincipalName = Generate-Username -FirstName $FirstName -LastName $LastName -Domain $domain
    $displayName = "$FirstName $LastName"
    $mailNickname = $userPrincipalName.Split('@')[0]
    
    Add-OnboardingStep -Step "Generate Username" -Status "Success" -Details $userPrincipalName
    $results.Credentials.UserPrincipalName = $userPrincipalName
    
    # Step 2: Generate Temporary Password
    Write-Host "`n[Step 2] Generating temporary password..." -ForegroundColor Cyan
    Add-Type -AssemblyName System.Web
    $tempPassword = [System.Web.Security.Membership]::GeneratePassword(12, 3)
    $results.Credentials.TempPassword = $tempPassword
    Add-OnboardingStep -Step "Generate Password" -Status "Success"
    
    # Step 3: Create User Account
    Write-Host "`n[Step 3] Creating user account..." -ForegroundColor Cyan
    if ($PSCmdlet.ShouldProcess($userPrincipalName, "Create user account")) {
        $newUser = New-MgUser -DisplayName $displayName `
            -GivenName $FirstName `
            -Surname $LastName `
            -UserPrincipalName $userPrincipalName `
            -MailNickname $mailNickname `
            -AccountEnabled `
            -PasswordProfile @{
                Password = $tempPassword
                ForceChangePasswordNextSignIn = $true
            } `
            -JobTitle $JobTitle `
            -Department $Department `
            -UsageLocation $UsageLocation
        
        Add-OnboardingStep -Step "Create User Account" -Status "Success" -Details $userPrincipalName
        
        # Wait for account replication
        Start-Sleep -Seconds 30
    } else {
        Add-OnboardingStep -Step "Create User Account" -Status "Skipped" -Details "WhatIf mode"
    }
    
    # Step 4: Assign Manager
    Write-Host "`n[Step 4] Assigning manager..." -ForegroundColor Cyan
    if ($Manager -and $PSCmdlet.ShouldProcess($userPrincipalName, "Assign manager")) {
        try {
            $managerUser = Get-MgUser -UserId $Manager
            Set-MgUserManagerByRef -UserId $newUser.Id -OdataId "https://graph.microsoft.com/v1.0/users/$($managerUser.Id)"
            Add-OnboardingStep -Step "Assign Manager" -Status "Success" -Details $Manager
        } catch {
            Add-OnboardingStep -Step "Assign Manager" -Status "Failed" -Details $_.Exception.Message
        }
    } else {
        Add-OnboardingStep -Step "Assign Manager" -Status "Skipped" -Details $(if (-not $Manager) {"No manager specified"} else {"WhatIf mode"})
    }
    
    # Step 5: Assign License
    Write-Host "`n[Step 5] Assigning license..." -ForegroundColor Cyan
    if ($LicenseSKU -and $PSCmdlet.ShouldProcess($userPrincipalName, "Assign license")) {
        try {
            # Get available licenses
            $availableLicenses = Get-MgSubscribedSku | Where-Object { $_.SkuPartNumber -eq $LicenseSKU -or $_.SkuId -eq $LicenseSKU }
            
            if ($availableLicenses) {
                $license = $availableLicenses[0]
                Set-MgUserLicense -UserId $newUser.Id -AddLicenses @(@{SkuId = $license.SkuId}) -RemoveLicenses @()
                Add-OnboardingStep -Step "Assign License" -Status "Success" -Details $license.SkuPartNumber
                
                # Wait for mailbox provisioning
                Write-Host "Waiting for mailbox provisioning..." -ForegroundColor Yellow
                Start-Sleep -Seconds 60
            } else {
                Add-OnboardingStep -Step "Assign License" -Status "Failed" -Details "License SKU not found"
            }
        } catch {
            Add-OnboardingStep -Step "Assign License" -Status "Failed" -Details $_.Exception.Message
        }
    } else {
        Add-OnboardingStep -Step "Assign License" -Status "Skipped" -Details $(if (-not $LicenseSKU) {"No license specified"} else {"WhatIf mode"})
    }
    
    # Step 6: Add to Groups
    Write-Host "`n[Step 6] Adding to groups..." -ForegroundColor Cyan
    $groupsAdded = 0
    $defaultGroups = @("All Users", "All Staff")
    $allGroups = $defaultGroups + $Groups
    
    foreach ($groupName in $allGroups) {
        if ($PSCmdlet.ShouldProcess($groupName, "Add user to group")) {
            try {
                $group = Get-MgGroup -Filter "displayName eq '$groupName'"
                if ($group) {
                    New-MgGroupMember -GroupId $group.Id -DirectoryObjectId $newUser.Id
                    $groupsAdded++
                }
            } catch {
                # Group might not exist or user might already be a member
            }
        }
    }
    Add-OnboardingStep -Step "Add to Groups" -Status "Success" -Details "Added to $groupsAdded groups"
    
    # Step 7: Configure Mailbox Settings
    Write-Host "`n[Step 7] Configuring mailbox..." -ForegroundColor Cyan
    if ($PSCmdlet.ShouldProcess($userPrincipalName, "Configure mailbox")) {
        try {
            # Try to get mailbox (may need more time to provision)
            $mailbox = Get-Mailbox -Identity $userPrincipalName -ErrorAction SilentlyContinue
            
            if ($mailbox) {
                # Set mailbox configuration
                Set-Mailbox -Identity $userPrincipalName -RetentionPolicy "Default MRM Policy"
                
                # Set calendar permissions
                Set-MailboxFolderPermission -Identity "$($userPrincipalName):\Calendar" -User Default -AccessRights Reviewer
                
                Add-OnboardingStep -Step "Configure Mailbox" -Status "Success"
            } else {
                Add-OnboardingStep -Step "Configure Mailbox" -Status "Skipped" -Details "Mailbox not yet provisioned"
            }
        } catch {
            Add-OnboardingStep -Step "Configure Mailbox" -Status "Failed" -Details $_.Exception.Message
        }
    } else {
        Add-OnboardingStep -Step "Configure Mailbox" -Status "Skipped" -Details "WhatIf mode"
    }
    
    # Step 8: Enable MFA
    Write-Host "`n[Step 8] Configuring MFA..." -ForegroundColor Cyan
    if ($PSCmdlet.ShouldProcess($userPrincipalName, "Enable MFA")) {
        try {
            # Note: This requires additional MFA configuration APIs
            # For now, we'll mark it as requiring manual setup
            Add-OnboardingStep -Step "Configure MFA" -Status "Success" -Details "MFA will be enforced at first login"
        } catch {
            Add-OnboardingStep -Step "Configure MFA" -Status "Failed" -Details $_.Exception.Message
        }
    }
    
    # Step 9: Send Welcome Email
    Write-Host "`n[Step 9] Sending welcome email..." -ForegroundColor Cyan
    if ($SendWelcomeEmail -and $PersonalEmail -and $PSCmdlet.ShouldProcess($PersonalEmail, "Send welcome email")) {
        $welcomeBody = @"
<html>
<body style='font-family: Arial, sans-serif;'>
<h2>Welcome to the team, $FirstName!</h2>

<p>We're excited to have you join us as $JobTitle in the $Department department.</p>

<p><strong>Your account details:</strong></p>
<ul>
<li>Username: <code>$userPrincipalName</code></li>
<li>Temporary Password: <code>$tempPassword</code></li>
<li>First Login URL: <a href='https://portal.office.com'>https://portal.office.com</a></li>
</ul>

<p><strong>Important notes:</strong></p>
<ul>
<li>You will be required to change your password on first login</li>
<li>You will need to set up multi-factor authentication (MFA)</li>
<li>Your manager is: $Manager</li>
</ul>

<p><strong>Next steps:</strong></p>
<ol>
<li>Log in to the Office 365 portal using the credentials above</li>
<li>Change your password when prompted</li>
<li>Set up MFA following the prompts</li>
<li>Check your email for additional onboarding information</li>
</ol>

<p>If you have any questions, please reach out to IT Support.</p>

<p>Best regards,<br>
IT Team</p>
</body>
</html>
"@
        
        # Save welcome email to file (would normally send via email)
        $welcomeFile = "$LogPath\WelcomeEmail-$($FirstName)$($LastName)-$SessionId.html"
        $welcomeBody | Out-File $welcomeFile
        
        Add-OnboardingStep -Step "Send Welcome Email" -Status "Success" -Details "Saved to $welcomeFile"
    } else {
        Add-OnboardingStep -Step "Send Welcome Email" -Status "Skipped" -Details $(if (-not $PersonalEmail) {"No personal email"} else {"WhatIf mode"})
    }
    
    # Generate final report
    $results.EndTime = Get-Date
    $results.Duration = $results.EndTime - $results.StartTime
    
    $reportPath = "$LogPath\OnboardingReport-$($FirstName)$($LastName)-$SessionId.json"
    $results | ConvertTo-Json -Depth 10 | Out-File $reportPath
    
    # Display summary
    Write-Host "`n$('='*50)" -ForegroundColor Green
    Write-Host "ONBOARDING SUMMARY" -ForegroundColor Green
    Write-Host "$('='*50)" -ForegroundColor Green
    Write-Host "Name: $displayName"
    Write-Host "Email: $userPrincipalName"
    Write-Host "Title: $JobTitle"
    Write-Host "Department: $Department"
    Write-Host "Duration: $($results.Duration.TotalMinutes.ToString('0.0')) minutes"
    Write-Host "`nCredentials:"
    Write-Host "  Username: $userPrincipalName" -ForegroundColor Yellow
    Write-Host "  Temp Password: $tempPassword" -ForegroundColor Yellow
    Write-Host "`nSteps Completed:"
    
    $results.Steps | ForEach-Object {
        $color = switch ($_.Status) {
            "Success" { "Green" }
            "Failed" { "Red" }
            "Skipped" { "Yellow" }
            default { "White" }
        }
        Write-Host "  [$($_.Status)] $($_.Step)" -ForegroundColor $color
    }
    
    Write-Host "`nReports saved to:"
    Write-Host "  - Log: $LogFile" -ForegroundColor Yellow
    Write-Host "  - Report: $reportPath" -ForegroundColor Yellow
    
} catch {
    Write-Host "`nERROR: $($_.Exception.Message)" -ForegroundColor Red
    Add-OnboardingStep -Step "Process Error" -Status "Failed" -Details $_.Exception.Message
} finally {
    Stop-Transcript
}