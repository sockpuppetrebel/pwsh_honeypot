# Export Azure AD Group Memberships for a User
# Gets all groups a user is a member of and exports to CSV

param(
    [string]$UserEmail = "kaila.trapani@optimizely.com",
    [string]$OutputPath = $null  # Will auto-generate if not provided
)

$ErrorActionPreference = 'Stop'
$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
$projectRoot = Split-Path -Parent (Split-Path -Parent $scriptPath)
$certPath = Join-Path $projectRoot "certificates/azure/azure_app_cert.pem"
$keyPath = Join-Path $projectRoot "certificates/azure/azure_app_key.pem"

# Load certificate
$certContent = Get-Content $certPath -Raw
$keyContent = Get-Content $keyPath -Raw
$cert = [System.Security.Cryptography.X509Certificates.X509Certificate2]::CreateFromPem($certContent, $keyContent)

Write-Host "`n=============== AZURE AD GROUP EXPORT ===============" -ForegroundColor Cyan
Write-Host "Target User: $UserEmail" -ForegroundColor White
Write-Host "Exporting all Azure AD group memberships..." -ForegroundColor Yellow

try {
    # Connect to Microsoft Graph using certificate
    Write-Host "`nConnecting to Microsoft Graph..." -ForegroundColor Yellow
    Connect-MgGraph -TenantId "3ec00d79-021a-42d4-aac8-dcb35973dff2" `
                    -ClientId "fe2a9efe-3000-4b02-96ea-344a2583dd52" `
                    -Certificate $cert `
                    -NoWelcome
    
    Write-Host "✓ Connected to Microsoft Graph" -ForegroundColor Green
    
    # Get the user first
    Write-Host "`nLooking up user: $UserEmail" -ForegroundColor Yellow
    $user = Get-MgUser -Filter "mail eq '$UserEmail' or userPrincipalName eq '$UserEmail'" -ErrorAction Stop
    
    if (-not $user) {
        Write-Host "❌ User not found: $UserEmail" -ForegroundColor Red
        exit 1
    }
    
    Write-Host "✓ Found user: $($user.DisplayName) ($($user.UserPrincipalName))" -ForegroundColor Green
    Write-Host "  Object ID: $($user.Id)" -ForegroundColor Gray
    Write-Host "  Account Enabled: $($user.AccountEnabled)" -ForegroundColor Gray
    
    # Get all group memberships (including nested groups)
    Write-Host "`nGetting group memberships..." -ForegroundColor Yellow
    $groupMemberships = Get-MgUserMemberOf -UserId $user.Id -All
    
    Write-Host "✓ Found $($groupMemberships.Count) group memberships" -ForegroundColor Green
    
    # Process groups and create export data
    $exportData = @()
    $securityGroups = 0
    $distributionGroups = 0
    $m365Groups = 0
    $dynamicGroups = 0
    
    foreach ($membership in $groupMemberships) {
        # Get detailed group information
        try {
            $group = Get-MgGroup -GroupId $membership.Id -ErrorAction Stop
            
            # Determine group type
            $groupType = "Unknown"
            $isSecurityGroup = $group.SecurityEnabled
            $isMailEnabled = $group.MailEnabled
            $isDynamic = $group.GroupTypes -contains "DynamicMembership"
            $isM365 = $group.GroupTypes -contains "Unified"
            
            if ($isM365) {
                $groupType = "Microsoft 365 Group"
                $m365Groups++
            } elseif ($isDynamic) {
                $groupType = "Dynamic Security Group"
                $dynamicGroups++
            } elseif ($isSecurityGroup -and -not $isMailEnabled) {
                $groupType = "Security Group"
                $securityGroups++
            } elseif (-not $isSecurityGroup -and $isMailEnabled) {
                $groupType = "Distribution Group"
                $distributionGroups++
            } elseif ($isSecurityGroup -and $isMailEnabled) {
                $groupType = "Mail-Enabled Security Group"
                $securityGroups++
            }
            
            $exportData += [PSCustomObject]@{
                DisplayName = $group.DisplayName
                GroupType = $groupType
                Mail = $group.Mail
                MailNickname = $group.MailNickname
                Description = $group.Description
                ObjectId = $group.Id
                SecurityEnabled = $group.SecurityEnabled
                MailEnabled = $group.MailEnabled
                IsDynamic = $isDynamic
                IsM365Group = $isM365
                CreatedDateTime = $group.CreatedDateTime
                Visibility = $group.Visibility
                MembershipRule = $group.MembershipRule
                OnPremisesSyncEnabled = $group.OnPremisesSyncEnabled
                UserEmail = $UserEmail
                UserDisplayName = $user.DisplayName
                UserObjectId = $user.Id
                ExportDate = Get-Date
            }
            
        } catch {
            Write-Host "⚠️  Could not get details for group: $($membership.Id)" -ForegroundColor Yellow
            
            # Add basic info even if we can't get full details
            $exportData += [PSCustomObject]@{
                DisplayName = "Unknown"
                GroupType = "Could not retrieve"
                Mail = ""
                MailNickname = ""
                Description = ""
                ObjectId = $membership.Id
                SecurityEnabled = ""
                MailEnabled = ""
                IsDynamic = ""
                IsM365Group = ""
                CreatedDateTime = ""
                Visibility = ""
                MembershipRule = ""
                OnPremisesSyncEnabled = ""
                UserEmail = $UserEmail
                UserDisplayName = $user.DisplayName
                UserObjectId = $user.Id
                ExportDate = Get-Date
            }
        }
    }
    
    # Generate output path if not provided
    if (-not $OutputPath) {
        $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
        $safeUserName = $UserEmail.Split('@')[0] -replace '[^\w]', '_'
        $OutputPath = Join-Path $scriptPath "AAD_Groups_${safeUserName}_${timestamp}.csv"
    }
    
    # Export to CSV
    Write-Host "`nExporting to CSV..." -ForegroundColor Yellow
    $exportData | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
    
    Write-Host "`n=============== EXPORT SUMMARY ===============" -ForegroundColor Cyan
    Write-Host "User: $($user.DisplayName) ($($user.UserPrincipalName))" -ForegroundColor White
    Write-Host "Total Groups: $($exportData.Count)" -ForegroundColor White
    Write-Host "  - Security Groups: $securityGroups" -ForegroundColor Green
    Write-Host "  - Distribution Groups: $distributionGroups" -ForegroundColor Blue
    Write-Host "  - Microsoft 365 Groups: $m365Groups" -ForegroundColor Cyan
    Write-Host "  - Dynamic Groups: $dynamicGroups" -ForegroundColor Yellow
    
    Write-Host "`n📄 Groups exported to: $OutputPath" -ForegroundColor Green
    
    # Show some sample groups
    if ($exportData.Count -gt 0) {
        Write-Host "`n📋 Sample Groups:" -ForegroundColor Yellow
        $exportData | Select-Object -First 10 | ForEach-Object {
            Write-Host "  - $($_.DisplayName) [$($_.GroupType)]" -ForegroundColor Gray
        }
        
        if ($exportData.Count -gt 10) {
            Write-Host "  ... and $($exportData.Count - 10) more groups" -ForegroundColor Gray
        }
    }
    
    # Check for SharePoint-related groups
    $sharePointGroups = $exportData | Where-Object { 
        $_.DisplayName -like "*SharePoint*" -or 
        $_.DisplayName -like "*Site*" -or 
        $_.Mail -like "*sharepoint*" 
    }
    
    if ($sharePointGroups.Count -gt 0) {
        Write-Host "`n🔍 SharePoint-Related Groups Found:" -ForegroundColor Cyan
        $sharePointGroups | ForEach-Object {
            Write-Host "  - $($_.DisplayName) [$($_.GroupType)]" -ForegroundColor Yellow
        }
    } else {
        Write-Host "`n⚠️  No obvious SharePoint-related groups found" -ForegroundColor Yellow
        Write-Host "Check the CSV file for groups that might control SharePoint access" -ForegroundColor Gray
    }
    
} catch {
    Write-Host "`n❌ Error: $($_.Exception.Message)" -ForegroundColor Red
    if ($_.Exception.InnerException) {
        Write-Host "Inner Exception: $($_.Exception.InnerException.Message)" -ForegroundColor Red
    }
} finally {
    # Disconnect from Graph
    try {
        Disconnect-MgGraph -ErrorAction SilentlyContinue
        Write-Host "`n✓ Disconnected from Microsoft Graph" -ForegroundColor Green
    } catch {
        # Ignore disconnect errors
    }
}

Write-Host "`n✅ Group export complete!" -ForegroundColor Green