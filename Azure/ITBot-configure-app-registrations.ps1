#Requires -Version 7.0
#Requires -Modules Microsoft.Graph.Authentication, Microsoft.Graph.Applications

<#
.SYNOPSIS
    Configures Azure AD app registrations for IT-Bot testing
.DESCRIPTION
    This script creates and configures Azure AD app registrations for:
    - IT-Bot Function App
    - Microsoft Graph API access
    - Okta integration (optional)
.PARAMETER TenantId
    The Azure AD tenant ID
.PARAMETER FunctionAppUrl
    The URL of the deployed Function App
.EXAMPLE
    .\configure-app-registrations.ps1 -TenantId "your-tenant-id" -FunctionAppUrl "https://itbot-testing-func-abc123.azurewebsites.net"
#>

param(
    [Parameter(Mandatory = $true)]
    [string]$TenantId,
    
    [Parameter(Mandatory = $true)]
    [string]$FunctionAppUrl
)

# Required Graph API permissions
$GraphApiPermissions = @(
    @{ Id = "User.Read.All"; Type = "Application" },
    @{ Id = "User.ReadWrite.All"; Type = "Application" },
    @{ Id = "Group.Read.All"; Type = "Application" },
    @{ Id = "Directory.Read.All"; Type = "Application" },
    @{ Id = "UserAuthenticationMethod.ReadWrite.All"; Type = "Application" }
)

function Write-Log {
    param($Message, $Level = "INFO")
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Write-Host "[$timestamp] [$Level] $Message"
}

function Connect-ToMicrosoftGraph {
    param([string]$TenantId)
    
    try {
        Write-Log "Connecting to Microsoft Graph..."
        Connect-MgGraph -TenantId $TenantId -Scopes "Application.ReadWrite.All", "Directory.ReadWrite.All"
        
        $context = Get-MgContext
        Write-Log "Connected successfully to tenant: $($context.TenantId)"
        return $true
    }
    catch {
        Write-Log "Failed to connect to Microsoft Graph: $($_.Exception.Message)" "ERROR"
        return $false
    }
}

function Get-GraphServicePrincipal {
    try {
        $graphSp = Get-MgServicePrincipal -Filter "appId eq '00000003-0000-0000-c000-000000000000'"
        return $graphSp
    }
    catch {
        Write-Log "Failed to get Microsoft Graph service principal: $($_.Exception.Message)" "ERROR"
        return $null
    }
}

function New-AppRegistration {
    param(
        [string]$DisplayName,
        [string]$ReplyUrl,
        [array]$RequiredResourceAccess
    )
    
    try {
        Write-Log "Creating app registration: $DisplayName"
        
        $webApp = @{
            RedirectUris = @($ReplyUrl)
            ImplicitGrantSettings = @{
                EnableAccessTokenIssuance = $false
                EnableIdTokenIssuance = $false
            }
        }
        
        $appParams = @{
            DisplayName = $DisplayName
            SignInAudience = "AzureADMyOrg"
            RequiredResourceAccess = $RequiredResourceAccess
            Web = $webApp
        }
        
        $app = New-MgApplication @appParams
        Write-Log "Created app registration: $($app.DisplayName) (App ID: $($app.AppId))"
        
        # Create service principal
        $spParams = @{
            AppId = $app.AppId
        }
        $servicePrincipal = New-MgServicePrincipal @spParams
        Write-Log "Created service principal for app: $($servicePrincipal.DisplayName)"
        
        # Create client secret
        $secretParams = @{
            PasswordCredential = @{
                DisplayName = "IT-Bot Testing Secret"
                EndDateTime = (Get-Date).AddMonths(12)
            }
        }
        $secret = Add-MgApplicationPassword -ApplicationId $app.Id @secretParams
        Write-Log "Created client secret (expires: $($secret.EndDateTime))"
        
        return @{
            Application = $app
            ServicePrincipal = $servicePrincipal
            ClientSecret = $secret.SecretText
        }
    }
    catch {
        Write-Log "Failed to create app registration: $($_.Exception.Message)" "ERROR"
        return $null
    }
}

function Grant-AdminConsent {
    param(
        [string]$ServicePrincipalId,
        [array]$Permissions,
        [object]$GraphServicePrincipal
    )
    
    try {
        Write-Log "Granting admin consent for permissions..."
        
        foreach ($permission in $Permissions) {
            $graphPermission = $GraphServicePrincipal.AppRoles | Where-Object { $_.Value -eq $permission.Id }
            
            if ($graphPermission) {
                $grantParams = @{
                    PrincipalId = $ServicePrincipalId
                    ResourceId = $GraphServicePrincipal.Id
                    AppRoleId = $graphPermission.Id
                }
                
                try {
                    New-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $ServicePrincipalId @grantParams
                    Write-Log "Granted permission: $($permission.Id)"
                }
                catch {
                    if ($_.Exception.Message -like "*Permission being assigned already exists*") {
                        Write-Log "Permission $($permission.Id) already granted" "INFO"
                    }
                    else {
                        Write-Log "Failed to grant permission $($permission.Id): $($_.Exception.Message)" "WARNING"
                    }
                }
            }
            else {
                Write-Log "Permission not found: $($permission.Id)" "WARNING"
            }
        }
    }
    catch {
        Write-Log "Failed to grant admin consent: $($_.Exception.Message)" "ERROR"
    }
}

# Main execution
Write-Log "Starting IT-Bot app registration configuration"

# Connect to Microsoft Graph
if (-not (Connect-ToMicrosoftGraph -TenantId $TenantId)) {
    Write-Log "Cannot proceed without Microsoft Graph connection" "ERROR"
    exit 1
}

# Get Microsoft Graph service principal
$graphSp = Get-GraphServicePrincipal
if (-not $graphSp) {
    Write-Log "Cannot proceed without Microsoft Graph service principal" "ERROR"
    exit 1
}

# Prepare required resource access for Microsoft Graph
$resourceAccess = @()
foreach ($permission in $GraphApiPermissions) {
    $graphPermission = $graphSp.AppRoles | Where-Object { $_.Value -eq $permission.Id }
    if ($graphPermission) {
        $resourceAccess += @{
            Id = $graphPermission.Id
            Type = "Role"
        }
    }
    else {
        Write-Log "Permission not found in Graph API: $($permission.Id)" "WARNING"
    }
}

$requiredResourceAccess = @(
    @{
        ResourceAppId = "00000003-0000-0000-c000-000000000000"  # Microsoft Graph
        ResourceAccess = $resourceAccess
    }
)

# Create main IT-Bot app registration
$replyUrl = "$FunctionAppUrl/.auth/login/aad/callback"
$itBotApp = New-AppRegistration -DisplayName "IT-Bot Testing" -ReplyUrl $replyUrl -RequiredResourceAccess $requiredResourceAccess

if (-not $itBotApp) {
    Write-Log "Failed to create IT-Bot app registration" "ERROR"
    exit 1
}

# Wait for service principal to be available
Write-Log "Waiting for service principal to be available..."
Start-Sleep -Seconds 10

# Grant admin consent
Grant-AdminConsent -ServicePrincipalId $itBotApp.ServicePrincipal.Id -Permissions $GraphApiPermissions -GraphServicePrincipal $graphSp

# Create configuration file
$configPath = "itbot-app-config.json"
$config = @{
    TenantId = $TenantId
    ClientId = $itBotApp.Application.AppId
    ClientSecret = $itBotApp.ClientSecret
    FunctionAppUrl = $FunctionAppUrl
    GraphApiUrl = "https://graph.microsoft.com"
    Scopes = @(
        "https://graph.microsoft.com/.default"
    )
    CreatedDate = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
} | ConvertTo-Json -Depth 3

$config | Out-File -FilePath $configPath -Encoding UTF8
Write-Log "Configuration saved to: $configPath"

# Create environment variables file
$envPath = "itbot-testing.env"
$envContent = @"
# IT-Bot Testing Environment Variables
# Generated on $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")

AZURE_TENANT_ID=$TenantId
AZURE_CLIENT_ID=$($itBotApp.Application.AppId)
AZURE_CLIENT_SECRET=$($itBotApp.ClientSecret)
GRAPH_API_URL=https://graph.microsoft.com
FUNCTION_APP_URL=$FunctionAppUrl

# For local development
AZURE_SUBSCRIPTION_ID=$(az account show --query id -o tsv)
AZURE_RESOURCE_GROUP=rg-itbot-testing
"@

$envContent | Out-File -FilePath $envPath -Encoding UTF8
Write-Log "Environment variables saved to: $envPath"

Write-Log ""
Write-Log "App registration configuration completed!"
Write-Log ""
Write-Log "App Registration Details:"
Write-Log "  Display Name: $($itBotApp.Application.DisplayName)"
Write-Log "  App ID: $($itBotApp.Application.AppId)"
Write-Log "  Object ID: $($itBotApp.Application.Id)"
Write-Log "  Reply URL: $replyUrl"
Write-Log ""
Write-Log "IMPORTANT SECURITY NOTES:"
Write-Log "1. Client secret is saved in the configuration files above"
Write-Log "2. Store these secrets securely (Azure Key Vault recommended)"
Write-Log "3. The client secret expires in 12 months"
Write-Log "4. Admin consent has been granted for required permissions"
Write-Log ""
Write-Log "Next Steps:"
Write-Log "1. Update your Function App with the configuration values"
Write-Log "2. Test the authentication flow"
Write-Log "3. Configure Okta integration if needed"

# Disconnect from Microsoft Graph
Disconnect-MgGraph
Write-Log "Disconnected from Microsoft Graph"