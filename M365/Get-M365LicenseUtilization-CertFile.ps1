#Requires -Modules Microsoft.Graph.Users, Microsoft.Graph.Identity.DirectoryManagement

<#
.SYNOPSIS
    Analyzes M365 license utilization using certificate file authentication
.DESCRIPTION
    Same as Get-M365LicenseUtilization.ps1 but uses a certificate file for auth
.PARAMETER CertificatePath
    Path to the certificate file (.pfx or .p12)
.PARAMETER CertificatePassword
    Certificate password (will prompt if not provided)
#>

param(
    [Parameter(Mandatory = $false)]
    [switch]$ExportToExcel,
    
    [Parameter(Mandatory = $false)]
    [switch]$ShowInactiveUsers,
    
    [Parameter(Mandatory = $false)]
    [int]$InactiveDays = 90,
    
    [Parameter(Mandatory = $true)]
    [string]$CertificatePath,
    
    [Parameter(Mandatory = $false)]
    [SecureString]$CertificatePassword
)

# Connect to Microsoft Graph with certificate file
try {
    $tenantId = '3ec00d79-021a-42d4-aac8-dcb35973dff2'
    $clientId = 'fe2a9efe-3000-4b02-96ea-344a2583dd52'
    
    if (-not $CertificatePassword) {
        $CertificatePassword = Read-Host -AsSecureString "Enter certificate password"
    }
    
    # Load certificate from file
    $cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2(
        $CertificatePath, 
        $CertificatePassword,
        [System.Security.Cryptography.X509Certificates.X509KeyStorageFlags]::PersistKeySet
    )
    
    Connect-MgGraph -ClientId $clientId -TenantId $tenantId -Certificate $cert
    Write-Host "Connected to Microsoft Graph using certificate file" -ForegroundColor Green
}
catch {
    Write-Error "Failed to connect to Microsoft Graph: $($_.Exception.Message)"
    exit 1
}

# Rest of the script remains the same...
# Copy the rest from the original script starting from line 58