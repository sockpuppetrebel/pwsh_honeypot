# Test PnP PowerShell connection with existing certificate

# Load configuration
. (Join-Path (Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)) "Config.ps1")

$certPath = $CertificatePaths.AzureCert
$keyPath = $CertificatePaths.AzureKey

# Check if certificate files exist
if (-not (Test-CertificateExists $certPath)) {
    exit 1
}
if (-not (Test-CertificateExists $keyPath)) {
    exit 1
}

# Create X509Certificate2 from PEM files
try {
    $certContent = Get-Content $certPath -Raw
    $keyContent = Get-Content $keyPath -Raw
    
    # Create certificate from PEM
    $cert = [System.Security.Cryptography.X509Certificates.X509Certificate2]::CreateFromPem($certContent, $keyContent)
    Write-Host "Certificate loaded successfully. Thumbprint: $($cert.Thumbprint)" -ForegroundColor Green
} catch {
    Write-Host "Failed to load certificate: $_" -ForegroundColor Red
    exit 1
}

# First, we need to add the certificate to the certificate store temporarily
Write-Host "`nAdding certificate to CurrentUser store temporarily..." -ForegroundColor Yellow
$store = New-Object System.Security.Cryptography.X509Certificates.X509Store("My", "CurrentUser")
$store.Open("ReadWrite")
$store.Add($cert)
$store.Close()
Write-Host "Certificate added with thumbprint: $($cert.Thumbprint)" -ForegroundColor Gray

# Connect to SharePoint Online using PnP
Write-Host "`nConnecting to SharePoint Online with PnP PowerShell..." -ForegroundColor Yellow

try {
    # Connect to the SharePoint admin center using thumbprint
    Connect-PnPOnline -Url "https://episerver99-admin.sharepoint.com" `
                      -ClientId "fe2a9efe-3000-4b02-96ea-344a2583dd52" `
                      -Tenant "3ec00d79-021a-42d4-aac8-dcb35973dff2" `
                      -Thumbprint $cert.Thumbprint
    
    Write-Host "✓ Successfully connected to SharePoint Admin Center!" -ForegroundColor Green
    
    # Get connection info
    $connection = Get-PnPConnection
    Write-Host "`nConnection Details:" -ForegroundColor Cyan
    Write-Host "  URL: $($connection.Url)" -ForegroundColor White
    Write-Host "  ClientId: $($connection.ClientId)" -ForegroundColor White
    Write-Host "  Tenant: $($connection.Tenant)" -ForegroundColor White
    
    # Test by getting tenant info
    Write-Host "`nTesting connection by getting tenant info..." -ForegroundColor Yellow
    $tenant = Get-PnPTenant
    Write-Host "✓ Tenant Title: $($tenant.Title)" -ForegroundColor Green
    
    # Test getting site collections
    Write-Host "`nTesting by getting first 5 site collections..." -ForegroundColor Yellow
    $sites = Get-PnPTenantSite | Select-Object -First 5
    
    if ($sites) {
        Write-Host "✓ Found site collections:" -ForegroundColor Green
        $sites | ForEach-Object {
            Write-Host "  - $($_.Title) ($($_.Url))" -ForegroundColor White
        }
    }
    
} catch {
    Write-Host "✗ Failed to connect: $_" -ForegroundColor Red
    Write-Host "Error type: $($_.Exception.GetType().FullName)" -ForegroundColor Red
    
    # Common issues
    if ($_.Exception.Message -match "401") {
        Write-Host "`nPossible issues:" -ForegroundColor Yellow
        Write-Host "- App registration needs SharePoint app-only permissions" -ForegroundColor Yellow
        Write-Host "- Certificate thumbprint mismatch" -ForegroundColor Yellow
    }
    
    if ($_.Exception.Message -match "AADSTS700027") {
        Write-Host "`nThe certificate is not properly registered with the app." -ForegroundColor Yellow
    }
}

# Disconnect
Disconnect-PnPOnline

# Clean up - remove certificate from store
Write-Host "`nCleaning up certificate from store..." -ForegroundColor Yellow
try {
    $store = New-Object System.Security.Cryptography.X509Certificates.X509Store("My", "CurrentUser")
    $store.Open("ReadWrite")
    $certToRemove = $store.Certificates | Where-Object { $_.Thumbprint -eq $cert.Thumbprint }
    if ($certToRemove) {
        $store.Remove($certToRemove)
        Write-Host "Certificate removed from store" -ForegroundColor Gray
    }
    $store.Close()
} catch {
    Write-Host "Warning: Could not remove certificate from store: $_" -ForegroundColor Yellow
}