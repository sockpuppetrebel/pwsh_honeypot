# Test authentication with detailed error output
$ErrorActionPreference = "Stop"

# Certificate paths
$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
$certPath = Join-Path $scriptPath "azure_app_cert.pem"
$keyPath = Join-Path $scriptPath "azure_app_key.pem"

# Load certificate
Write-Host "Loading certificate..." -ForegroundColor Yellow
$certContent = Get-Content $certPath -Raw
$keyContent = Get-Content $keyPath -Raw
$cert = [System.Security.Cryptography.X509Certificates.X509Certificate2]::CreateFromPem($certContent, $keyContent)

Write-Host "Certificate loaded. Thumbprint: $($cert.Thumbprint)" -ForegroundColor Green

# Try to connect with verbose output
Write-Host "`nAttempting to connect to Microsoft Graph..." -ForegroundColor Yellow
try {
    Connect-MgGraph -TenantId "3ec00d79-021a-42d4-aac8-dcb35973dff2" `
                    -ClientId "fe2a9efe-3000-4b02-96ea-344a2583dd52" `
                    -Certificate $cert `
                    -NoWelcome `
                    -Debug
    
    Write-Host "✓ Successfully connected!" -ForegroundColor Green
    
    # Test a simple Graph call
    Write-Host "`nTesting Graph API access..." -ForegroundColor Yellow
    $context = Get-MgContext
    Write-Host "Connected as: $($context.ClientId)" -ForegroundColor Green
    Write-Host "Auth Type: $($context.AuthType)" -ForegroundColor Green
    Write-Host "Scopes: $($context.Scopes -join ', ')" -ForegroundColor Green
    
    # Try to list sites
    Write-Host "`nTrying to list SharePoint sites..." -ForegroundColor Yellow
    $sites = Get-MgSite -Top 1
    Write-Host "✓ Can access sites!" -ForegroundColor Green
    
} catch {
    Write-Host "✗ Authentication failed!" -ForegroundColor Red
    Write-Host "Error: $_" -ForegroundColor Red
    Write-Host "`nFull error details:" -ForegroundColor Yellow
    Write-Host $_.Exception.ToString() -ForegroundColor Red
    
    if ($_.Exception.InnerException) {
        Write-Host "`nInner exception:" -ForegroundColor Yellow
        Write-Host $_.Exception.InnerException.ToString() -ForegroundColor Red
    }
} finally {
    if (Get-MgContext) {
        Disconnect-MgGraph -NoWelcome
    }
}