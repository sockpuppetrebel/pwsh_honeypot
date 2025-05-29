# Certificate diagnostics script
$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
$certPath = Join-Path $scriptPath "azure_app_cert.pem"
$keyPath = Join-Path $scriptPath "azure_app_key.pem"

Write-Host "`nCertificate Diagnostics" -ForegroundColor Cyan
Write-Host "======================" -ForegroundColor Cyan

# Check if files exist
Write-Host "`nChecking certificate files..."
if (Test-Path $certPath) {
    Write-Host "✓ Certificate file found: $certPath" -ForegroundColor Green
} else {
    Write-Host "✗ Certificate file NOT found: $certPath" -ForegroundColor Red
    exit 1
}

if (Test-Path $keyPath) {
    Write-Host "✓ Key file found: $keyPath" -ForegroundColor Green
} else {
    Write-Host "✗ Key file NOT found: $keyPath" -ForegroundColor Red
    exit 1
}

# Try to load certificate
Write-Host "`nLoading certificate..."
try {
    $certContent = Get-Content $certPath -Raw
    $keyContent = Get-Content $keyPath -Raw
    
    $cert = [System.Security.Cryptography.X509Certificates.X509Certificate2]::CreateFromPem($certContent, $keyContent)
    Write-Host "✓ Certificate loaded successfully" -ForegroundColor Green
    
    # Display certificate details
    Write-Host "`nCertificate Details:" -ForegroundColor Yellow
    Write-Host "Subject: $($cert.Subject)"
    Write-Host "Issuer: $($cert.Issuer)"
    Write-Host "Thumbprint: $($cert.Thumbprint)"
    Write-Host "Serial Number: $($cert.SerialNumber)"
    Write-Host "Not Before: $($cert.NotBefore)"
    Write-Host "Not After: $($cert.NotAfter)"
    
    # Check if certificate is valid
    if ($cert.NotBefore -gt (Get-Date)) {
        Write-Host "`n⚠ Certificate is not yet valid!" -ForegroundColor Red
    } elseif ($cert.NotAfter -lt (Get-Date)) {
        Write-Host "`n⚠ Certificate has expired!" -ForegroundColor Red
    } else {
        Write-Host "`n✓ Certificate is within validity period" -ForegroundColor Green
    }
    
    # Display app registration info
    Write-Host "`nApp Registration Info:" -ForegroundColor Yellow
    Write-Host "Tenant ID: 3ec00d79-021a-42d4-aac8-dcb35973dff2"
    Write-Host "Client ID: fe2a9efe-3000-4b02-96ea-344a2583dd52"
    
    Write-Host "`nIMPORTANT: The certificate thumbprint shown above must be registered in your Azure AD app registration!" -ForegroundColor Cyan
    Write-Host "To fix the authentication error:" -ForegroundColor Yellow
    Write-Host "1. Go to Azure Portal > Azure Active Directory > App registrations"
    Write-Host "2. Find your app (Client ID: fe2a9efe-3000-4b02-96ea-344a2583dd52)"
    Write-Host "3. Go to 'Certificates & secrets' > 'Certificates' tab"
    Write-Host "4. Upload the certificate or verify the thumbprint matches: $($cert.Thumbprint)"
    Write-Host "5. Ensure the app has the required API permissions for Microsoft Graph"
    
} catch {
    Write-Host "✗ Failed to load certificate: $_" -ForegroundColor Red
    Write-Host "`nError details: $($_.Exception.Message)" -ForegroundColor Red
}