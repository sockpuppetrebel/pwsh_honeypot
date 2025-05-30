# Recreate PEM certificates for Azure app authentication

Write-Host "Generating new self-signed certificate..." -ForegroundColor Yellow

# Generate certificate
$certName = "AzureAppAuth-$(Get-Date -Format 'yyyyMMdd')"
$cert = New-SelfSignedCertificate -Subject "CN=$certName" `
                                  -CertStoreLocation "Cert:\CurrentUser\My" `
                                  -KeyExportPolicy Exportable `
                                  -KeySpec Signature `
                                  -KeyLength 2048 `
                                  -KeyAlgorithm RSA `
                                  -HashAlgorithm SHA256 `
                                  -NotAfter (Get-Date).AddYears(2)

Write-Host "Certificate created:" -ForegroundColor Green
Write-Host "  Subject: $($cert.Subject)" -ForegroundColor White
Write-Host "  Thumbprint: $($cert.Thumbprint)" -ForegroundColor White
Write-Host "  Expires: $($cert.NotAfter)" -ForegroundColor White

# Export to PFX with temporary password
$pfxPassword = ConvertTo-SecureString -String "TempPassword123!" -Force -AsPlainText
$pfxPath = ".\temp_cert.pfx"
Export-PfxCertificate -Cert $cert -FilePath $pfxPath -Password $pfxPassword

Write-Host "`nConverting to PEM format..." -ForegroundColor Yellow

# Convert PFX to PEM using OpenSSL (if available)
try {
    # Export certificate (public key) to PEM
    $certPemPath = ".\azure_app_cert.pem"
    openssl pkcs12 -in $pfxPath -clcerts -nokeys -out $certPemPath -password pass:TempPassword123! -passin pass:TempPassword123!
    
    # Export private key to PEM
    $keyPemPath = ".\azure_app_key.pem"
    openssl pkcs12 -in $pfxPath -nocerts -nodes -out $keyPemPath -password pass:TempPassword123! -passin pass:TempPassword123!
    
    # Clean up temp PFX
    Remove-Item $pfxPath -Force
    
    Write-Host "✓ PEM files created:" -ForegroundColor Green
    Write-Host "  Certificate: $certPemPath" -ForegroundColor White
    Write-Host "  Private Key: $keyPemPath" -ForegroundColor White
    
    Write-Host "`n⚠ IMPORTANT: You need to upload the certificate to Azure:" -ForegroundColor Yellow
    Write-Host "1. Go to Azure App Registration (fe2a9efe-3000-4b02-96ea-344a2583dd52)" -ForegroundColor White
    Write-Host "2. Go to Certificates & secrets" -ForegroundColor White
    Write-Host "3. Upload the certificate file: $certPemPath" -ForegroundColor White
    Write-Host "4. Replace the old certificate" -ForegroundColor White
    
} catch {
    Write-Host "OpenSSL not found. Using PowerShell method..." -ForegroundColor Yellow
    
    # Alternative: Use PowerShell native methods
    $certBase64 = [Convert]::ToBase64String($cert.RawData, 'InsertLineBreaks')
    $certPem = "-----BEGIN CERTIFICATE-----`n$certBase64`n-----END CERTIFICATE-----"
    
    # Export certificate to PEM
    $certPem | Out-File -FilePath ".\azure_app_cert.pem" -Encoding ASCII
    
    # For private key, we need to export and convert
    $privateKeyBytes = $cert.PrivateKey.ExportPkcs8PrivateKey()
    $privateKeyBase64 = [Convert]::ToBase64String($privateKeyBytes, 'InsertLineBreaks')
    $privateKeyPem = "-----BEGIN PRIVATE KEY-----`n$privateKeyBase64`n-----END PRIVATE KEY-----"
    
    $privateKeyPem | Out-File -FilePath ".\azure_app_key.pem" -Encoding ASCII
    
    Write-Host "✓ PEM files created using PowerShell method" -ForegroundColor Green
    Write-Host "  Certificate: azure_app_cert.pem" -ForegroundColor White
    Write-Host "  Private Key: azure_app_key.pem" -ForegroundColor White
}

Write-Host "`nCertificate thumbprint for scripts: $($cert.Thumbprint)" -ForegroundColor Cyan