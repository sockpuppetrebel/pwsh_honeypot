# Generate new certificate for Azure app authentication
$certName = "AzureAppAuth-$(Get-Date -Format 'yyyyMMdd')"
$cert = New-SelfSignedCertificate -Subject "CN=$certName" -CertStoreLocation "Cert:\CurrentUser\My" -KeyExportPolicy Exportable -KeySpec Signature -KeyLength 2048 -KeyAlgorithm RSA -HashAlgorithm SHA256 -NotAfter (Get-Date).AddYears(2)

Write-Host "Certificate created successfully!" -ForegroundColor Green
Write-Host "Subject: $($cert.Subject)"
Write-Host "Thumbprint: $($cert.Thumbprint)"
Write-Host "Expires: $($cert.NotAfter)"

# Export certificate for Azure upload
$certPath = ".\$certName.cer"
Export-Certificate -Cert $cert -FilePath $certPath
Write-Host "Certificate exported to: $certPath" -ForegroundColor Yellow

Write-Host "`nNext steps:" -ForegroundColor Cyan
Write-Host "1. Upload $certPath to Azure App Registration (replace cert ID: 4458824c-6df6-42e3-b8e9-4a334e46a42b)"
Write-Host "2. Update your PowerShell script with thumbprint: $($cert.Thumbprint)"