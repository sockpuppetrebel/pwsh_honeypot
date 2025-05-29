# Quick removal from known SharePoint sites
param(
    [string]$UserEmail = "kaila.trapani@optimizely.com"
)

$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
$certPath = Join-Path $scriptPath "azure_app_cert.pem"  
$keyPath = Join-Path $scriptPath "azure_app_key.pem"

# Load certificate
$certContent = Get-Content $certPath -Raw
$keyContent = Get-Content $keyPath -Raw
$cert = [System.Security.Cryptography.X509Certificates.X509Certificate2]::CreateFromPem($certContent, $keyContent)

# Add to store
$store = New-Object System.Security.Cryptography.X509Certificates.X509Store("My", "CurrentUser")
$store.Open("ReadWrite")
$store.Add($cert)
$store.Close()

# Known sites
$knownSites = @(
    @{Url = "https://episerver99.sharepoint.com/sites/APACOppDusk"; Group = "Medlemmar på (APAC Opp) Dusk"},
    @{Url = "https://episerver99.sharepoint.com/sites/NET5"; Group = ".NET 5 Members"}
)

Write-Host "`n=============== REMOVING FROM KNOWN SITES ===============" -ForegroundColor Cyan
Write-Host "User: $UserEmail" -ForegroundColor White

foreach ($site in $knownSites) {
    Write-Host "`nSite: $($site.Url)" -ForegroundColor Yellow
    Write-Host "Known group: $($site.Group)" -ForegroundColor Gray
    
    try {
        Connect-PnPOnline -Url $site.Url `
                          -ClientId "fe2a9efe-3000-4b02-96ea-344a2583dd52" `
                          -Tenant "3ec00d79-021a-42d4-aac8-dcb35973dff2" `
                          -Thumbprint $cert.Thumbprint
        
        # Find the group
        $group = Get-PnPGroup | Where-Object { $_.Title -eq $site.Group }
        if ($group) {
            # Get member
            $member = Get-PnPGroupMember -Group $group | Where-Object { $_.Email -eq $UserEmail }
            if ($member) {
                # Remove member
                Remove-PnPGroupMember -Group $group -LoginName $member.LoginName
                Write-Host "✓ Removed from group: $($group.Title)" -ForegroundColor Green
            } else {
                Write-Host "✗ User not found in group" -ForegroundColor Red
            }
        } else {
            Write-Host "✗ Group not found" -ForegroundColor Red
        }
        
        Disconnect-PnPOnline
        
    } catch {
        Write-Host "✗ Error: $_" -ForegroundColor Red
    }
}

# Clean up
$store.Open("ReadWrite")
$certToRemove = $store.Certificates | Where-Object { $_.Thumbprint -eq $cert.Thumbprint }
if ($certToRemove) { $store.Remove($certToRemove) }
$store.Close()

Write-Host "`n✓ Complete!" -ForegroundColor Green
Write-Host "`nFor comprehensive removal, run the full audit script overnight." -ForegroundColor Yellow