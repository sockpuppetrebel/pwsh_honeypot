# Smart search - prioritize likely sites based on patterns
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

$store = New-Object System.Security.Cryptography.X509Certificates.X509Store("My", "CurrentUser")
$store.Open("ReadWrite")
$store.Add($cert)
$store.Close()

Write-Host "`n=============== SMART SITE SEARCH ===============" -ForegroundColor Cyan

# Connect to admin
Connect-PnPOnline -Url "https://episerver99-admin.sharepoint.com" `
                  -ClientId "fe2a9efe-3000-4b02-96ea-344a2583dd52" `
                  -Tenant "3ec00d79-021a-42d4-aac8-dcb35973dff2" `
                  -Thumbprint $cert.Thumbprint

# Get all sites
Write-Host "Getting all sites and analyzing patterns..." -ForegroundColor Yellow
$allSites = Get-PnPTenantSite | Where-Object { $_.Template -notlike "REDIRECT*" }

# Prioritize sites based on patterns
$prioritizedSites = $allSites | ForEach-Object {
    $priority = 5  # Default low priority
    
    # Higher priority for project/team sites
    if ($_.Url -match "NET|APAC|project|team|dev|tech|opp") { $priority = 1 }
    # Medium priority for general collaboration
    elseif ($_.Url -match "collab|work|share") { $priority = 2 }
    # Lower priority for corporate/HR sites
    elseif ($_.Url -match "HR|admin|corporate|policy") { $priority = 4 }
    # Skip personal sites
    elseif ($_.Url -match "-my\.sharepoint|personal") { $priority = 10 }
    
    [PSCustomObject]@{
        Url = $_.Url
        Title = $_.Title
        Priority = $priority
    }
} | Sort-Object Priority

Write-Host "Prioritized $($prioritizedSites.Count) sites" -ForegroundColor Green
Write-Host "`nTop priority sites:" -ForegroundColor Yellow
$prioritizedSites | Select-Object -First 20 | ForEach-Object {
    Write-Host "  [$($_.Priority)] $($_.Title) - $($_.Url)" -ForegroundColor White
}

Disconnect-PnPOnline

# Check top priority sites
Write-Host "`nChecking high-priority sites first..." -ForegroundColor Yellow
$findings = @()

$highPriority = $prioritizedSites | Where-Object { $_.Priority -le 2 } | Select-Object -First 100

foreach ($site in $highPriority) {
    Write-Host "`rChecking: $($site.Title)" -NoNewline
    
    try {
        Connect-PnPOnline -Url $site.Url `
                          -ClientId "fe2a9efe-3000-4b02-96ea-344a2583dd52" `
                          -Tenant "3ec00d79-021a-42d4-aac8-dcb35973dff2" `
                          -Thumbprint $cert.Thumbprint `
                          -ErrorAction Stop
        
        # Quick user check
        try {
            $user = Get-PnPUser -Identity $UserEmail -ErrorAction Stop
            
            # User exists, check groups
            $groups = Get-PnPGroup
            foreach ($group in $groups) {
                $members = Get-PnPGroupMember -Group $group -ErrorAction SilentlyContinue
                if ($members | Where-Object { $_.Email -eq $UserEmail }) {
                    $findings += [PSCustomObject]@{
                        SiteUrl = $site.Url
                        SiteTitle = $site.Title
                        GroupName = $group.Title
                    }
                    Write-Host "`n✓ FOUND in: $($site.Title) - Group: $($group.Title)" -ForegroundColor Green
                }
            }
        } catch {
            # User not in site
        }
        
        Disconnect-PnPOnline -ErrorAction SilentlyContinue
    } catch {
        # Skip
    }
}

Write-Host "`n`nFound $($findings.Count) additional sites with permissions" -ForegroundColor Cyan

if ($findings.Count -gt 0) {
    $csvPath = Join-Path $scriptPath "smart_search_results.csv"
    $findings | Export-Csv -Path $csvPath -NoTypeInformation
    Write-Host "Results saved to: $csvPath" -ForegroundColor Green
}

# Clean up
$store.Open("ReadWrite")
$certToRemove = $store.Certificates | Where-Object { $_.Thumbprint -eq $cert.Thumbprint }
if ($certToRemove) { $store.Remove($certToRemove) }
$store.Close()

Write-Host "`n✓ Complete!" -ForegroundColor Green