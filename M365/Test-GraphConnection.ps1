# Test Microsoft Graph connection and permissions

Write-Host "Testing Microsoft Graph Connection..." -ForegroundColor Cyan

# Try to connect
try {
    Connect-MgGraph -Scopes "User.Read.All", "Directory.Read.All", "Organization.Read.All" -NoWelcome
    Write-Host "✓ Connected to Microsoft Graph" -ForegroundColor Green
} catch {
    Write-Error "Failed to connect: $_"
    exit 1
}

# Check current context
$context = Get-MgContext
Write-Host "`nCurrent Context:" -ForegroundColor Yellow
Write-Host "Account: $($context.Account)" -ForegroundColor White
Write-Host "TenantId: $($context.TenantId)" -ForegroundColor White
Write-Host "ClientId: $($context.ClientId)" -ForegroundColor White
Write-Host "AuthType: $($context.AuthType)" -ForegroundColor White

# Check scopes/permissions
Write-Host "`nGranted Scopes:" -ForegroundColor Yellow
$context.Scopes | ForEach-Object { Write-Host "  - $_" -ForegroundColor White }

# Test basic queries
Write-Host "`nTesting API Access:" -ForegroundColor Yellow

# Test 1: Get organization info
try {
    $org = Get-MgOrganization
    Write-Host "✓ Can read organization info: $($org.DisplayName)" -ForegroundColor Green
} catch {
    Write-Host "✗ Cannot read organization info: $_" -ForegroundColor Red
}

# Test 2: Get subscribed SKUs
try {
    $skus = Get-MgSubscribedSku -Top 1
    Write-Host "✓ Can read subscribed SKUs" -ForegroundColor Green
} catch {
    Write-Host "✗ Cannot read subscribed SKUs: $_" -ForegroundColor Red
}

# Test 3: Get users
try {
    $users = Get-MgUser -Top 1
    Write-Host "✓ Can read users" -ForegroundColor Green
} catch {
    Write-Host "✗ Cannot read users: $_" -ForegroundColor Red
}

Write-Host "`nDiagnostics complete!" -ForegroundColor Cyan

# Disconnect
Disconnect-MgGraph