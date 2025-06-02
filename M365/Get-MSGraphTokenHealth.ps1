#Requires -Modules Microsoft.Graph.Authentication

<#
.SYNOPSIS
    Monitors and validates Microsoft Graph API token health and permissions
.DESCRIPTION
    Checks Graph API connectivity, token expiration, permission scopes, and
    rate limiting status. Essential for troubleshooting Graph-based automation.
.PARAMETER CheckAllConnections
    Test connectivity to all commonly used Graph endpoints
.PARAMETER ValidatePermissions
    Check if current token has required permissions for common operations
.EXAMPLE
    .\Get-MSGraphTokenHealth.ps1 -CheckAllConnections -ValidatePermissions
#>

param(
    [Parameter(Mandatory = $false)]
    [switch]$CheckAllConnections,
    
    [Parameter(Mandatory = $false)]
    [switch]$ValidatePermissions
)

Write-Host "=== MICROSOFT GRAPH TOKEN HEALTH CHECK ===" -ForegroundColor Cyan

# Check if already connected
$context = Get-MgContext

if (-not $context) {
    Write-Host "Not connected to Microsoft Graph. Attempting connection..." -ForegroundColor Yellow
    try {
        Connect-MgGraph -Scopes "User.Read.All", "Directory.Read.All"
        $context = Get-MgContext
    }
    catch {
        Write-Error "Failed to connect to Microsoft Graph: $($_.Exception.Message)"
        exit 1
    }
}

Write-Host "✓ Connected to Microsoft Graph" -ForegroundColor Green

# Display current context information
Write-Host "`n--- Current Graph Context ---" -ForegroundColor Yellow
Write-Host "Tenant ID: $($context.TenantId)" -ForegroundColor White
Write-Host "Client ID: $($context.ClientId)" -ForegroundColor White
Write-Host "Account: $($context.Account)" -ForegroundColor White
Write-Host "Auth Type: $($context.AuthType)" -ForegroundColor White
Write-Host "Token Expires: $($context.TokenCredential.Token.ExpiresOn)" -ForegroundColor White

# Check token expiration
$tokenExpiry = $context.TokenCredential.Token.ExpiresOn
$timeUntilExpiry = $tokenExpiry - (Get-Date)

if ($timeUntilExpiry.TotalMinutes -lt 5) {
    Write-Host "⚠ WARNING: Token expires in $([math]::Round($timeUntilExpiry.TotalMinutes,1)) minutes" -ForegroundColor Red
} elseif ($timeUntilExpiry.TotalMinutes -lt 15) {
    Write-Host "⚠ CAUTION: Token expires in $([math]::Round($timeUntilExpiry.TotalMinutes,1)) minutes" -ForegroundColor Yellow
} else {
    Write-Host "✓ Token valid for $([math]::Round($timeUntilExpiry.TotalHours,1)) hours" -ForegroundColor Green
}

# Display current scopes
Write-Host "`n--- Granted Scopes ---" -ForegroundColor Yellow
$scopes = $context.Scopes
if ($scopes) {
    $scopes | Sort-Object | ForEach-Object { Write-Host "  $_" -ForegroundColor White }
} else {
    Write-Host "  No scopes information available" -ForegroundColor Gray
}

# Test basic Graph connectivity
Write-Host "`n--- Testing Basic Connectivity ---" -ForegroundColor Yellow

$connectivityTests = @(
    @{ Name = "Current User"; Command = { Get-MgUser -UserId (Get-MgContext).Account -Select "Id,DisplayName" } },
    @{ Name = "Organization Info"; Command = { Get-MgOrganization -Select "Id,DisplayName" } },
    @{ Name = "Directory Stats"; Command = { Get-MgDirectoryStatistic } }
)

foreach ($test in $connectivityTests) {
    try {
        $result = & $test.Command
        Write-Host "✓ $($test.Name): OK" -ForegroundColor Green
    }
    catch {
        Write-Host "✗ $($test.Name): Failed - $($_.Exception.Message)" -ForegroundColor Red
    }
}

# Extended connectivity tests
if ($CheckAllConnections) {
    Write-Host "`n--- Extended Connectivity Tests ---" -ForegroundColor Yellow
    
    $extendedTests = @(
        @{ Name = "Users"; Command = { Get-MgUser -Top 1 -Select "Id" } },
        @{ Name = "Groups"; Command = { Get-MgGroup -Top 1 -Select "Id" } },
        @{ Name = "Applications"; Command = { Get-MgApplication -Top 1 -Select "Id" } },
        @{ Name = "Service Principals"; Command = { Get-MgServicePrincipal -Top 1 -Select "Id" } },
        @{ Name = "Devices"; Command = { Get-MgDevice -Top 1 -Select "Id" } },
        @{ Name = "Directory Roles"; Command = { Get-MgDirectoryRole -Top 1 -Select "Id" } }
    )
    
    foreach ($test in $extendedTests) {
        try {
            $startTime = Get-Date
            $result = & $test.Command
            $duration = (Get-Date) - $startTime
            Write-Host "✓ $($test.Name): OK ($([math]::Round($duration.TotalMilliseconds))ms)" -ForegroundColor Green
        }
        catch {
            $errorMessage = $_.Exception.Message
            if ($errorMessage -like "*Insufficient privileges*") {
                Write-Host "⚠ $($test.Name): Insufficient permissions" -ForegroundColor Yellow
            } elseif ($errorMessage -like "*throttled*" -or $errorMessage -like "*rate limit*") {
                Write-Host "⚠ $($test.Name): Rate limited" -ForegroundColor Yellow
            } else {
                Write-Host "✗ $($test.Name): Failed - $errorMessage" -ForegroundColor Red
            }
        }
    }
}

# Permission validation
if ($ValidatePermissions) {
    Write-Host "`n--- Permission Validation ---" -ForegroundColor Yellow
    
    $permissionTests = @(
        @{ 
            Name = "Read Users"
            Scope = "User.Read.All"
            Test = { Get-MgUser -Top 1 -Select "Id" }
        },
        @{ 
            Name = "Read Groups"
            Scope = "Group.Read.All"
            Test = { Get-MgGroup -Top 1 -Select "Id" }
        },
        @{ 
            Name = "Read Directory"
            Scope = "Directory.Read.All"
            Test = { Get-MgDirectoryRole -Top 1 -Select "Id" }
        },
        @{ 
            Name = "Read Devices"
            Scope = "Device.Read.All"
            Test = { Get-MgDevice -Top 1 -Select "Id" }
        },
        @{ 
            Name = "Read Applications"
            Scope = "Application.Read.All"
            Test = { Get-MgApplication -Top 1 -Select "Id" }
        }
    )
    
    foreach ($permTest in $permissionTests) {
        $hasScope = $scopes -contains $permTest.Scope
        
        if ($hasScope) {
            try {
                & $permTest.Test | Out-Null
                Write-Host "✓ $($permTest.Name): Permission granted and functional" -ForegroundColor Green
            }
            catch {
                Write-Host "⚠ $($permTest.Name): Permission granted but test failed" -ForegroundColor Yellow
            }
        } else {
            Write-Host "✗ $($permTest.Name): Missing scope '$($permTest.Scope)'" -ForegroundColor Red
        }
    }
}

# Rate limiting information
Write-Host "`n--- Rate Limiting Status ---" -ForegroundColor Yellow

try {
    $startTime = Get-Date
    
    # Make several quick requests to test rate limiting
    for ($i = 1; $i -le 5; $i++) {
        Get-MgUser -Top 1 -Select "Id" | Out-Null
        Start-Sleep -Milliseconds 100
    }
    
    $totalTime = (Get-Date) - $startTime
    $avgResponseTime = $totalTime.TotalMilliseconds / 5
    
    if ($avgResponseTime -lt 200) {
        Write-Host "✓ Response times normal (avg: $([math]::Round($avgResponseTime))ms)" -ForegroundColor Green
    } elseif ($avgResponseTime -lt 500) {
        Write-Host "⚠ Response times elevated (avg: $([math]::Round($avgResponseTime))ms)" -ForegroundColor Yellow
    } else {
        Write-Host "⚠ Response times high (avg: $([math]::Round($avgResponseTime))ms) - possible throttling" -ForegroundColor Red
    }
}
catch {
    if ($_.Exception.Message -like "*throttled*" -or $_.Exception.Message -like "*rate limit*") {
        Write-Host "⚠ Currently being rate limited" -ForegroundColor Red
    } else {
        Write-Host "✗ Error testing rate limits: $($_.Exception.Message)" -ForegroundColor Red
    }
}

# Recommendations
Write-Host "`n--- Recommendations ---" -ForegroundColor Yellow

if ($timeUntilExpiry.TotalMinutes -lt 15) {
    Write-Host "• Refresh your token soon to avoid authentication errors" -ForegroundColor Cyan
}

if ($scopes -notcontains "User.Read.All") {
    Write-Host "• Consider adding User.Read.All scope for user management tasks" -ForegroundColor Cyan
}

if ($scopes -notcontains "Group.Read.All") {
    Write-Host "• Consider adding Group.Read.All scope for group management tasks" -ForegroundColor Cyan
}

Write-Host "• Use certificate-based authentication for production automation" -ForegroundColor Cyan
Write-Host "• Monitor rate limits and implement exponential backoff in scripts" -ForegroundColor Cyan
Write-Host "• Cache results when possible to reduce API calls" -ForegroundColor Cyan

# Token refresh suggestion
Write-Host "`n--- Token Management ---" -ForegroundColor Yellow
Write-Host "To refresh token: Disconnect-MgGraph; Connect-MgGraph -Scopes <required-scopes>" -ForegroundColor Cyan
Write-Host "To check permissions: Get-MgContext | Select-Object Scopes" -ForegroundColor Cyan

Write-Host "`n=== HEALTH CHECK COMPLETE ===" -ForegroundColor Cyan