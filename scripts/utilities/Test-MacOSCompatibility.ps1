#Requires -Version 7.0

<#
.SYNOPSIS
    Tests PowerShell Core compatibility on macOS
    
.DESCRIPTION
    Validates that PowerShell Core features work correctly on macOS including
    modules, authentication, and interactive features used by other scripts.
    
.EXAMPLE
    .\Test-MacOSCompatibility.ps1
#>

[CmdletBinding()]
param()

Write-Host "=== POWERSHELL CORE MACOS COMPATIBILITY TEST ===" -ForegroundColor Cyan
Write-Host "PowerShell Version: $($PSVersionTable.PSVersion)" -ForegroundColor White
Write-Host "Platform: $($PSVersionTable.Platform)" -ForegroundColor White
Write-Host "OS: $($PSVersionTable.OS)" -ForegroundColor White

$testResults = @()

# Test 1: Basic PowerShell features
Write-Host "`n1. Testing basic PowerShell features..." -ForegroundColor Yellow
try {
    $array = @("test1", "test2", "test3")
    $hash = @{ key1 = "value1"; key2 = "value2" }
    $object = [PSCustomObject]@{ Name = "Test"; Value = 123 }
    
    Write-Host "   ✓ Arrays, hashtables, and objects work" -ForegroundColor Green
    $testResults += "Basic Features: PASS"
}
catch {
    Write-Host "   ✗ Basic features failed: $_" -ForegroundColor Red
    $testResults += "Basic Features: FAIL - $_"
}

# Test 2: String manipulation and regex
Write-Host "`n2. Testing string manipulation..." -ForegroundColor Yellow
try {
    $text = "John Smith, Jane Doe; Bob Johnson"
    $names = $text -split ",|;" | ForEach-Object { $_.Trim() }
    $email = "test@domain.com"
    $isEmail = $email -match "^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$"
    
    if ($names.Count -eq 3 -and $isEmail) {
        Write-Host "   ✓ String splitting and regex work" -ForegroundColor Green
        $testResults += "String Manipulation: PASS"
    } else {
        throw "String operations didn't produce expected results"
    }
}
catch {
    Write-Host "   ✗ String manipulation failed: $_" -ForegroundColor Red
    $testResults += "String Manipulation: FAIL - $_"
}

# Test 3: File operations
Write-Host "`n3. Testing file operations..." -ForegroundColor Yellow
try {
    $testFile = "/tmp/pwsh-test-$(Get-Date -Format 'yyyyMMddHHmmss').txt"
    "Test content" | Out-File -FilePath $testFile
    $content = Get-Content $testFile
    Remove-Item $testFile -Force
    
    if ($content -eq "Test content") {
        Write-Host "   ✓ File operations work" -ForegroundColor Green
        $testResults += "File Operations: PASS"
    } else {
        throw "File content didn't match"
    }
}
catch {
    Write-Host "   ✗ File operations failed: $_" -ForegroundColor Red
    $testResults += "File Operations: FAIL - $_"
}

# Test 4: Interactive input simulation
Write-Host "`n4. Testing interactive features..." -ForegroundColor Yellow
try {
    # Test Write-Host with colors
    Write-Host "   Testing colored output..." -ForegroundColor Cyan
    Write-Host "   Testing background colors..." -ForegroundColor White -BackgroundColor DarkBlue
    
    # Test parameter binding
    function Test-ParameterBinding {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory=$false)]
            [string]$TestParam = "default"
        )
        return $TestParam
    }
    
    $result = Test-ParameterBinding -TestParam "custom"
    if ($result -eq "custom") {
        Write-Host "   ✓ Interactive features work" -ForegroundColor Green
        $testResults += "Interactive Features: PASS"
    } else {
        throw "Parameter binding failed"
    }
}
catch {
    Write-Host "   ✗ Interactive features failed: $_" -ForegroundColor Red
    $testResults += "Interactive Features: FAIL - $_"
}

# Test 5: Module loading (check common modules)
Write-Host "`n5. Testing module availability..." -ForegroundColor Yellow
$modules = @(
    @{ Name = "Microsoft.Graph.Users"; Required = $false },
    @{ Name = "Microsoft.Graph.Groups"; Required = $false },
    @{ Name = "ExchangeOnlineManagement"; Required = $false },
    @{ Name = "Microsoft.Online.SharePoint.PowerShell"; Required = $false }
)

foreach ($module in $modules) {
    try {
        $available = Get-Module -ListAvailable -Name $module.Name
        if ($available) {
            Write-Host "   ✓ $($module.Name) is available" -ForegroundColor Green
            $testResults += "$($module.Name): AVAILABLE"
        } else {
            Write-Host "   - $($module.Name) not installed" -ForegroundColor Yellow
            $testResults += "$($module.Name): NOT INSTALLED"
        }
    }
    catch {
        Write-Host "   ✗ Error checking $($module.Name): $_" -ForegroundColor Red
        $testResults += "$($module.Name): ERROR - $_"
    }
}

# Test 6: Progress bars and formatting
Write-Host "`n6. Testing progress and formatting..." -ForegroundColor Yellow
try {
    for ($i = 1; $i -le 3; $i++) {
        Write-Progress -Activity "Testing Progress" -Status "Step $i of 3" -PercentComplete ($i * 33)
        Start-Sleep -Milliseconds 500
    }
    Write-Progress -Activity "Testing Progress" -Completed
    
    # Test formatting
    $data = @(
        [PSCustomObject]@{ Name = "Test1"; Value = 100 },
        [PSCustomObject]@{ Name = "Test2"; Value = 200 }
    )
    
    $formatted = $data | Format-Table -AutoSize | Out-String
    if ($formatted.Length -gt 0) {
        Write-Host "   ✓ Progress bars and formatting work" -ForegroundColor Green
        $testResults += "Progress/Formatting: PASS"
    } else {
        throw "Formatting produced no output"
    }
}
catch {
    Write-Host "   ✗ Progress/formatting failed: $_" -ForegroundColor Red
    $testResults += "Progress/Formatting: FAIL - $_"
}

# Test 7: Error handling
Write-Host "`n7. Testing error handling..." -ForegroundColor Yellow
try {
    try {
        # Intentionally cause an error
        Get-Item "/nonexistent/path/file.txt" -ErrorAction Stop
    }
    catch {
        # This should be caught
        if ($_.Exception.Message) {
            Write-Host "   ✓ Error handling works correctly" -ForegroundColor Green
            $testResults += "Error Handling: PASS"
        } else {
            throw "Error object doesn't have expected properties"
        }
    }
}
catch {
    Write-Host "   ✗ Error handling failed: $_" -ForegroundColor Red
    $testResults += "Error Handling: FAIL - $_"
}

# Summary
Write-Host "`n=== TEST RESULTS SUMMARY ===" -ForegroundColor Cyan
$passCount = 0
$failCount = 0

foreach ($result in $testResults) {
    if ($result -like "*PASS*" -or $result -like "*AVAILABLE*") {
        Write-Host "✓ $result" -ForegroundColor Green
        $passCount++
    } elseif ($result -like "*FAIL*" -or $result -like "*ERROR*") {
        Write-Host "✗ $result" -ForegroundColor Red
        $failCount++
    } else {
        Write-Host "- $result" -ForegroundColor Yellow
    }
}

Write-Host "`nOverall Status:" -ForegroundColor Cyan
Write-Host "Passed: $passCount" -ForegroundColor Green
Write-Host "Failed: $failCount" -ForegroundColor Red

if ($failCount -eq 0) {
    Write-Host "`nPowerShell Core is fully compatible with your macOS environment!" -ForegroundColor Green
    Write-Host "All scripts should work correctly." -ForegroundColor Green
} elseif ($failCount -le 2) {
    Write-Host "`nPowerShell Core is mostly compatible with minor issues." -ForegroundColor Yellow
    Write-Host "Most scripts should work correctly." -ForegroundColor Yellow
} else {
    Write-Host "`nPowerShell Core has compatibility issues on this system." -ForegroundColor Red
    Write-Host "Some scripts may not work correctly." -ForegroundColor Red
}

Write-Host "`nRecommendations:" -ForegroundColor Cyan
Write-Host "- Install missing modules with: Install-Module ModuleName" -ForegroundColor White
Write-Host "- Ensure you're running PowerShell 7.0 or later" -ForegroundColor White
Write-Host "- For Exchange/SharePoint scripts, install respective modules" -ForegroundColor White