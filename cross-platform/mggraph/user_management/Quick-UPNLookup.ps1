#Requires -Modules Microsoft.Graph.Users

<#
.SYNOPSIS
    Quick UPN lookup - paste names directly when prompted
    
.DESCRIPTION
    Simple script that prompts you to paste a list of names and automatically
    looks up their UPNs. Perfect for quick lookups from emails or documents.
    
.PARAMETER ExactMatch
    Require exact display name match (default: true)
    
.EXAMPLE
    .\Quick-UPNLookup.ps1
    
.EXAMPLE
    .\Quick-UPNLookup.ps1 -ExactMatch:$false
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory=$false)]
    [bool]$ExactMatch = $true
)

# Connect to Microsoft Graph if not already connected
try {
    $context = Get-MgContext
    if (-not $context) {
        Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Yellow
        Connect-MgGraph -Scopes "User.Read.All" -NoWelcome
        Write-Host "Connected successfully" -ForegroundColor Green
    }
} catch {
    Write-Error "Failed to connect to Microsoft Graph: $_"
    exit 1
}

Write-Host "`n=== QUICK UPN LOOKUP ===" -ForegroundColor Cyan
Write-Host "Paste your list of names below and press Enter twice when done:" -ForegroundColor Yellow
Write-Host "(Names can be separated by newlines, commas, or semicolons)" -ForegroundColor Gray
Write-Host ""

# Collect input until empty line
$inputLines = @()
do {
    $line = Read-Host
    if ($line.Trim() -ne "") {
        $inputLines += $line
    }
} while ($line.Trim() -ne "")

if ($inputLines.Count -eq 0) {
    Write-Host "No names provided. Exiting." -ForegroundColor Red
    exit 0
}

# Parse the input
$allText = $inputLines -join "`n"
$names = $allText -split "`n|`r`n|,|;" | 
    ForEach-Object { $_.Trim() } | 
    Where-Object { $_ -ne "" -and $_ -notmatch "^\s*$" }

Write-Host "`nParsed $($names.Count) names:" -ForegroundColor Green
$names | ForEach-Object { Write-Host "  - $_" -ForegroundColor White }

Write-Host "`nSearching..." -ForegroundColor Yellow

# Results
$results = @()
$found = 0
$notFound = 0

foreach ($name in $names) {
    Write-Host "`nLooking up: $name" -ForegroundColor Cyan
    
    try {
        if ($ExactMatch) {
            $users = Get-MgUser -Filter "DisplayName eq '$name'" -ConsistencyLevel eventual -ErrorAction SilentlyContinue
        } else {
            $users = Get-MgUser -Filter "startswith(DisplayName,'$name')" -ConsistencyLevel eventual -ErrorAction SilentlyContinue
            if (-not $users) {
                $allUsers = Get-MgUser -All -Property DisplayName,UserPrincipalName -ErrorAction SilentlyContinue
                $users = $allUsers | Where-Object { $_.DisplayName -like "*$name*" }
            }
        }
        
        if (-not $users) {
            Write-Host "  ✗ NOT FOUND" -ForegroundColor Red
            $notFound++
            $results += "$name -> NOT FOUND"
        }
        elseif ($users.Count -eq 1) {
            Write-Host "  ✓ $($users[0].UserPrincipalName)" -ForegroundColor Green
            $found++
            $results += "$name -> $($users[0].UserPrincipalName)"
        }
        else {
            Write-Host "  ⚠ Multiple matches:" -ForegroundColor Yellow
            foreach ($user in $users) {
                Write-Host "    - $($user.DisplayName) -> $($user.UserPrincipalName)" -ForegroundColor Gray
                $results += "$name -> $($user.UserPrincipalName) (multiple matches)"
            }
            $found += $users.Count
        }
    }
    catch {
        Write-Host "  ✗ ERROR: $_" -ForegroundColor Red
        $notFound++
        $results += "$name -> ERROR"
    }
}

# Summary
Write-Host "`n=== RESULTS SUMMARY ===" -ForegroundColor Cyan
Write-Host "Found: $found" -ForegroundColor Green
Write-Host "Not Found: $notFound" -ForegroundColor Red

Write-Host "`n=== COPY/PASTE RESULTS ===" -ForegroundColor Cyan
$results | ForEach-Object { Write-Host $_ -ForegroundColor White }

# Option to save results
Write-Host "`nSave results to file? (Y/N): " -ForegroundColor Yellow
$save = Read-Host
if ($save -eq 'Y' -or $save -eq 'y') {
    $timestamp = Get-Date -Format "yyyy-MM-dd-HHmm"
    $filename = "UPN-Results-$timestamp.txt"
    $results | Out-File -FilePath $filename -Encoding UTF8
    Write-Host "Results saved to: $filename" -ForegroundColor Green
}

Write-Host "`nLookup complete!" -ForegroundColor Green