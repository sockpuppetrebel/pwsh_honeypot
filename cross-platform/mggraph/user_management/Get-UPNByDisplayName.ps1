#Requires -Modules Microsoft.Graph.Users

<#
.SYNOPSIS
    Retrieves User Principal Names (UPNs) based on display names
    
.DESCRIPTION
    This script looks up users in Azure AD/Entra ID by their display names and returns
    their corresponding User Principal Names (email addresses). Supports both single
    user lookup and batch processing from a file.
    
.PARAMETER DisplayName
    Single display name to look up
    
.PARAMETER DisplayNames
    Array of display names to look up
    
.PARAMETER InputFile
    Path to text file containing display names (one per line)
    
.PARAMETER RawText
    Raw text block containing names (paste directly from email/document)
    Automatically parses names separated by newlines, commas, or semicolons
    
.PARAMETER ExportToCsv
    Export results to CSV file
    
.PARAMETER ExactMatch
    Require exact display name match (default uses contains)
    
.EXAMPLE
    .\Get-UPNByDisplayName.ps1 -DisplayName "John Smith"
    
.EXAMPLE
    .\Get-UPNByDisplayName.ps1 -DisplayNames @("John Smith", "Jane Doe")
    
.EXAMPLE
    .\Get-UPNByDisplayName.ps1 -InputFile "names.txt" -ExportToCsv
    
.EXAMPLE
    .\Get-UPNByDisplayName.ps1 -DisplayName "John" -ExactMatch:$false
    
.EXAMPLE
    .\Get-UPNByDisplayName.ps1 -RawText "John Smith
    Jane Doe
    Bob Johnson"
    
.EXAMPLE
    .\Get-UPNByDisplayName.ps1 -RawText "John Smith, Jane Doe, Bob Johnson"
    
.NOTES
    Author: System Administrator
    Version: 1.0
    Requires: Microsoft Graph PowerShell SDK with User.Read.All permission
#>

[CmdletBinding(DefaultParameterSetName = "Single")]
param(
    [Parameter(Mandatory=$true, ParameterSetName = "Single")]
    [string]$DisplayName,
    
    [Parameter(Mandatory=$true, ParameterSetName = "Multiple")]
    [string[]]$DisplayNames,
    
    [Parameter(Mandatory=$true, ParameterSetName = "File")]
    [string]$InputFile,
    
    [Parameter(Mandatory=$true, ParameterSetName = "RawText")]
    [string]$RawText,
    
    [Parameter(Mandatory=$false)]
    [switch]$ExportToCsv,
    
    [Parameter(Mandatory=$false)]
    [bool]$ExactMatch = $true
)

# Function to write colored output
function Write-ColorOutput {
    param(
        [string]$Message,
        [string]$ForegroundColor = "White"
    )
    Write-Host $Message -ForegroundColor $ForegroundColor
}

# Function to search for user by display name
function Get-UserByDisplayName {
    param(
        [string]$Name,
        [bool]$Exact = $true
    )
    
    try {
        if ($Exact) {
            # Exact match search
            $users = Get-MgUser -Filter "DisplayName eq '$Name'" -ConsistencyLevel eventual -ErrorAction SilentlyContinue
        } else {
            # Contains search (less precise but more flexible)
            $users = Get-MgUser -Filter "startswith(DisplayName,'$Name')" -ConsistencyLevel eventual -ErrorAction SilentlyContinue
            if (-not $users) {
                # Try contains search if startswith fails
                $allUsers = Get-MgUser -All -Property DisplayName,UserPrincipalName -ErrorAction SilentlyContinue
                $users = $allUsers | Where-Object { $_.DisplayName -like "*$Name*" }
            }
        }
        
        return $users
    }
    catch {
        Write-Warning "Error searching for '$Name': $($_.Exception.Message)"
        return $null
    }
}

# Connect to Microsoft Graph if not already connected
try {
    $context = Get-MgContext
    if (-not $context) {
        Write-ColorOutput "Connecting to Microsoft Graph..." -ForegroundColor Yellow
        Connect-MgGraph -Scopes "User.Read.All" -NoWelcome
        Write-ColorOutput "Connected successfully" -ForegroundColor Green
    } else {
        Write-ColorOutput "Already connected to Microsoft Graph" -ForegroundColor Green
    }
} catch {
    Write-Error "Failed to connect to Microsoft Graph: $_"
    exit 1
}

Write-ColorOutput "`n=== UPN LOOKUP BY DISPLAY NAME ===" -ForegroundColor Cyan

# Determine which names to process
$namesToProcess = @()

switch ($PSCmdlet.ParameterSetName) {
    "Single" {
        $namesToProcess = @($DisplayName)
    }
    "Multiple" {
        $namesToProcess = $DisplayNames
    }
    "File" {
        if (-not (Test-Path $InputFile)) {
            Write-Error "Input file not found: $InputFile"
            exit 1
        }
        $namesToProcess = Get-Content $InputFile | Where-Object { $_.Trim() -ne "" }
        Write-ColorOutput "Loaded $($namesToProcess.Count) names from file" -ForegroundColor White
    }
    "RawText" {
        # Parse raw text - handle multiple separators
        $rawNames = $RawText -split "`n|`r`n|,|;" | 
            ForEach-Object { $_.Trim() } | 
            Where-Object { $_ -ne "" -and $_ -notmatch "^\s*$" }
        
        $namesToProcess = $rawNames
        Write-ColorOutput "Parsed $($namesToProcess.Count) names from raw text" -ForegroundColor White
        Write-ColorOutput "Names found: $($namesToProcess -join ', ')" -ForegroundColor Gray
    }
}

# Results collection
$results = @()
$foundCount = 0
$notFoundCount = 0
$multipleMatchCount = 0

Write-ColorOutput "`nSearching for $($namesToProcess.Count) user(s)..." -ForegroundColor Yellow
Write-ColorOutput "Match Type: $(if ($ExactMatch) { 'Exact' } else { 'Contains' })" -ForegroundColor Gray

foreach ($name in $namesToProcess) {
    Write-ColorOutput "`nSearching: $name" -ForegroundColor Cyan
    
    $users = Get-UserByDisplayName -Name $name -Exact $ExactMatch
    
    if (-not $users) {
        Write-ColorOutput "  ✗ Not found" -ForegroundColor Red
        $notFoundCount++
        $results += [PSCustomObject]@{
            SearchName = $name
            DisplayName = "NOT FOUND"
            UserPrincipalName = "NOT FOUND"
            Status = "Not Found"
            MatchCount = 0
        }
    }
    elseif ($users.Count -eq 1) {
        Write-ColorOutput "  ✓ Found: $($users[0].DisplayName)" -ForegroundColor Green
        Write-ColorOutput "    UPN: $($users[0].UserPrincipalName)" -ForegroundColor White
        $foundCount++
        $results += [PSCustomObject]@{
            SearchName = $name
            DisplayName = $users[0].DisplayName
            UserPrincipalName = $users[0].UserPrincipalName
            Status = "Found"
            MatchCount = 1
        }
    }
    else {
        Write-ColorOutput "  ⚠ Multiple matches found ($($users.Count)):" -ForegroundColor Yellow
        $multipleMatchCount++
        
        foreach ($user in $users) {
            Write-ColorOutput "    - $($user.DisplayName) ($($user.UserPrincipalName))" -ForegroundColor Gray
            $results += [PSCustomObject]@{
                SearchName = $name
                DisplayName = $user.DisplayName
                UserPrincipalName = $user.UserPrincipalName
                Status = "Multiple Matches"
                MatchCount = $users.Count
            }
        }
    }
}

# Display summary
Write-ColorOutput "`n=== SUMMARY ===" -ForegroundColor Cyan
Write-ColorOutput "Names searched: $($namesToProcess.Count)" -ForegroundColor White
Write-ColorOutput "Found (unique): $foundCount" -ForegroundColor Green
Write-ColorOutput "Not found: $notFoundCount" -ForegroundColor Red
Write-ColorOutput "Multiple matches: $multipleMatchCount" -ForegroundColor Yellow

# Export results if requested
if ($ExportToCsv) {
    $timestamp = Get-Date -Format "yyyy-MM-dd-HHmm"
    $csvPath = "UPN-Lookup-Results-$timestamp.csv"
    $results | Export-Csv -Path $csvPath -NoTypeInformation
    Write-ColorOutput "`nResults exported to: $csvPath" -ForegroundColor Green
}

# Show not found names for easy copy/paste
$notFoundNames = $results | Where-Object { $_.Status -eq "Not Found" } | Select-Object -ExpandProperty SearchName
if ($notFoundNames.Count -gt 0) {
    Write-ColorOutput "`nNot Found Names (for easy copy/paste):" -ForegroundColor Yellow
    $notFoundNames | ForEach-Object { Write-ColorOutput "  $_" -ForegroundColor Red }
}

# Disconnect if we connected in this script
if (-not $context) {
    Disconnect-MgGraph -ErrorAction SilentlyContinue
}