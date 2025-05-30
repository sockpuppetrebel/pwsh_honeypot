# Connect to Microsoft Graph
Connect-MgGraph -Scopes "User.Read.All" -UseDeviceAuthentication -NoWelcome

# Missing users with potential variations
$missingUsers = @(
    @{
        Name = "Matthew Payne"
        FirstNameVariations = @("Matthew", "Matt")
        LastName = "Payne"
    },
    @{
        Name = "Marcus Hoffman"  
        FirstNameVariations = @("Marcus")
        LastName = @("Hoffman", "Hoffmann")
    },
    @{
        Name = "Rob Saunders"
        FirstNameVariations = @("Rob", "Robert", "Bob")
        LastName = "Saunders"
    },
    @{
        Name = "Zachary Coulter"
        FirstNameVariations = @("Zachary", "Zach", "Zack")
        LastName = "Coulter"
    },
    @{
        Name = "Terry McGregor"
        FirstNameVariations = @("Terry", "Terrence", "Terence")
        LastName = @("McGregor", "MacGregor")
    }
)

Write-Host "🔍 Searching for missing UPNs with flexible name matching..." -ForegroundColor Yellow
Write-Host ""

foreach ($user in $missingUsers) {
    Write-Host "Searching for: $($user.Name)" -ForegroundColor Cyan
    $found = $false
    
    # Search by first name variations + last name
    foreach ($firstName in $user.FirstNameVariations) {
        if ($user.LastName -is [array]) {
            $lastNames = $user.LastName
        } else {
            $lastNames = @($user.LastName)
        }
        
        foreach ($lastName in $lastNames) {
            # Try exact display name match
            $fullName = "$firstName $lastName"
            $result = Get-MgUser -Filter "DisplayName eq '$fullName'" -ConsistencyLevel eventual -ErrorAction SilentlyContinue
            
            if ($result) {
                Write-Host "  ✔ Found: $($result.DisplayName) -> $($result.UserPrincipalName)" -ForegroundColor Green
                $found = $true
                break
            }
            
            # Try partial matches using contains
            $result = Get-MgUser -Filter "startswith(DisplayName,'$firstName') and contains(DisplayName,'$lastName')" -ConsistencyLevel eventual -ErrorAction SilentlyContinue
            
            if ($result) {
                foreach ($r in $result) {
                    Write-Host "  ✔ Found: $($r.DisplayName) -> $($r.UserPrincipalName)" -ForegroundColor Green
                    $found = $true
                }
                break
            }
        }
        if ($found) { break }
    }
    
    # If still not found, try searching by last name only
    if (-not $found) {
        if ($user.LastName -is [array]) {
            $lastNames = $user.LastName
        } else {
            $lastNames = @($user.LastName)
        }
        
        foreach ($lastName in $lastNames) {
            $result = Get-MgUser -Filter "contains(DisplayName,'$lastName')" -ConsistencyLevel eventual -ErrorAction SilentlyContinue
            
            if ($result) {
                Write-Host "  📋 Potential matches for '$lastName':" -ForegroundColor Yellow
                foreach ($r in $result) {
                    Write-Host "     $($r.DisplayName) -> $($r.UserPrincipalName)" -ForegroundColor Yellow
                }
                $found = $true
                break
            }
        }
    }
    
    if (-not $found) {
        Write-Host "  ❌ No matches found for $($user.Name)" -ForegroundColor Red
    }
    
    Write-Host ""
}

Write-Host "🔍 Search complete!" -ForegroundColor Green