# ====== INTERACTIVE AUTH CONFIG ======
$tenantId = "3ec00d79-021a-42d4-aac8-dcb35973dff2"

# ====== CONNECT TO MICROSOFT GRAPH ======
try {
    Write-Host "🔐 Connecting to Microsoft Graph (device code login)..." -ForegroundColor Yellow
    Connect-MgGraph -TenantId $tenantId `
                    -Scopes "Group.ReadWrite.All", "User.Read.All" `
                    -UseDeviceAuthentication `
                    -NoWelcome
} catch {
    Write-Host "❌ Failed to connect to Microsoft Graph: $($_.Exception.Message)" -ForegroundColor Red
    return
}

# ====== VERIFY CONTEXT ======
$context = Get-MgContext
if (-not $context) {
    Write-Host "❌ Authentication failed - no valid context" -ForegroundColor Red
    return
}
Write-Host "✔ Connected to Microsoft Graph successfully" -ForegroundColor Green
Write-Host "ℹ Account: $($context.Account)" -ForegroundColor Cyan

# ====== LOOK UP GROUP ID ======
$groupName = "SG_Shareworks_Participant"
$group = Get-MgGroup -Filter "displayName eq '$groupName'"
if (-not $group) {
    Write-Host "❌ Group '$groupName' not found." -ForegroundColor Red
    return
}
$groupId = $group.Id
Write-Host "✔ Found group '$groupName' with ID: $groupId" -ForegroundColor Green

# ====== DEFINE UPN LIST ======
$upns = @(
    "jack.mcclean@optimizely.com",
    "paul.gray@optimizely.com",
    "william.lapuma@optimizely.com",
    "mark.ryan@optimizely.com",
    "david.bingham@optimizely.com",
    "anna.parback@optimizely.com",
    "jack.joseph@episerver.com",
    "thomas.mckenzie@optimizely.com",
    "sean.groat@optimizely.com",
    "brandon.halvorson@optimizely.com",
    "daniel.martell@optimizely.com",
    "nuno.figueiredo@optimizely.com",
    "marcus.hoffmann@optimizely.com",
    "zach.coulter@optimizely.com",
    "shannon.gray@optimizely.com",
    "anatoliy.savinov@optimizely.com",
    "brett.samuels@optimizely.com",
    "anna.redmile@optimizely.com",
    "aidan.dodd@optimizely.com",
    "vimi.kaul@optimizely.com",
    "mark.wakelin@optimizely.com",
    "chynna.roberts@optimizely.com",
    "alexandra.vanheel@optimizely.com",
    "jennifer.lovett@optimizely.com",
    "robin.leclerc@optimizely.com"
)

# ====== PREPARE CSV OUTPUT ======
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$csvPath = "group_membership_results_$timestamp.csv"
$results = @()

Write-Host "ℹ Results will be saved to: $csvPath" -ForegroundColor Cyan

# ====== ADD USERS TO GROUP ======
foreach ($upn in $upns) {
    $result = [PSCustomObject]@{
        UserPrincipalName = $upn
        Action = ""
        Status = ""
        ErrorMessage = ""
        Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    }
    
    try {
        $user = Get-MgUser -UserId $upn -ErrorAction Stop
        
        # Check if user is already a member
        $existingMember = Get-MgGroupMember -GroupId $groupId | Where-Object { $_.Id -eq $user.Id }
        if ($existingMember) {
            Write-Host "ℹ $upn is already a member of $groupName" -ForegroundColor Cyan
            $result.Action = "Skipped"
            $result.Status = "Already Member"
            continue
        }
        
        New-MgGroupMember -GroupId $groupId -DirectoryObjectId $user.Id -ErrorAction Stop
        Write-Host "✔ Added $upn to $groupName" -ForegroundColor Green
        $result.Action = "Added"
        $result.Status = "Success"
    } catch {
        if ($_.Exception.Message -like "*One or more added object references already exist*") {
            Write-Host "ℹ $upn is already a member of $groupName" -ForegroundColor Cyan
            $result.Action = "Skipped"
            $result.Status = "Already Member"
        } else {
            Write-Host "⚠ Failed to add ${upn}: $($_.Exception.Message)" -ForegroundColor Yellow
            $result.Action = "Failed"
            $result.Status = "Error"
            $result.ErrorMessage = $_.Exception.Message
        }
    }
    
    $results += $result
}

# ====== EXPORT RESULTS TO CSV ======
$results | Export-Csv -Path $csvPath -NoTypeInformation
Write-Host "`n📄 Results exported to: $csvPath" -ForegroundColor Green

# ====== SUMMARY ======
$added = ($results | Where-Object { $_.Action -eq "Added" }).Count
$skipped = ($results | Where-Object { $_.Action -eq "Skipped" }).Count
$failed = ($results | Where-Object { $_.Action -eq "Failed" }).Count

Write-Host "`n📊 SUMMARY:" -ForegroundColor Magenta
Write-Host "   ✔ Added: $added users" -ForegroundColor Green
Write-Host "   ℹ Skipped: $skipped users (already members)" -ForegroundColor Cyan
Write-Host "   ⚠ Failed: $failed users" -ForegroundColor Yellow
