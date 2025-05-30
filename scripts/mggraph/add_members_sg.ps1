# ====== CERTIFICATE AUTH CONFIG ======
$tenantId = "3ec00d79-021a-42d4-aac8-dcb35973dff2"
$clientId = "fe2a9efe-3000-4b02-96ea-344a2583dd52"
$pfxPath = "$env:USERPROFILE\azureappauth.pfx"

# ====== LOAD CERTIFICATE ======
if (-not (Test-Path $pfxPath)) {
    Write-Host "❌ Certificate file not found: $pfxPath" -ForegroundColor Red
    return
}

try {
    $cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2
    $cert.Import($pfxPath, $null, "Exportable,PersistKeySet")
} catch {
    Write-Host "❌ Failed to load certificate: $($_.Exception.Message)" -ForegroundColor Red
    return
}

# ====== CONNECT TO MICROSOFT GRAPH ======
try {
    Connect-MgGraph -TenantId $tenantId `
                    -ClientId $clientId `
                    -Certificate $cert `
                    -NoWelcome
} catch {
    Write-Host "❌ Failed to connect to Microsoft Graph: $($_.Exception.Message)" -ForegroundColor Red
    return
}

# ====== VERIFY CONTEXT ======
$context = Get-MgContext
if ($context.AuthType -ne "AppOnly") {
    Write-Host "❌ Authentication failed. Expected AppOnly, got: $($context.AuthType)" -ForegroundColor Red
    return
}
Write-Host "✔ Connected to Microsoft Graph with AppOnly authentication" -ForegroundColor Green

# ====== LOOK UP GROUP ID ======
$groupName = "sg_shareworks"
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

# ====== ADD USERS TO GROUP ======
foreach ($upn in $upns) {
    try {
        $user = Get-MgUser -UserId $upn -ErrorAction Stop
        
        # Check if user is already a member
        $existingMember = Get-MgGroupMember -GroupId $groupId | Where-Object { $_.Id -eq $user.Id }
        if ($existingMember) {
            Write-Host "ℹ $upn is already a member of $groupName" -ForegroundColor Cyan
            continue
        }
        
        Add-MgGroupMember -GroupId $groupId -DirectoryObjectId $user.Id -ErrorAction Stop
        Write-Host "✔ Added $upn to $groupName" -ForegroundColor Green
    } catch {
        if ($_.Exception.Message -like "*One or more added object references already exist*") {
            Write-Host "ℹ $upn is already a member of $groupName" -ForegroundColor Cyan
        } else {
            Write-Host "⚠ Failed to add ${upn}: $($_.Exception.Message)" -ForegroundColor Yellow
        }
    }
}
