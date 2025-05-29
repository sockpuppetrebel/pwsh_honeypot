# Check if the Exchange Online PowerShell module is installed
if (!(Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
    Write-Warning "[WARNING] ExchangeOnlineManagement module is not installed."
    Write-Host "[INFO] Please install it with: Install-Module -Name ExchangeOnlineManagement -Force"
    $moduleInstalled = $false
} else {
    $moduleInstalled = $true
    
    # Import the module
    Import-Module ExchangeOnlineManagement

    # Try to connect to Exchange Online
    try {
        # Check if already connected
        Get-Command Get-Mailbox -ErrorAction Stop | Out-Null
        Write-Host "[INFO] Already connected to Exchange Online"
    } catch {
        Write-Host "[INFO] Please authenticate with your Exchange Online credentials..."
        Write-Host "[INFO] A popup window may appear for authentication"
    }
}

$emails = @(
"christopher.koutis@optimizely.com",
"Dyel.KassaKoumba@optimizely.com",
"line.wilkens-lintrup@optimizely.com",
"salome.isanovic@optimizely.com",
"Emily.Harford@optimizely.com",
"Ella.Clark@optimizely.com",
"Sofia.Larsson@optimizely.com",
"Audrey.Hungerman@optimizely.com",
"Derrick.Arakaki@optimizely.com",
"Matthew.Harris@optimizely.com",
"Simone.AraujoCoelho@optimizely.com",
"Drew.Elston@optimizely.com",
"Craig.Stryker@optimizely.com",
"Mathias.Lehniger@optimizely.com",
"Yanni.Panacopoulos@optimizely.com",
"Michelle.Williams@optimizely.com",
"Sterling.Nostedt@optimizely.com",
"Katerina.Gavalas@optimizely.com",
"jacob.khan@optimizely.com",
"Jeremy.Davis@optimizely.com",
"jon.greene@optimizely.com",
"Matthew.Kingham@optimizely.com",
"Joe.Duffell@optimizely.com",
"Nils.Qvarfordt@optimizely.com",
"Simon.McDonald@optimizely.com",
"Jenna.Esterson@optimizely.com",
"Ravi.Khera@optimizely.com",
"Cliff.Hill@optimizely.com",
"Terrence.McGregor@optimizely.com",
"Katarina.Lister@optimizely.com",
"Lisa.Martwichuck@optimizely.com"
)

$GroupName = "Revenue Leaders Forecast"

foreach ($email in $emails) {
    try {
        Add-DistributionGroupMember -Identity $GroupName -Member $email -ErrorAction Stop
        Write-Host "[SUCCESS] Added $email to $GroupName"
    } catch {
        Write-Warning "[ERROR] Failed to add $email - $($_.Exception.Message)"
    }
}

