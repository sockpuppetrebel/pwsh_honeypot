# PowerShell Scripts Configuration
# Central configuration for certificate paths and common settings

# Get the root directory of the PowerShell-Scripts folder
$ScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path

# Certificate paths
$CertificatePaths = @{
    # Microsoft Graph certificates
    GraphCert = Join-Path $ScriptRoot "certificates\graph\jslater-graph.crt"
    GraphKey = Join-Path $ScriptRoot "certificates\graph\jslater-graph.key" 
    GraphPfx = Join-Path $ScriptRoot "certificates\graph\jslater-graph.pfx"
    GraphPfxNoPw = Join-Path $ScriptRoot "certificates\graph\jslater-graph-nopw.pfx"
    GraphPem = Join-Path $ScriptRoot "certificates\graph\jslater-graph.pem"
    GraphCsr = Join-Path $ScriptRoot "certificates\graph\jslater-graph.csr"
    
    # Azure App certificates  
    AzureCert = Join-Path $ScriptRoot "certificates\azure\azure_app_cert.pem"
    AzureKey = Join-Path $ScriptRoot "certificates\azure\azure_app_key.pem"
    GenericCert = Join-Path $ScriptRoot "certificates\azure\cert.pem"
    GenericKey = Join-Path $ScriptRoot "certificates\azure\key.pem"
}

# Output directories
$OutputPaths = @{
    SharePointAudit = Join-Path $ScriptRoot "output\sharepoint\audit_reports"
    SharePointPermissions = Join-Path $ScriptRoot "output\sharepoint\permission_reports"
}

# Function to validate certificate exists
function Test-CertificateExists {
    param([string]$CertPath)
    
    if (Test-Path $CertPath) {
        Write-Host "✓ Certificate found: $CertPath" -ForegroundColor Green
        return $true
    } else {
        Write-Warning "✗ Certificate not found: $CertPath"
        return $false
    }
}

# Function to get certificate path by name
function Get-CertificatePath {
    param([string]$CertName)
    
    if ($CertificatePaths.ContainsKey($CertName)) {
        return $CertificatePaths[$CertName]
    } else {
        throw "Certificate '$CertName' not found in configuration"
    }
}

# Export the paths for use in other scripts
Export-ModuleMember -Variable CertificatePaths, OutputPaths -Function Test-CertificateExists, Get-CertificatePath