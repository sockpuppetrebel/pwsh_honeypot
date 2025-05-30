Import-Module Microsoft.Graph -Global

# Force load all submodules
Get-Module Microsoft.Graph -ListAvailable |
    ForEach-Object {
        Import-Module $_.Name -Global -Force
    }

