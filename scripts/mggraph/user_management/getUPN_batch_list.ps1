# Connect to Microsoft Graph
Connect-MgGraph -Scopes "User.Read.All" -UseDeviceAuthentication -NoWelcome

$names = @(
"First Last"
"John Smith"
"Jane Doe"
"Sample User"
"Test User"
)

foreach ($name in $names) {
  $user = Get-MgUser -Filter "DisplayName eq '$name'" -ConsistencyLevel eventual
  if ($user) {
    $user | Select-Object DisplayName, UserPrincipalName
  } else {
    Write-Warning "Not found: $name"
  }
}
