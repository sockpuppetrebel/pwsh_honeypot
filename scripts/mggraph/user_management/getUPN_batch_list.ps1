# Connect to Microsoft Graph
Connect-MgGraph -Scopes "User.Read.All" -UseDeviceAuthentication -NoWelcome

$names = @(
"Jack McClean"
"Mike Cartwright"
"Paul Gray"
"William LaPuma"
"Mark Ryan"
"Matthew Payne"
"David Bingham"
"Anna Parback"
"Jack Joseph"
"Thomas McKenzie"
"Sean Groat"
"Brandon Halvorson"
"Daniel Martell"
"Jon Jones"
"Rob Stoves"
"Nuno Figueiredo"
"Marcus Hoffman"
"Rob Saunders"
"Zachary Coulter"
"Phil Yates"
"Shannon Gray"
"Anatoliy Savinov"
"Brett Samuels"
"Anna Redmile"
"Aidan Dodd"
"Vimi Kaul"
"Mark Wakelin"
"Chynna Roberts"
"Tarik Antunes"
"Alexandra Van Heel"
"Jennifer Lovett"
"Terry McGregor"
"Robin LeClerc"
)

foreach ($name in $names) {
  $user = Get-MgUser -Filter "DisplayName eq '$name'" -ConsistencyLevel eventual
  if ($user) {
    $user | Select-Object DisplayName, UserPrincipalName
  } else {
    Write-Warning "Not found: $name"
  }
}
