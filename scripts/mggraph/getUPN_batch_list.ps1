$names = @(
"Chris Koutis"
"Dyel Kassa Koumba"
"Line Wilkens-Lintrup"
"Salome Isanovic"
"Emily Harford"
"Ella Clark"
"Sofia Larsson"
"Audrey Hungerman"
"Derrick Arakaki"
"Matthew Harris"
"Stephen Gemous"
"Craig Ferrera"
"Simone Araujo Coelho"
"Drew Elston"
"Craig Stryker"
"Mathias Lehniger"
"Yanni Panacopoulos"
"Michelle LeBlanc Williams"
"Sterling Nostedt"
"Opti-CsOps@optimizely.com"
"Katerina Gavalas"
"Jacob Khan"
"Jeremy Davis"
"Jon Greene"
"Matthew Kingham"
"Joe Duffell"
"Nils Qvarfordt"
"Simon McDonald"
"Jeron Schuijt"
"Jenna Esterson"
"Ravi Khera"
"Cliff Hill"
"Terrence McGregor"
"Katarina Lister"
"Lisa Martwichuck"

)

foreach ($name in $names) {
  $user = Get-MgUser -Filter "DisplayName eq '$name'" -ConsistencyLevel eventual
  if ($user) {
    $user | Select-Object DisplayName, UserPrincipalName
  } else {
    Write-Warning "Not found: $name"
  }
}
