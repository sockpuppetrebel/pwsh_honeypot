# IT-Bot Generated Script
# Created: 2025-06-14T22:39:39.008Z
# User: test-admin
# Auto-saved from chat interaction

$oldDate = (Get-Date).AddDays(-90)
Get-AzureADUser | Where-Object { $_.LastSignInDateTime -lt $oldDate }