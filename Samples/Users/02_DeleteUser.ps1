# Deletes a user
# Minimum Application Permission: User.ReadWrite.All
# https://docs.microsoft.com/en-us/graph/api/user-delete

Param(
  [Parameter(Mandatory)]
  [string]$UserPrincipalName
)

# Get config and helper
$root = Split-Path (Split-Path -Path $PSScriptRoot -Parent) -Parent
$config = Get-Content "$root\config\clientconfiguration.json" -Raw | ConvertFrom-Json
Connect-Graph -ClientId $config.ClientId -TenantId $config.TenantId -CertificateThumbprint $config.Thumbprint

# Query MS Graph - Get guest users
Remove-MgUser -UserId $UserPrincipalName