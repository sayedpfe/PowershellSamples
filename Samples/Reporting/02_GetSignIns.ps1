# Retrieves the Azure AD user sign-ins for your tenant
# Minimum Application Permission: AuditLog.Read.All and Directory.Read.All
# https://docs.microsoft.com/en-us/graph/api/signin-list
# Requires at least Azure P1

# Get config and helper
$root = Split-Path (Split-Path -Path $PSScriptRoot -Parent) -Parent
$config = Get-Content "$root\config\clientconfiguration.json" -Raw | ConvertFrom-Json

# Connect to Graph using Client Credential Flow with Certificate (Application Permission)
Connect-Graph -ClientId $config.ClientId -TenantId $config.TenantId -CertificateThumbprint $config.Thumbprint

# Query MS Graph - Sign Ins
$signIns = Get-MgAuditLogSignIn | Select-Object -First 5
$signIns | fl

# Disconnect MS Graph
Disconnect-Graph