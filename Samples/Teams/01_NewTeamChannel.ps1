# Creates a channel on an existing MS Team
# Minimum Application Permission: Group.ReadWrite.All
# https://docs.microsoft.com/en-us/graph/api/channel-post

# Variables
$teamDisplayName    = "New SP Team"
$channelDisplayName = "Knorr-Bremse"
$channelDescription = "Knorr-Bremse"

# Get config and helper
$root = Split-Path (Split-Path -Path $PSScriptRoot -Parent) -Parent
$config = Get-Content "$root\config\clientconfiguration.json" -Raw | ConvertFrom-Json

# Connect to Graph using Client Credential Flow with Certificate (Application Permission)
Connect-Graph -ClientId $config.ClientId -TenantId $config.TenantId -CertificateThumbprint $config.Thumbprint

# Query MS Graph - Get Team
$group = Get-MgGroup -Filter "displayName eq '$teamDisplayName'"
#$team = Get-MgTeam -TeamId $group.Id

# Query MS Graph - Create channel
$body = @{
    displayName = $channelDisplayName
    description = $channelDescription
}
$teamChannel = New-MgTeamChannel -TeamId $group.Id -BodyParameter $body
$teamChannel

# Disconnect MS Graph
Disconnect-Graph