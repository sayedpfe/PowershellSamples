# Creates Team
# Minimum Application Permission: Directory.ReadWrite.All, Group.ReadWrite.All, Team.Create
# https://learn.microsoft.com/en-us/powershell/module/microsoft.graph.teams/?view=graph-powershell-1.0

# Variables
$teamDisplayName = "KB Second Product Team"
$teamDescription = "KBProduct Team for internal collaboration"
$teamOwnerUpn    = "admin@M365CPI90282478.onmicrosoft.com"
$channelDisplayName = "Knorr-Bremse"
$channelDescription = "Knorr-Bremse"

# Get config and helper
$root = Split-Path (Split-Path -Path $PSScriptRoot -Parent) -Parent
$config = Get-Content "$root\config\clientconfiguration.json" -Raw | ConvertFrom-Json

# Connect to Graph using Client Credential Flow with Certificate (Application Permission)
Connect-Graph -ClientId $config.ClientId -TenantId $config.TenantId -CertificateThumbprint $config.Thumbprint

$params = @{
  "Template@odata.bind" = "https://graph.microsoft.com/v1.0/teamsTemplates('standard')"
  DisplayName           = $teamDisplayName
  Description           = $teamDescription
  Members               = @(
    @{
      "@odata.type"     = "#microsoft.graph.aadUserConversationMember"
      Roles             = @(
        "owner"
      )
      "User@odata.bind" = "https://graph.microsoft.com/v1.0/users('$teamOwnerUpn')"
    }
  )
}

$newCreatedTeam=New-MgTeam -BodyParameter $params

Get-MgUserJoinedTeam -UserId $teamOwnerUpn

$body = @{
    displayName = $channelDisplayName
    description = $channelDescription
}
$teamChannel = New-MgTeamChannel -TeamId $newCreatedTeam.Id -BodyParameter $body
$teamChannel


# Disconnect MS Graph
Disconnect-Graph