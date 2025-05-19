<#
.SYNOPSIS
    Get a message header for a given message.

.EXAMPLE
    .\04_GetMessageHeader.ps1 -User user@contoso.com -SearchString "Weekly digest: Office 365"
#>

param (
    [Parameter(Mandatory = $true)]
    [string]$User,

    [Parameter(Mandatory = $true)]
    [string]$SearchString
)

# Minimum Application Permission: Mail.Read
# https://docs.microsoft.com/en-us/graph/api/user-list-messages

# Variables
# $user = 'user@contoso.com'
# $searchString = "Weekly digest: Office 365"
# Get config and helper
$root = Split-Path (Split-Path -Path $PSScriptRoot -Parent) -Parent
$config = Get-Content "$root\config\clientconfiguration.json" -Raw | ConvertFrom-Json

# Connect to Graph using Client Credential Flow with Certificate (Application Permission)
Connect-Graph -ClientId $config.ClientId -TenantId $config.TenantId -CertificateThumbprint $config.Thumbprint

# Query MS Graph - get messages
$messages = Get-MgUserMessage -UserId $user -Search `"$searchString`"

# Query MS Graph - get headers from first message
$messageId = ($messages | Select-Object -First 1).id
$m = Get-MgUserMessage -UserId $user -MessageId $messageId -Property internetMessageHeaders

# Output
$rawHeader = $m.internetMessageHeaders
$rawHeader

# Disconnect MS Graph
Disconnect-Graph