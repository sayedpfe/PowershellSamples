<#
.SYNOPSIS
    Search emails by subject for a given user using Microsoft Graph.

.EXAMPLE
    # Run this script from PowerShell:
    .\02_SearchEmailBySubject.ps1 -User "user@domain.com" -SearchString "Operations Department"
#>

param (
    [Parameter(Mandatory = $true)]
    [string]$User,

    [Parameter(Mandatory = $true)]
    [string]$SearchString
)

# Gets email for a given user that meet search criteria
# Minimum Application Permission: Mail.Read
# https://docs.microsoft.com/en-us/graph/api/user-list-messages



function Format-EmailMessages {

  [CmdletBinding()]
  param
  (
    [Parameter(Mandatory = $true, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
    [object]$messages
  )
    
  $messages | Select-Object `
    Id, `
    Subject, `
    ConversationId, `
    ParentFolderId, `
    CreatedDateTime, `
  @{Label="From"; Expression={ $_.From.EmailAddress.Address } }, `
    IsRead `
  | Sort-Object CreatedDateTime -Descending
}

# Variables
$user = 'AmberR@M365CPI90282478.onmicrosoft.com'
$searchString = "Operations Department"

# Get config and helper
$root = Split-Path (Split-Path -Path $PSScriptRoot -Parent) -Parent
$config = Get-Content "$root\config\clientconfiguration.json" -Raw | ConvertFrom-Json

# Connect to Graph using Client Credential Flow with Certificate (Application Permission)
Connect-Graph -ClientId $config.ClientId -TenantId $config.TenantId -CertificateThumbprint $config.Thumbprint

# Query MS Graph - search for messages
$messages = Get-MgUserMessage -UserId $user -Search `"$searchString`"

# Output - Format messages
Format-EmailMessages $messages

# Disconnect MS Graph
Disconnect-Graph