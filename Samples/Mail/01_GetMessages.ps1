<#
.SYNOPSIS
    Get the Top 5 messages for a given user using Microsoft Graph.

.EXAMPLE
    # Run this script from PowerShell:
    .\01_GetMessages.ps1 -user "user@domain.com"
#>


# Gets All email messages for a given user (including deleted ones)
# Minimum Application Permission: Mail.Read
# https://docs.microsoft.com/en-us/graph/api/user-list-messages
# 

if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Mail)) {
    Write-Host "Microsoft.Graph module not found. Installing..."
    Install-Module -Name Microsoft.Graph.Mail -Scope CurrentUser -Force
}

param(
    [Parameter(Mandatory = $true)]
    [string]$user
)

# Gets All email messages for a given user (including deleted ones)
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
# $user = 'AmberR@M365CPI90282478.OnMicrosoft.com'   # <-- Remove or comment out this line

# Get config
$root = Split-Path (Split-Path -Path $PSScriptRoot -Parent) -Parent
$config = Get-Content "$root\config\clientconfiguration.json" -Raw | ConvertFrom-Json

# Connect to Graph using Client Credential Flow with Certificate (Application Permission)
Connect-Graph -ClientId $config.ClientId -TenantId $config.TenantId -CertificateThumbprint $config.Thumbprint

# Query MS Graph
$messages = Get-MgUserMessage -UserId $user -Top 20

# Output - Format messages
Format-EmailMessages $messages

# Disconnect MS Graph
Disconnect-Graph