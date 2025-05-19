<#
.SYNOPSIS
Sends an email using Microsoft Graph.

.EXAMPLE
.\05_SendEmail.ps1 -mailFromUpn "admin@M365CPI90282478.onmicrosoft.com" -mailToUpn "AdilE@M365CPI90282478.OnMicrosoft.com" -mailsubject "Hello MS Graph!" -mailContent "This is a sample mail sent via MS Graph. How cool is this?"
#>

param (
    [Parameter(Mandatory = $true)]
    [string]$mailFromUpn,

    [Parameter(Mandatory = $true)]
    [string]$mailToUpn,

    [Parameter(Mandatory = $true)]
    [string]$mailsubject,

    [Parameter(Mandatory = $true)]
    [string]$mailContent
)

# Sends an email
# Minimum Application Permission: Mail.Send
# https://docs.microsoft.com/en-us/graph/api/user-sendmail

# Get config and helper
$root = Split-Path (Split-Path -Path $PSScriptRoot -Parent) -Parent
$config = Get-Content "$root\config\clientconfiguration.json" -Raw | ConvertFrom-Json

# Connect to Graph using Client Credential Flow with Certificate (Application Permission)
Connect-Graph -ClientId $config.ClientId -TenantId $config.TenantId -CertificateThumbprint $config.Thumbprint

# Create message
$message = @{
  subject = $mailsubject
  body = @{
    contentType = "Text"
    content = $mailContent
  }
  toRecipients = @(
    @{
      emailAddress = @{
        address = $mailToUpn
      }
    }
  )
}

# Query MS Graph - send email
$mailFrom = Get-MGUser -UserId $mailFromUpn
Send-MgUserMail -UserId $mailFrom.id -BodyParameter @{message = $message}

# Disconnect MS Graph
Disconnect-Graph