# Sends an email
# Minimum Application Permission: Mail.Send
# https://docs.microsoft.com/en-us/graph/api/user-sendmail

# Variables
$mailFromUpn = "admin@M365CPI90282478.onmicrosoft.com"
$mailToUpn   = "AdilE@M365CPI90282478.OnMicrosoft.com"
$mailsubject = "Hello MS Graph!"
$mailContent = "This is a sample mail sent via MS Graph. How cool is this?"

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