# Gets the sign in information of guest users
# Minimum Application Permission: 
#   - To list users: User.Read.All
#   - To get sign in info: AuditLog.Read.All and Directory.Read.All
# https://docs.microsoft.com/en-us/graph/api/user-list  
# https://docs.microsoft.com/en-us/graph/api/signin-list


# Get config and helper
$root = Split-Path (Split-Path -Path $PSScriptRoot -Parent) -Parent
$config = Get-Content "$root\config\clientconfiguration.json" -Raw | ConvertFrom-Json
Connect-Graph -ClientId $config.ClientId -TenantId $config.TenantId -CertificateThumbprint $config.Thumbprint

$guestUsers = Get-MgUser -Filter "userType eq 'Guest'" -Property id,mail,userPrincipalName

$results = [System.Collections.Generic.List[PSCustomObject]]::new()

foreach ($guestUser in $guestUsers) {

  # /auditLogs/signIns default sort is newest to oldest, so we only need the first result
  $lastSignIn = Get-MgAuditLogSignIn -Filter "userId eq '$($guestUser.Id)'" -Top 1 | Select-Object -ExpandProperty CreatedDateTime
  if($null -eq $lastSignIn) {
    $lastSignIn = 'Earlier than the audit log period'
  }

  # Create the output object
  $result = [PSCustomObject]@{
    UPN        = $guestUser.userPrincipalName
    Email      = $guestUser.mail
    LastSignIn = $lastSignIn
  }

  # Populate the last sign in based on count of signins
  $results.Add($result)

}

# Output
$results