#requires -version 5.1
#requires -Modules Microsoft.Graph.Authentication, Microsoft.Graph.Applications

# Get all service principals in the tenant using Windows Azure Active Directory
# Minimum application permission: Application.Read.All, DelegatedPermissionGrant.ReadWrite.All
# Alternate application permission: Directory.Read.All

Param(

  # Optionally resolve owner names for apps registered in your tenant. This makes additional API
  # calls and will increase the time it takes for this script to run
  [Parameter()]
  [switch]$ResolveAppRegistrationOwners,

  # Don't change this
  [Parameter()]
  [guid]$WindowsAzureAdApiId = '00000003-0000-0000-c000-000000000000'

)

# Get config and helper
$root = Split-Path (Split-Path -Path $PSScriptRoot -Parent) -Parent
$config = Get-Content "$root\config\clientconfiguration.json" -Raw | ConvertFrom-Json
Connect-Graph -ClientId $config.ClientId -TenantId $config.TenantId -CertificateThumbprint $config.Thumbprint

# We'll need the tenant ID to determine if apps are located in your tenant, or if they're 3rd party
$tenantId = (Get-MgContext).TenantId

# For other ways to connect to Microsoft Graph, please visit:
# https://learn.microsoft.com/en-us/powershell/microsoftgraph/authentication-commands

# Get the object ID for Windows Azure Active Directory (AAD) on the current tenant
$windowsAzureAdGraphPrincipal = Get-MgServicePrincipal -Filter "appId eq '$WindowsAzureAdApiId'" -Property Id

# Get all the apps where a user or administrator has granted a Windows AAD permission
$forWindowsAadGraph = @{
  All      = $true
  Filter   = "resourceId eq '$($windowsAzureAdGraphPrincipal.Id)'"
  Property = 'clientId'
}
$windowsAadGraphConsents = Get-MgOauth2PermissionGrant @forWindowsAadGraph

# Extract only the unique applications using Windows AAD permissions
$uniqueWindowsAadAppIds = $windowsAadGraphConsents.ClientId | Select-Object

# Optimized way to ask Graph for multiple service principals with a single query
for($index = 0; $index -lt $uniqueWindowsAadAppIds.Count; $index += 15) {

  # 15 at a time because it's the maximum for the /servicePrincipals endpoint
  $fifteenPrincipalsAtATime = @{
    Filter   = "id in ('{0}')" -f ($uniqueWindowsAadAppIds[$index..($index+14)] -join "','")
    Property = @('id', 'appId', 'appOwnerOrganizationId', 'displayName')
  }
  $appPrincipals = Get-MgServicePrincipal @fifteenPrincipalsAtATime

  # Output what we found (ManageUrl probably won't be clickable from a terminal window)
  foreach($appPrincipal in $appPrincipals) {

    # Determine what will appear in the "OwnedBy" column. First is this app native to this tenant?
    $appOwnedBy = if($appPrincipal.AppOwnerOrganizationId -eq $tenantId) {

      # Next, should we get the app registration owners?
      if($ResolveAppRegistrationOwners) {

        # Optimize the query to include owners' display name and UPN only
        $withAppOwners = @{
          Filter         = "appId eq '$($appPrincipal.appId)'"
          Property       = @('id')
          ExpandProperty = 'owners($select=displayName,userPrincipalName)'
        }
        $appRegistration = Get-MgApplication @withAppOwners

        foreach($owner in $appRegistration.Owners) {
          '{0} ({1})' -f $owner.AdditionalProperties['displayName'], $owner.AdditionalProperties['userPrincipalName']
        }
      }

      # If we don't want individual owners, simply report that it's an app in this tenant
      else { 'This tenant' }
    }
    # If this isn't an app in this tenant, report it as third-party
    else { 'Third-party' }

    # Output (recommend exporting to CSV)
    [PSCustomObject]@{
      DisplayName        = $appPrincipal.DisplayName
      ServicePrincipalId = $appPrincipal.Id
      AppOwnedBy         = ($appOwnedBy -join '; ')
      ManageUrl          = "https://entra.microsoft.com/#view/Microsoft_AAD_IAM/ManagedAppMenuBlade/~/Overview/objectId/$($appPrincipal.Id)/appId/$($appPrincipal.AppId)/preferredSingleSignOnMode~/null/servicePrincipalType/Application"
    }
  }

}