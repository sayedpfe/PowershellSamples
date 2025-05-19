# Grants an app registration with Sites.Selected read, write, or full control on
# Minimum Application Permission:
#   - To query sites and add permissions to a site: Sites.FullControl.All

# To best demo this:
# - Create an additional app registration (other than the one you set up in clientconfiguration.json)
# - Grant the additional app registration the Sites.Selected application permission
# - Update the values of the parameters below

Param(
  # URL of the site that will receive permissions
  [Parameter()]
  [Uri]$SiteToGrantPermissions = 'https://m365cpi90282478.sharepoint.com/sites/ProductionDepartment-shared70',

  # Client/application ID of the app registration with Sites.Selected
  [Parameter()]
  [Guid]$SitesSelectedAppId = '0a96d36f-b1cf-4227-9fcf-e247b629027e',

  # Name of the app registration with Sites.Selected
  [Parameter()]
  [string]$SitesSelectedAppName = 'ADO Wiki Graph Connector',

  # The permission that the app registration with Sites.Selected will receive on the SharePoint site in question
  [Parameter()]
  [ValidateSet('read', 'write', 'fullcontrol')]
  [string]$RoleToGrant = 'read'
)

# Get config
$root = Split-Path (Split-Path -Path $PSScriptRoot -Parent) -Parent
$config = Get-Content "$root\config\clientconfiguration.json" -Raw | ConvertFrom-Json

# Connect to Graph using Client Credential Flow with Certificate (Application Permission)
$null = Connect-Graph -ClientId $config.ClientId -TenantId $config.TenantId -CertificateThumbprint $config.Thumbprint

# Get the site resource so we have its Graph-friendly ID
$siteObject = Get-MgSite -SiteId ('{0}:{1}' -f $siteToGrantPermissions.DnsSafeHost, $siteToGrantPermissions.PathAndQuery) -Select Id

# If granting full control, a permission must be initially created with read or write, then updated after the fact to allow full control
$grantFullControl = ($RoleToGrant -eq 'fullcontrol')
if($grantFullControl) { $RoleToGrant = 'read' }

# Create the new permission
$sitePermission = @{
  SiteId              = $siteObject.Id
  Roles               = $RoleToGrant
  GrantedToIdentities = @(
    @{
      application = @{
        id          = $SitesSelectedAppId
        displayName = $SitesSelectedAppName
      }
    }
  )
}
$newPermission = New-MgSitePermission @sitePermission

# Update the new permission to use full control if it was specified in the parameters
if($grantFullControl) {
  $fullControlPermission = @{
    SiteId       = $siteObject.Id
    Roles        = 'fullcontrol'
    PermissionId = $newPermission.Id
  }
  Update-MgSitePermission @fullControlPermission
}