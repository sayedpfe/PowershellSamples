# Gets a list of deleted O365 groups and check if their corresponding SPO site is also deleted or orphan (not deleted)
# Minimum Application Permission: 
#   - To get deleted groups: Group.Read.All
#   - To get SPO Site: Sites.Read.All
# https://docs.microsoft.com/en-us/graph/api/directory-deleteditems-list
# https://docs.microsoft.com/en-us/graph/api/site-get

# To best demo this:
# - Create an O365 group. Wait for it to provision the SPO site.
# - Delete the O365 group through the UI. The SPO site will not get immediately deleted and therefore will be flagged as orphan.

# Get config
$root = Split-Path (Split-Path -Path $PSScriptRoot -Parent) -Parent
$config = Get-Content "$root\config\clientconfiguration.json" -Raw | ConvertFrom-Json

# Connect to Graph using Client Credential Flow with Certificate (Application Permission)
$CertThumbprint = $config.Thumbprint
$cert = Get-Item "Cert:\CurrentUser\My\$CertThumbprint"
$token = Get-MsalToken -ClientId $config.ClientId -TenantId $config.TenantId -ClientCertificate $cert -Scopes "https://graph.microsoft.com/.default"
$headers = @{ Authorization = "Bearer $($token.AccessToken)" }

# Query MS Graph - Get deleted groups
$uri = "https://graph.microsoft.com/v1.0/directory/deletedItems/microsoft.graph.group"
$response = Invoke-RestMethod -Uri $uri -Headers @{Authorization = "Bearer $($token.AccessToken)"}
$deletedGroups = $response.value
#$deletedGroups

$orphanSites = @()
foreach ($deletedGroup in $deletedGroups){
    # Query MS Graph - Get SPO Site based on O365 Group
    $uri = "https://graph.microsoft.com/v1.0/sites?search="+$deletedGroup.DisplayName
    $response = Invoke-RestMethod -Uri $uri -Headers @{Authorization = "Bearer $($token.AccessToken)"}
    $spoSite = $response.value
    if ([System.String]::IsNullOrEmpty($spoSite)){
        Write-Host "No orphan sites found that is associated to deleted group '$($deletedGroup.displayName)'" -ForegroundColor Green
    }
    else{
        # Write-Host "Orphan site found - Group '$($deletedGroup.displayName)' was deleted but its associated SPO site '$($sposite[0].webUrl)' was not." -ForegroundColor Yellow
        $orphanSites += $spoSite
    }
}
Write-Host "`nThe below sites are orphan (associated O365 Group was deleted):" -ForegroundColor Cyan
$orphanSites | Select webUrl
