<#
.SYNOPSIS
Gets the parent folder of email messages that meet search criteria for a given user.

.EXAMPLE
.\03_MessageFolderName.ps1 -User user@contoso.com -SearchString "Weekly digest: Office 365"
#>

param (
    [Parameter(Mandatory = $true)]
    [string]$User,

    [Parameter(Mandatory = $true)]
    [string]$SearchString
)

# Minimum Application Permission: Mail.Read
# https://docs.microsoft.com/en-us/graph/api/user-list-messages
# https://docs.microsoft.com/en-us/graph/api/mailfolder-get

# Variables
# $user = 'user@contoso.com'
# $searchString = "Weekly digest: Office 365"
# Get config and helper
$root = Split-Path (Split-Path -Path $PSScriptRoot -Parent) -Parent
$config = Get-Content "$root\config\clientconfiguration.json" -Raw | ConvertFrom-Json

# Connect to Graph using Client Credential Flow with Certificate (Application Permission)
Connect-Graph -ClientId $config.ClientId -TenantId $config.TenantId -CertificateThumbprint $config.Thumbprint

# Query MS Graph - get messages
$results = @()
$messages = Get-MgUserMessage -UserId $user -Search `"$searchString`"

# Query MS Graph - get parent folder
$messages | ForEach-Object {
    $m = Get-MgUserMailFolder -UserId $user -MailFolderId $_.ParentFolderId
    $results += $m 
}

# Output
$results | fl

# Disconnect MS Graph
Disconnect-Graph