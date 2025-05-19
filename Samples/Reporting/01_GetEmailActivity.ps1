# Gets details about email activity users have performed.
# Minimum Application Permission: Reports.Read.All
# https://docs.microsoft.com/en-us/graph/api/reportroot-getemailactivityuserdetail

<#
ISSUES:
   - Currently, SDK does not accept non-json responses. This is an open issue:
   - https://github.com/microsoftgraph/msgraph-sdk-powershell/issues/182
#>

# Variables
$period  = "D30" # Supported values: D7,D30,D90,D180
$outPath = "C:\temp"

# Get config and helper
$root = Split-Path (Split-Path -Path $PSScriptRoot -Parent) -Parent
$config = Get-Content "$root\config\clientconfiguration.json" -Raw | ConvertFrom-Json

# Connect to Graph using Client Credential Flow with Certificate (Application Permission)
Connect-Graph -ClientId $config.ClientId -TenantId $config.TenantId -CertificateThumbprint $config.Thumbprint

# Query MS Graph - Get email activity
$activity = Get-MgReportEmailActivityUserDetail -Period $period -OutFile $outPath
$activity

# Disconnect MS Graph
Disconnect-Graph