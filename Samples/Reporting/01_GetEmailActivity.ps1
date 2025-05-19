<#
.SYNOPSIS
Gets details about email activity users have performed.

.EXAMPLE
.\01_GetEmailActivity.ps1 -period D30 -outPath "C:\temp"
#>


param(
    [Parameter(Mandatory=$true)]
    [ValidateSet("D7", "D30", "D90", "D180")]
    [string]$period,

    [Parameter(Mandatory=$true)]
    [string]$outPath
)

$fileName = "EmailActivity_{0}_{1}.csv" -f $period, (Get-Date -Format "yyyyMMdd")
$fullPath = Join-Path -Path $outPath -ChildPath $fileName

# Get config and helper
$root = Split-Path (Split-Path -Path $PSScriptRoot -Parent) -Parent
$config = Get-Content "$root\config\clientconfiguration.json" -Raw | ConvertFrom-Json

# Connect to Graph using Client Credential Flow with Certificate (Application Permission)
Connect-Graph -ClientId $config.ClientId -TenantId $config.TenantId -CertificateThumbprint $config.Thumbprint

# Query MS Graph - Get email activity
$activity = Get-MgReportEmailActivityUserDetail -period $period -OutFile $outPath

# Save output manually to file
$activity | Out-File -FilePath $fullPath -Encoding utf8

# Display activity
$activity

# Disconnect MS Graph
Disconnect-Graph