# Creates a user and assigns licenses with specific service plans enabled.
# Minimum Application Permission (to create users, assign license, and get SKU information) : User.ReadWrite.All, Directory.Read.All
# https://docs.microsoft.com/en-us/graph/api/user-post-users
# https://docs.microsoft.com/en-us/graph/api/user-assignlicense

Param(
  # User's display name
  [Parameter()]
  [string]$DisplayName = 'Graph PowerShell',

  # User's email alias
  [Parameter()]
  [string]$MailNickname = 'graphposh',

  # Desired user principal name
  [Parameter()]
  [string]$UserPrincipalName = 'graphposh@M365CPI90282478.onmicrosoft.com',

  # User's first name
  [Parameter()]
  [string]$GivenName = 'Graph',

  # User's last name
  [Parameter()]
  [string]$Surname = 'PowerShell',

  # User's job title
  [Parameter()]
  [string]$JobTitle = 'PowerShell Pro',

  # User's preferred data location
  [Parameter()]
  [string]$usageLocation = 'US',

  # User's password
  [Parameter(Mandatory)]
  [securestring]$Password,

  # The license that the user will receive. To see which are available, run:
  # PS> Get-MgSubscribedSku | Select-Object -Property SkuPartNumber
  [Parameter()]
  [string]$LicenseSku = 'Microsoft_365_Copilot',

  # Licenses from the SKU to disable. To see which are available, run:
  # PS> (Get-MgSubScribedSku | Where-Object -Property SkuPartNumber -EQ -Value '<SKU name>').ServicePlans
  # Or visit https://docs.microsoft.com/en-us/azure/active-directory/enterprise-users/licensing-service-plan-reference
  [Parameter()]
  [string[]]$DisabledLicenses = @('SWAY', 'BI_AZURE_P2')
)

# Get config and helper
$root = Split-Path (Split-Path -Path $PSScriptRoot -Parent) -Parent
$config = Get-Content "$root\config\clientconfiguration.json" -Raw | ConvertFrom-Json
Connect-Graph -ClientId $config.ClientId -TenantId $config.TenantId -CertificateThumbprint $config.Thumbprint

# Create new user body
$newUserRequest = @{
  AccountEnabled    = $true
  DisplayName       = $DisplayName
  MailNickName      = $MailNickName
  UserPrincipalName = $userPrincipalName
  GivenName         = $givenName
  Surname           = $surname
  JobTitle          = $jobTitle
  UsageLocation     = $usageLocation
  PasswordProfile   =  @{
    password = [System.Runtime.InteropServices.Marshal]::PtrToStringUni([System.Runtime.InteropServices.Marshal]::SecureStringToCoTaskMemUnicode($Password))
    forceChangePasswordNextSignIn = $false
  }
  PasswordPolicies  = "DisablePasswordExpiration, DisableStrongPassword"
}

$newUser = New-MgUser @newUserRequest -ErrorAction Stop

# Validate SKU
$skus = Get-MgSubscribedSku
$sku  = $skus | Where-Object -Property SkuPartNumber -IEQ -Value $LicenseSku | Select-Object -First 1
if(-not $sku) {
  throw 'No subscribed SKU named {0}; available SKUs are {1}' -f $LicenseSku, ($skus.SkuPartNumber -join ', ')
}

# Validate that licenses to disable exist for the SKU
for($d = $DisabledLicenses.Count-1; $d -ge 0; $d--) {
  $disabledLicense = $DisabledLicenses[$d]

  # If license is a name and not a GUID, try to resolve it to a GUID
  if($disabledLicense -notmatch '(?im)^[{(]?[0-9A-F]{8}[-]?(?:[0-9A-F]{4}[-]?){3}[0-9A-F]{12}[)}]?$') {
    $licenseName     = $disabledLicense
    $disabledLicense = ($sku.ServicePlans | Where-Object -Property ServicePlanName -EQ -Value $licenseName).ServicePlanId

    # Remove the disabled license if it's not valid for the SKU
    if([string]::IsNullOrEmpty($disabledLicense)) {
      Write-Warning -Message "No license available for SKU $LicenseSku named $licenseName"
      $DisabledLicenses.RemoveAt($d)
    }
    # If it's a valid license, replace the name with a GUID
    else {
      $DisabledLicenses[$d] = $disabledLicense
    }
  }
}

# Build license body
$addLicenses = @(
  @{
    disabledPlans = $DisabledLicenses
    skuId         = $sku.SkuId
  }
)

Set-MgUserLicense -UserId $newUser.Id -AddLicenses $addLicenses -RemoveLicenses @()