# Gets groups that are either orphan (no owners) or only have one owner.
# Minimum Application Permission: Group.Read.All
# https://docs.microsoft.com/en-us/graph/api/group-get

# Get config and helper
$root = Split-Path (Split-Path -Path $PSScriptRoot -Parent) -Parent
$config = Get-Content "$root\config\clientconfiguration.json" -Raw | ConvertFrom-Json
Connect-Graph -ClientId $config.ClientId -TenantId $config.TenantId -CertificateThumbprint $config.Thumbprint

$groups = Get-MgGroup -Filter "groupTypes/any(c:c eq 'Unified')" -All

$results = [System.Collections.Generic.List[PSCustomObject]]::new()

foreach($group in $groups) {

  $owners  = Get-MgGroupOwner  -GroupId $group.Id
  $members = Get-MgGroupMember -GroupId $group.Id

  # Microsoft.Graph.Groups module v1.0.1 doesn't report count correctly for $owners or $members
  # Using $owners.Id.Count and $members.Id.Count will get the correct count
  # Keep an eye on https://github.com/microsoftgraph/msgraph-sdk-powershell/issues/418

  if($owners.Id.Count -in 0..1) {
    $results.Add([PSCustomObject]@{
      DisplayName = $group.DisplayName
      OwnersCount = $owners.Id.Count
      MembersCount = $members.Id.Count
      Id = $group.Id
    })
  }

}

# Output
$results