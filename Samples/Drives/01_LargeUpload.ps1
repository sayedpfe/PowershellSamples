# Large file upload to a drive
# Minimum Application Permission: Files.ReadWrite.All
# Adapted from https://learn.microsoft.com/en-us/answers/questions/587655/graph-upload-large-file-with-powershell.html

Param(
  # Path to the local file
  [Parameter()]
  [ValidateScript({ if(-not (Test-Path -Path (Resolve-Path -Path $_))) { throw "No file at $_" } else { $true } })]
  [string]$LocalFilePath,

  # Drive ID
  [Parameter()]
  [string]$DriveId,

  [Parameter()]
  [ValidateScript({ if($_ % 320 -eq 0) { $true } else { throw 'Chunk size must be evenly divisible by 320' } })]
  [int]$ChunkSize = 320
)

# Get config and helper
$root = Split-Path (Split-Path -Path $PSScriptRoot -Parent) -Parent
$config = Get-Content "$root\config\clientconfiguration.json" -Raw | ConvertFrom-Json
Connect-Graph -ClientId $config.ClientId -TenantId $config.TenantId -CertificateThumbprint $config.Thumbprint

# Default chunk size of 320 = 327680
$partSize   = $ChunkSize * 1024

# Get some file information
$fileName   = Split-Path -Path $LocalFilePath -Leaf
$fileBytes  = [System.IO.File]::ReadAllBytes($LocalFilePath)
$fileLength = $fileBytes.Length

$newSession = @{
  DriveId       = $DriveId
  DriveItemId   = "root:/${fileName}:"
  BodyParameter = @{
    fileSize = $fileLength
    name     = $fileName
  }
}
$uploadSession = New-MgDriveItemUploadSession @newSession -ErrorAction Stop

for($index = $start = $end = 0; $fileLength -gt ($end+1) ; $index++) {

  $start = $index * $partSize
  if (($start + $partSize - 1) -lt $fileLength) {
    $end = ($start + $partSize - 1)
  }
  else {
    $end = ($start + ($fileLength - ($index * $partSize)) - 1)
  }

  [byte[]]$body = $fileBytes[$start..$end]

  $partUpload = @{
    Method               = 'Put'
    Uri                  = $uploadSession.UploadUrl
    Body                 = $body
    SkipHeaderValidation = $true
    Headers              = @{
      'Content-Range' = "bytes $start-$end/$fileLength"
    }
  }
  try  { $response = Invoke-RestMethod @partUpload -ErrorAction Stop }
  catch {
    "Error while uploading bytes $start-$end/$fileLength - $($_.Exception.Message)" | Write-Warning
  }

}

write-verbose 'finished'