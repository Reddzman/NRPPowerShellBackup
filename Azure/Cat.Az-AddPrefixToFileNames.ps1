<#
.SYNOPSIS
Adds a prefix to every file name in a specified ADLS Gen2 (HNS-enabled) directory using true server-side rename.

.DESCRIPTION
This script renames files in Azure Data Lake Storage Gen2 (ADLS Gen2) by performing a
server-side path rename/move operation. This capability requires a storage account with
Hierarchical Namespace (HNS) enabled.

Background: Why ADLS Gen2 / HNS is required
- Standard Azure Blob Storage (HNS disabled) does not support a true rename operation. Many tools implement
  "rename" as Copy -> Wait for copy success -> Delete old blob.
- ADLS Gen2 (HNS enabled) adds filesystem semantics, including atomic rename/move operations against paths.
  Renames occur server-side without copying the file content.

How this script stays safe
- Uses Move-AzDataLakeGen2Item for true rename.
- Forces destination to remain in the SAME filesystem/container by always setting:
    -DestFileSystem $FileSystemName
  This prevents accidental cross-filesystem moves.
- Supports -WhatIf for dry-run.
- By default, skips if the destination already exists (no overwrite unless you opt in).

REQUIREMENTS
- Storage account: ADLS Gen2 (EnableHierarchicalNamespace = True)
- PowerShell module: Az (Az.Storage)
- Permissions:
  - RBAC: Storage Blob Data Contributor (or higher) at account/filesystem scope
  - AND/OR POSIX ACL permissions on the directory/files (ADLS Gen2)

.PARAMETER ResourceGroupName
Resource group containing the storage account.

.PARAMETER StorageAccountName
Name of the ADLS Gen2-enabled storage account.

.PARAMETER FileSystemName
The ADLS Gen2 filesystem name (container name).

.PARAMETER DirectoryPath
Directory (folder) path within the filesystem to process. Example: "inbox/2026-03"
Use "" (empty string) to target the filesystem root.

.PARAMETER PrefixToAdd
Prefix string to add to each leaf file name. Example: "ARCHIVED-"

.PARAMETER Recurse
If set, processes all files in all subdirectories under DirectoryPath.

.PARAMETER SkipIfAlreadyPrefixed
If set, skips files whose leaf file name already starts with PrefixToAdd.

.PARAMETER OverwriteExisting
If set, overwrites destination if it already exists. Otherwise destinations that exist are skipped.

.PARAMETER MaxErrors
Maximum number of rename errors before the script stops (default 25).

.OUTPUTS
Objects with OldPath/NewPath/Status (and Error if applicable). You can export to CSV.

.EXAMPLE
# Dry-run (safe): show what would happen, do not change anything
.\Cat.Az-AddPrefixToFileNames.ps1 `
  -ResourceGroupName "rg-prac-prod-file-share-migration-discovery-01" `
  -StorageAccountName "stdatamigrationeus2" `
  -FileSystemName "datamigration" `
  -DirectoryPath "bronze/MeritageUserDrives/RawData/Test Rename" `
  -PrefixToAdd "TEST" `
  -Recurse `
  -SkipIfAlreadyPrefixed `
  -WhatIf

.EXAMPLE
# Execute: rename files only in the specified directory (no recursion)
.\Cat.Az-AddPrefixToFileNames.ps1 `
  -ResourceGroupName "rg-prac-prod-file-share-migration-discovery-01" `
  -StorageAccountName "stdatamigrationeus2" `
  -FileSystemName "datamigration" `
  -DirectoryPath "bronze/MeritageUserDrives/RawData/Test Rename" `
  -PrefixToAdd "TEST" `
  -Recurse `
  -SkipIfAlreadyPrefixed

.EXAMPLE
# Execute: recurse and overwrite if destination already exists
.\Cat.Az-AddPrefixToFileNames.ps1 `
  -ResourceGroupName "rg-prac-prod-file-share-migration-discovery-01" `
  -StorageAccountName "stdatamigrationeus2" `
  -FileSystemName "datamigration" `
  -DirectoryPath "bronze/MeritageUserDrives/RawData/Test Rename" `
  -PrefixToAdd "TEST" `
  -Recurse `
  -OverwriteExisting
#>

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = "Medium")]
param(
  [Parameter(Mandatory)] [string] $ResourceGroupName,
  [Parameter(Mandatory)] [string] $StorageAccountName,
  [Parameter(Mandatory)] [string] $FileSystemName,
  [Parameter(Mandatory)] [string] $DirectoryPath,
  [Parameter(Mandatory)] [string] $PrefixToAdd,

  [switch] $Recurse,
  [switch] $SkipIfAlreadyPrefixed,
  [switch] $OverwriteExisting,
  [int] $MaxErrors = 25
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Normalize-AdlsPath {
  param([string]$Path)
  if ($null -eq $Path) { return "" }
  $p = $Path.Trim()
  $p = $p.TrimStart("/")
  $p = $p.TrimEnd("/")
  return $p
}

function Join-AdlsPath {
  param([string]$Left, [string]$Right)
  $l = Normalize-AdlsPath $Left
  $r = Normalize-AdlsPath $Right
  if ([string]::IsNullOrEmpty($l)) { return $r }
  if ([string]::IsNullOrEmpty($r)) { return $l }
  return "$l/$r"
}

function Get-LeafName {
  param([string]$Path)
  if ([string]::IsNullOrEmpty($Path)) { return "" }
  $p = $Path.Replace("\","/")
  return ($p.Split("/") | Select-Object -Last 1)
}

function Get-ParentPath {
  param([string]$Path)
  if ([string]::IsNullOrEmpty($Path)) { return "" }
  $p = $Path.Replace("\","/")
  $idx = $p.LastIndexOf("/")
  if ($idx -lt 0) { return "" }
  return $p.Substring(0, $idx)
}

# ----- Validate inputs -----
if ([string]::IsNullOrWhiteSpace($PrefixToAdd)) {
  throw "PrefixToAdd cannot be empty."
}

$DirectoryPath = Normalize-AdlsPath $DirectoryPath

# ----- Ensure Az module -----
if (-not (Get-Module Az.Storage -ListAvailable)) {
  throw "Az.Storage module not found. Install with: Install-Module Az -Scope CurrentUser"
}

# ----- Interactive login -----
try { Get-AzContext | Out-Null } catch { Connect-AzAccount | Out-Null }

# ----- Get storage account context -----
$acct = Get-AzStorageAccount -ResourceGroupName $ResourceGroupName -Name $StorageAccountName

if (-not $acct.EnableHierarchicalNamespace) {
  throw "HNS is NOT enabled on storage account '$StorageAccountName'. This script requires ADLS Gen2 (HNS=True)."
}

$ctx = $acct.Context

Write-Host "ADLS Gen2 confirmed (HNS=True). Using true rename/move operations." -ForegroundColor Green
Write-Host "Target filesystem: $FileSystemName" -ForegroundColor Cyan
Write-Host "Target directory : $(if($DirectoryPath){$DirectoryPath}else{'(root)'})" -ForegroundColor Cyan
Write-Host "Prefix to add    : $PrefixToAdd" -ForegroundColor Cyan
Write-Host "Recurse          : $Recurse" -ForegroundColor Cyan
Write-Host "OverwriteExisting: $OverwriteExisting" -ForegroundColor Cyan
Write-Host "SkipIfAlreadyPrefixed: $SkipIfAlreadyPrefixed" -ForegroundColor Cyan
Write-Host "WhatIf           : $WhatIfPreference" -ForegroundColor Cyan

# ----- Enumerate items -----
Write-Host "Listing items..." -ForegroundColor Cyan

$items = Get-AzDataLakeGen2ChildItem -Context $ctx -FileSystem $FileSystemName -Path $DirectoryPath -Recurse:$Recurse

# Filter to files only (skip directories)
$files = $items | Where-Object { -not $_.IsDirectory }

if (-not $files -or $files.Count -eq 0) {
  Write-Host "No files found under the specified directory." -ForegroundColor Yellow
  return
}

Write-Host ("Found {0} file(s) to evaluate." -f $files.Count) -ForegroundColor Cyan

# ----- Process renames -----
$results = New-Object System.Collections.Generic.List[object]
$errCount = 0
$i = 0

foreach ($f in $files) {
  $i++

  # In Az cmdlets, the file path is typically in the "Path" property
  $oldPath = $f.Path.Replace("\","/")

  $leaf   = Get-LeafName $oldPath
  $parent = Get-ParentPath $oldPath

  if ($SkipIfAlreadyPrefixed -and $leaf.StartsWith($PrefixToAdd)) {
    $results.Add([pscustomobject]@{
      OldPath = $oldPath
      NewPath = $null
      Status  = "SkippedAlreadyPrefixed"
    })
    continue
  }

  $newLeaf = $PrefixToAdd + $leaf
  $newPath = Join-AdlsPath $parent $newLeaf

  if ($oldPath -eq $newPath) {
    $results.Add([pscustomobject]@{
      OldPath = $oldPath
      NewPath = $newPath
      Status  = "NoChange"
    })
    continue
  }

  # Check for existing destination (safe by default)
  $destExists = $false
  try {
    $null = Get-AzDataLakeGen2Item -Context $ctx -FileSystem $FileSystemName -Path $newPath -ErrorAction Stop
    $destExists = $true
  } catch {
    $destExists = $false
  }

  if ($destExists -and -not $OverwriteExisting) {
    $results.Add([pscustomobject]@{
      OldPath = $oldPath
      NewPath = $newPath
      Status  = "SkippedDestExists"
    })
    continue
  }

  $action = "Rename to '$newPath' (same filesystem enforced)"
  if ($PSCmdlet.ShouldProcess($oldPath, $action)) {
    try {
      # TRUE SERVER-SIDE RENAME (HNS enabled)
      # SAFETY: Enforce SAME filesystem for destination to prevent accidental moves across containers.
      Move-AzDataLakeGen2Item `
        -Context $ctx `
        -FileSystem $FileSystemName `
        -Path $oldPath `
        -DestFileSystem $FileSystemName `
        -DestPath $newPath `
        -Force:$OverwriteExisting | Out-Null

      $results.Add([pscustomobject]@{
        OldPath = $oldPath
        NewPath = $newPath
        Status  = "Renamed"
      })
    }
    catch {
      $errCount++
      $results.Add([pscustomobject]@{
        OldPath = $oldPath
        NewPath = $newPath
        Status  = "Error"
        Error   = $_.Exception.Message
      })

      Write-Warning "Failed: $oldPath -> $newPath : $($_.Exception.Message)"

      if ($errCount -ge $MaxErrors) {
        throw "Reached MaxErrors=$MaxErrors. Stopping."
      }
    }
  }
  else {
    $results.Add([pscustomobject]@{
      OldPath = $oldPath
      NewPath = $newPath
      Status  = "WhatIf"
    })
  }

  if ($i % 200 -eq 0) {
    Write-Host ("Processed {0}/{1}..." -f $i, $files.Count) -ForegroundColor DarkCyan
  }
}

# ----- Summary -----
Write-Host "`nSummary:" -ForegroundColor Cyan
$results | Group-Object Status | Sort-Object Name | ForEach-Object {
  "{0,-24} {1,8}" -f $_.Name, $_.Count
} | ForEach-Object { Write-Host $_ }

# Return details for exporting/reporting
$results