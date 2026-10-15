<#
.SYNOPSIS
  Finds output CSV files whose first N lines reference any UNC path derived from
  input CSV rows where Status = "No" (case-insensitive).

.DESCRIPTION
  - Imports a CSV with Path and Status columns
  - Filters to Status = No
  - Builds one or more compiled regex patterns that match any "\\<Path>" (UNC prefix)
  - Recursively scans all *.csv files under OutputRoot
  - Reads only the first N lines (default 5)
  - Records at most ONE match per file (first match found)
  - Exports results to ReportPath

.PARAMETER InputCsv
  CSV containing Path and Status columns.

.PARAMETER OutputRoot
  Root folder containing output CSVs to scan (recursive).

.PARAMETER ReportPath
  CSV output file listing files that matched.

.PARAMETER LinesToCheck
  Number of lines from the start of each file to scan (default 5).

.PARAMETER RegexChunkSize
  Default 500 is usually safe and fast.

.EXAMPLE
cd C:\GitHub\CatPowerShell\Reporting
.\Cat.Files-CompareDiscoveryErrorsToOutput.ps1 `
    -InputCsv "C:\Users\npatterson.CORP\OneDrive - Quisitive\Clients\Meritage\p220fsm009\2601291019 - p220fsm009\Cat.Files-NetworkShareDiscovery_WithStatus.csv" `
    -OutputRoot "C:\Users\npatterson.CORP\OneDrive - Quisitive\Clients\Meritage\p220fsm009\2601291019 - p220fsm009" `
    -ReportPath "C:\Users\npatterson.CORP\OneDrive - Quisitive\Clients\Meritage\p220fsm009\2601291019 - p220fsm009\MatchesToReview.csv"
#>

[CmdletBinding()]
param(
  [Parameter(Mandatory = $true)]
  [string]$InputCsv,

  [Parameter(Mandatory = $true)]
  [string]$OutputRoot,

  [Parameter(Mandatory = $false)]
  [string]$ReportPath = ".\MatchesToReview.csv",

  [Parameter(Mandatory = $false)]
  [ValidateRange(1,200)]
  [int]$LinesToCheck = 5,

  [Parameter(Mandatory = $false)]
  [ValidateRange(10,5000)]
  [int]$RegexChunkSize = 500
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

Write-Host "Loading input CSV: $InputCsv"
$rows = Import-Csv -Path $InputCsv

# Collect unique paths where Status == "no" (case-insensitive)
$noPaths =
  $rows |
  Where-Object { ($_.Status -as [string]).Trim().ToLowerInvariant() -eq "no" } |
  ForEach-Object { ($_.Path -as [string]).Trim() } |
  Where-Object { $_ -and $_.Length -gt 0 } |
  Sort-Object -Unique

if (-not $noPaths -or $noPaths.Count -eq 0) {
  Write-Warning "No paths found with Status = 'No'. Nothing to scan."
  return
}

Write-Host ("Found {0} 'No' paths to check." -f $noPaths.Count)

# Build escaped regex tokens for "\\<Path>"
# Example Path: 575f10001\users\josh.calera -> token matches \\575f10001\users\josh.calera
$tokens = foreach ($p in $noPaths) {
  [Regex]::Escape("\\$p")
}

# Helper: chunk an array
function Get-Chunks {
  param(
    [Parameter(Mandatory=$true)][object[]]$Array,
    [Parameter(Mandatory=$true)][int]$Size
  )
  for ($i = 0; $i -lt $Array.Count; $i += $Size) {
    $end = [Math]::Min($i + $Size - 1, $Array.Count - 1)
    ,$Array[$i..$end]
  }
}

# Build one or more compiled regex objects from token chunks
$regexOptions = [System.Text.RegularExpressions.RegexOptions]::IgnoreCase `
              -bor [System.Text.RegularExpressions.RegexOptions]::Compiled `
              -bor [System.Text.RegularExpressions.RegexOptions]::CultureInvariant

$regexList = New-Object System.Collections.Generic.List[System.Text.RegularExpressions.Regex]
$chunks = Get-Chunks -Array $tokens -Size $RegexChunkSize

Write-Host ("Compiling {0} regex chunk(s)..." -f $chunks.Count)

foreach ($chunk in $chunks) {
  # Non-capturing group alternation for speed
  $pattern = "(?:" + ($chunk -join "|") + ")"
  $regexList.Add([System.Text.RegularExpressions.Regex]::new($pattern, $regexOptions))
}

Write-Host "Scanning output CSV files under: $OutputRoot"
$files = Get-ChildItem -Path $OutputRoot -Recurse -File -Filter "*.csv" -ErrorAction Stop

Write-Host ("Found {0} CSV files to scan." -f $files.Count)

$results = New-Object System.Collections.Generic.List[object]
$scanned = 0

foreach ($f in $files) {
  $scanned++
  if ($scanned % 200 -eq 0) {
    Write-Host ("Scanned {0}/{1}..." -f $scanned, $files.Count)
  }

  # Read only the first N lines
  $head = $null
  try {
    $head = Get-Content -LiteralPath $f.FullName -TotalCount $LinesToCheck -ErrorAction Stop
  }
  catch {
    # Unreadable file: skip
    continue
  }

  if (-not $head) { continue }

  $matched = $false

  for ($i = 0; $i -lt $head.Count -and -not $matched; $i++) {
    $line = $head[$i]
    if ([string]::IsNullOrWhiteSpace($line)) { continue }

    foreach ($rx in $regexList) {
      $m = $rx.Match($line)
      if ($m.Success) {
        # Record ONLY the first match per file
        $results.Add([pscustomobject]@{
          FileName        = $f.Name
          ParentDirectory = $f.DirectoryName
          FullPath        = $f.FullName
          MatchedUNC      = $m.Value
          LineNumber      = ($i + 1)
          LineText        = $line
        })
        $matched = $true
        break
      }
    }
  }
}

Write-Host ("Done. Found {0} matching file(s)." -f $results.Count)

# Export report
$results |
  Sort-Object ParentDirectory, FileName |
  Export-Csv -Path $ReportPath -NoTypeInformation -Encoding UTF8

Write-Host "Report written to: $ReportPath"
Write-Host "Tip: Review the report and delete matching files before rescanning."