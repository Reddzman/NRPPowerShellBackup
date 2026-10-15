<#
.SYNOPSIS
Find potential duplicate CSV files across multiple root folders by (Family + FileName),
then compute exact match and percent content overlap using chunk-hash similarity.

.DESCRIPTION
This script:
1) Scans provided root folders for *.csv (optionally recursively)
2) Assigns a "Family" based on filename tokens (e.g., "-Dir-", "-File-", "-Perm-")
3) Groups files by (Family + FileName) (case-insensitive name match)
4) Keeps only groups that appear in 2+ distinct roots (candidate duplicates)
5) Compares each unique pair in each group:
   - If binary SHA256 matches => Identical=True, PercentMatch=100
   - Else compute PercentMatch using chunk-hash overlap:
       SharedChunks / Max(TotalChunksA, TotalChunksB) * 100
     (Also computes Jaccard% = Shared / Union)

.OUTPUTS
- Pairwise CSV: one row per candidate duplicate pair
- Summary CSV: one row per (Family + FileName) group with min/avg/max percent

.NOTES
- Chunk-based comparison does NOT require the files to be valid RFC CSV.
- Reading OneDrive cloud-only files may trigger download.
- Pairwise comparisons per group are n*(n-1)/2; use -MaxFilesPerGroup to control runtime.

.PARAMETER FolderPaths
Root folders to scan.

.PARAMETER Recurse
Scan subfolders recursively.

.PARAMETER OutputDirectory
Directory for CSV reports.

.PARAMETER FamilyTokens
Filename tokens that define families (default: -Dir-, -File-, -Perm-).

.PARAMETER ChunkSizeKB
Chunk size in KB for chunk-hash similarity (default: 64KB). Larger is faster, less granular.

.PARAMETER Algorithm
Hash algorithm for chunk hashes and whole-file hash shortcut (default: SHA256).

.PARAMETER CaseInsensitive
If true (default), match file names case-insensitively.

.PARAMETER MaxFilesPerGroup
Safety limit: if a single (Family+Name) group has more than this many files, it will be skipped.
0 = no limit.

.EXAMPLE
.\Cat.Files-CompareDiscoveryCSVForDuplicates.ps1 `
  -FolderPaths `
    "C:\Users\npatterson.CORP\OneDrive - Quisitive\Clients\Meritage\p220fsm006",
    "C:\Users\npatterson.CORP\OneDrive - Quisitive\Clients\Meritage\p220fsm007",
    "C:\Users\npatterson.CORP\OneDrive - Quisitive\Clients\Meritage\p220fsm008",
    "C:\Users\npatterson.CORP\OneDrive - Quisitive\Clients\Meritage\p220fsm009",
    "C:\Users\npatterson.CORP\OneDrive - Quisitive\Clients\Meritage\p220fsm010" `
  -Recurse `
  -OutputDirectory "C:\Users\npatterson.CORP\OneDrive - Quisitive\Clients\Meritage" `
  -FamilyTokens "-Dir-","-File-","-Perm-" `
  -ChunkSizeKB 64 `
  -Algorithm SHA256 `
  -CaseInsensitive $true `
  -MaxFilesPerGroup 25
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string[]] $FolderPaths,

    [switch] $Recurse,

    [string] $OutputDirectory = (Join-Path $PWD "OneDriveCsvDuplicateReport"),

    [string[]] $FamilyTokens = @("-Dir-","-File-","-Perm-"),

    [ValidateRange(1, 10240)]
    [int] $ChunkSizeKB = 64,

    [ValidateSet("SHA256","SHA1","SHA512","MD5")]
    [string] $Algorithm = "SHA256",

    [bool] $CaseInsensitive = $true,

    [ValidateRange(0, 1000000)]
    [int] $MaxFilesPerGroup = 0
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Resolve-FullPath {
    param([Parameter(Mandatory)][string]$Path)
    (Resolve-Path -LiteralPath $Path).Path
}

function Get-RelativePath {
    param(
        [Parameter(Mandatory)][string]$Base,
        [Parameter(Mandatory)][string]$Full
    )
    $baseNorm = (Resolve-FullPath $Base).TrimEnd('\')
    $fullNorm = (Resolve-FullPath $Full)

    if ($fullNorm.StartsWith($baseNorm, [System.StringComparison]::OrdinalIgnoreCase)) {
        return $fullNorm.Substring($baseNorm.Length).TrimStart('\')
    }
    return $fullNorm
}

function Get-Family {
    param(
        [Parameter(Mandatory)][string]$FileName,
        [Parameter(Mandatory)][string[]]$Tokens
    )
    foreach ($t in $Tokens) {
        if ($FileName.IndexOf($t, [System.StringComparison]::OrdinalIgnoreCase) -ge 0) {
            return $t
        }
    }
    return $null
}

# -----------------------------
# Hash caches (avoid recompute)
# -----------------------------
$WholeFileHashCache = @{}     # FullPath -> Hash (hex) or $null
$ChunkSignatureCache = @{}    # FullPath -> PSCustomObject (dict counts + totalChunks) or $null

function Get-CachedWholeFileHash {
    param([Parameter(Mandatory)][string]$Path, [Parameter(Mandatory)][string]$Algorithm)
    if ($WholeFileHashCache.ContainsKey($Path)) { return $WholeFileHashCache[$Path] }
    try {
        $h = (Get-FileHash -LiteralPath $Path -Algorithm $Algorithm).Hash
        $WholeFileHashCache[$Path] = $h
        return $h
    } catch {
        $WholeFileHashCache[$Path] = $null
        return $null
    }
}

function Convert-BytesToHex {
    param([Parameter(Mandatory)][byte[]]$Bytes)
    $sb = New-Object System.Text.StringBuilder ($Bytes.Length * 2)
    foreach ($b in $Bytes) { [void]$sb.AppendFormat("{0:x2}", $b) }
    $sb.ToString().ToUpperInvariant()
}

function Get-ChunkSignature {
    <#
      Returns a signature object:
      - TotalChunks
      - Dict: Dictionary[string,int] mapping ChunkHash -> count
      Uses streaming reads to avoid loading entire file.
    #>
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][int]$ChunkSizeBytes,
        [Parameter(Mandatory)][string]$Algorithm
    )

    $dict = New-Object 'System.Collections.Generic.Dictionary[string,int]'
    $total = 0

    $hasher = [System.Security.Cryptography.HashAlgorithm]::Create($Algorithm)
    if (-not $hasher) { throw "Unable to create hash algorithm: $Algorithm" }

    $fs = [System.IO.File]::Open($Path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
    try {
        $buffer = New-Object byte[] $ChunkSizeBytes

        while ($true) {
            $read = $fs.Read($buffer, 0, $buffer.Length)
            if ($read -le 0) { break }
            $total++

            # Compute hash of this chunk only (0..read)
            $hashBytes = $hasher.ComputeHash($buffer, 0, $read)
            $key = Convert-BytesToHex -Bytes $hashBytes

            if ($dict.ContainsKey($key)) { $dict[$key]++ } else { $dict[$key] = 1 }
        }
    }
    finally {
        $fs.Dispose()
        $hasher.Dispose()
    }

    return [pscustomobject]@{
        TotalChunks = $total
        Dict        = $dict
    }
}

function Get-CachedChunkSignature {
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][int]$ChunkSizeBytes,
        [Parameter(Mandatory)][string]$Algorithm
    )
    # Chunk size & algorithm affect signature, so include them in the cache key
    $key = "$Path|CS=$ChunkSizeBytes|ALG=$Algorithm"
    if ($ChunkSignatureCache.ContainsKey($key)) { return $ChunkSignatureCache[$key] }

    try {
        $sig = Get-ChunkSignature -Path $Path -ChunkSizeBytes $ChunkSizeBytes -Algorithm $Algorithm
        $ChunkSignatureCache[$key] = $sig
        return $sig
    } catch {
        $ChunkSignatureCache[$key] = $null
        return $null
    }
}

function Compare-ChunkPercent {
    <#
      Computes overlap and Jaccard between two chunk signatures.
      Shared = sum(min(countA, countB))
      Union  = sum(max(countA, countB))
      PercentMatch = Shared / max(TotalA, TotalB) * 100
      JaccardPercent = Shared / Union * 100
    #>
    param(
        [Parameter(Mandatory)]$SigA,
        [Parameter(Mandatory)]$SigB
    )

    $dictA = $SigA.Dict
    $dictB = $SigB.Dict
    $shared = 0
    $union = 0

    foreach ($k in $dictA.Keys) {
        $aCount = $dictA[$k]
        $bCount = if ($dictB.ContainsKey($k)) { $dictB[$k] } else { 0 }
        $shared += [System.Math]::Min($aCount, $bCount)
        $union  += [System.Math]::Max($aCount, $bCount)
    }
    foreach ($k in $dictB.Keys) {
        if (-not $dictA.ContainsKey($k)) {
            $union += $dictB[$k]
        }
    }

    $totalA = [int]$SigA.TotalChunks
    $totalB = [int]$SigB.TotalChunks
    $maxTot = [System.Math]::Max($totalA, $totalB)

    $pct = if ($maxTot -eq 0) { 0.0 } else { [System.Math]::Round(($shared / $maxTot) * 100, 2) }
    $jac = if ($union  -eq 0) { 0.0 } else { [System.Math]::Round(($shared / $union)  * 100, 2) }

    return [pscustomobject]@{
        SharedChunks   = $shared
        UnionChunks    = $union
        TotalChunksA   = $totalA
        TotalChunksB   = $totalB
        PercentMatch   = $pct
        JaccardPercent = $jac
    }
}

# -----------------------------
# Prep output
# -----------------------------
$OutputDirectory = (Resolve-FullPath (New-Item -ItemType Directory -Force -Path $OutputDirectory).FullName)
$timestamp = Get-Date -Format "yyyyMMdd-HHmmss"

$pairwiseCsv = Join-Path $OutputDirectory "CsvDup_Pairwise_$timestamp.csv"
$summaryCsv  = Join-Path $OutputDirectory "CsvDup_Summary_$timestamp.csv"

Write-Host "Output directory: $OutputDirectory" -ForegroundColor Cyan
Write-Host "FamilyTokens: $($FamilyTokens -join ', ')" -ForegroundColor Cyan
Write-Host "ChunkSizeKB : $ChunkSizeKB" -ForegroundColor Cyan
Write-Host "Algorithm   : $Algorithm" -ForegroundColor Cyan

# -----------------------------
# Resolve & validate roots
# -----------------------------
$roots = foreach ($p in $FolderPaths) {
    $rp = Resolve-FullPath $p
    if (-not (Test-Path -LiteralPath $rp -PathType Container)) {
        throw "Folder does not exist or is not a directory: $rp"
    }
    $rp
}

# -----------------------------
# Enumerate candidate files (CSV + family)
# -----------------------------
Write-Host "Scanning CSV files..." -ForegroundColor Cyan

$fileList = New-Object System.Collections.Generic.List[object]
$rootIndex = 0

foreach ($root in $roots) {
    $rootIndex++
    Write-Progress -Activity "Enumerating CSVs" -Status "Root $rootIndex of $($roots.Count): $root" -PercentComplete (($rootIndex / $roots.Count) * 100)

    $gciParams = @{
        LiteralPath = $root
        File        = $true
        Force       = $true
        Filter      = "*.csv"
    }
    if ($Recurse) { $gciParams.Recurse = $true }

    Get-ChildItem @gciParams | ForEach-Object {
        $fam = Get-Family -FileName $_.Name -Tokens $FamilyTokens
        if (-not $fam) { return }

        $nameKey = if ($CaseInsensitive) { $_.Name.ToLowerInvariant() } else { $_.Name }

        $fileList.Add([pscustomobject]@{
            Root         = $root
            Family       = $fam
            Name         = $_.Name
            NameKey      = $nameKey
            FullPath     = $_.FullName
            RelativePath = Get-RelativePath -Base $root -Full $_.FullName
            Length       = $_.Length
            LastWriteUtc = $_.LastWriteTimeUtc
        })
    }
}
Write-Progress -Activity "Enumerating CSVs" -Completed

if ($fileList.Count -eq 0) {
    Write-Warning "No CSV files found that match the specified family tokens."
    return
}

# -----------------------------
# Group by (Family + NameKey), keep only groups spanning 2+ roots (candidate duplicates)
# -----------------------------
Write-Host "Selecting candidate duplicate groups (Family + FileName across roots)..." -ForegroundColor Cyan

$groups = $fileList |
    Group-Object -Property { "$($_.Family)|$($_.NameKey)" } |
    Where-Object {
        $_.Count -gt 1 -and
        ($_.Group | Select-Object -ExpandProperty Root -Unique).Count -gt 1
    }

if (-not $groups -or $groups.Count -eq 0) {
    Write-Host "No candidate duplicates found (same family + same file name across different roots)." -ForegroundColor Yellow
    return
}

# -----------------------------
# Compare pairwise within each candidate group
# -----------------------------
Write-Host "Comparing candidate duplicates (pairwise)..." -ForegroundColor Cyan

$chunkSizeBytes = $ChunkSizeKB * 1KB

$pairwiseRows = New-Object System.Collections.Generic.List[object]
$summaryRows  = New-Object System.Collections.Generic.List[object]

$groupId = 0
foreach ($g in $groups) {
    $groupId++
    $items = $g.Group

    if ($MaxFilesPerGroup -gt 0 -and $items.Count -gt $MaxFilesPerGroup) {
        $summaryRows.Add([pscustomobject]@{
            GroupId           = $groupId
            Family            = $items[0].Family
            FileName          = $items[0].Name
            Instances         = $items.Count
            DistinctRoots     = ($items | Select-Object -ExpandProperty Root -Unique).Count
            ComparisonsDone   = 0
            MinPercentMatch   = $null
            AvgPercentMatch   = $null
            MaxPercentMatch   = $null
            AnyExactMatch     = $false
            Note              = "Skipped: group exceeds MaxFilesPerGroup=$MaxFilesPerGroup"
        })
        continue
    }

    Write-Progress -Activity "Pairwise comparisons" -Status "Group $groupId of $($groups.Count): $($items[0].Name)" -PercentComplete (($groupId / $groups.Count) * 100)

    $pctList = New-Object System.Collections.Generic.List[double]
    $anyExact = $false
    $comparisonCount = 0

    for ($i = 0; $i -lt $items.Count; $i++) {
        for ($j = $i + 1; $j -lt $items.Count; $j++) {
            $a = $items[$i]
            $b = $items[$j]

            $comparisonCount++

            # Whole-file shortcut
            $ha = $null
            $hb = $null
            $identical = $false
            $percent = $null
            $jaccard = $null
            $shared = $null
            $union = $null
            $chunksA = $null
            $chunksB = $null
            $method = $null
            $note = $null

            if ($a.Length -eq $b.Length) {
                $ha = Get-CachedWholeFileHash -Path $a.FullPath -Algorithm $Algorithm
                $hb = Get-CachedWholeFileHash -Path $b.FullPath -Algorithm $Algorithm
                if ($ha -and $hb -and ($ha -eq $hb)) {
                    $identical = $true
                    $percent = 100.0
                    $jaccard = 100.0
                    $method = "WholeFileHash"
                    $anyExact = $true
                }
            }

            if (-not $identical) {
                # Chunk signature similarity
                $sigA = Get-CachedChunkSignature -Path $a.FullPath -ChunkSizeBytes $chunkSizeBytes -Algorithm $Algorithm
                $sigB = Get-CachedChunkSignature -Path $b.FullPath -ChunkSizeBytes $chunkSizeBytes -Algorithm $Algorithm

                if ($sigA -eq $null -or $sigB -eq $null) {
                    $note = "FailedToReadOrHash"
                    $method = "ChunkHash"
                } else {
                    $cmp = Compare-ChunkPercent -SigA $sigA -SigB $sigB
                    $shared  = $cmp.SharedChunks
                    $union   = $cmp.UnionChunks
                    $chunksA = $cmp.TotalChunksA
                    $chunksB = $cmp.TotalChunksB
                    $percent = $cmp.PercentMatch
                    $jaccard = $cmp.JaccardPercent
                    $identical = ($percent -eq 100.0)
                    $method = "ChunkHash"
                    if ($identical) { $anyExact = $true }
                }
            }

            if ($percent -ne $null) { [void]$pctList.Add([double]$percent) }

            $pairwiseRows.Add([pscustomobject]@{
                GroupId        = $groupId
                Family         = $items[0].Family
                FileName       = $items[0].Name

                RootA          = $a.Root
                RelativePathA  = $a.RelativePath
                FullPathA      = $a.FullPath
                SizeBytesA     = $a.Length
                LastWriteUtcA  = $a.LastWriteUtc

                RootB          = $b.Root
                RelativePathB  = $b.RelativePath
                FullPathB      = $b.FullPath
                SizeBytesB     = $b.Length
                LastWriteUtcB  = $b.LastWriteUtc

                Identical      = $identical
                PercentMatch   = $percent
                JaccardPercent = $jaccard
                SharedChunks   = $shared
                UnionChunks    = $union
                TotalChunksA   = $chunksA
                TotalChunksB   = $chunksB

                ChunkSizeKB    = $ChunkSizeKB
                Algorithm      = $Algorithm
                Method         = $method
                Note           = $note
            })
        }
    }

    $minPct = if ($pctList.Count -gt 0) { [System.Math]::Round(($pctList | Measure-Object -Minimum).Minimum, 2) } else { $null }
    $avgPct = if ($pctList.Count -gt 0) { [System.Math]::Round(($pctList | Measure-Object -Average).Average, 2) } else { $null }
    $maxPct = if ($pctList.Count -gt 0) { [System.Math]::Round(($pctList | Measure-Object -Maximum).Maximum, 2) } else { $null }

    $summaryRows.Add([pscustomobject]@{
        GroupId           = $groupId
        Family            = $items[0].Family
        FileName          = $items[0].Name
        Instances         = $items.Count
        DistinctRoots     = ($items | Select-Object -ExpandProperty Root -Unique).Count
        ComparisonsDone   = $comparisonCount
        MinPercentMatch   = $minPct
        AvgPercentMatch   = $avgPct
        MaxPercentMatch   = $maxPct
        AnyExactMatch     = $anyExact
        ChunkSizeKB       = $ChunkSizeKB
        Algorithm         = $Algorithm
        Note              = $null
    })
}

Write-Progress -Activity "Pairwise comparisons" -Completed

# -----------------------------
# Export only candidate duplicates (by design)
# -----------------------------
$pairwiseRows |
    Sort-Object GroupId, Family, FileName, PercentMatch -Descending |
    Export-Csv -NoTypeInformation -Encoding UTF8 -Path $pairwiseCsv

$summaryRows |
    Sort-Object GroupId |
    Export-Csv -NoTypeInformation -Encoding UTF8 -Path $summaryCsv

Write-Host "`nDONE" -ForegroundColor Green
Write-Host "Pairwise report: $pairwiseCsv" -ForegroundColor Green
Write-Host "Summary report : $summaryCsv" -ForegroundColor Green