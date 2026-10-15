param (
    [string]$InputCsv = "UNCPaths.csv",       # CSV file with a column named 'Path'
    [string]$OutputCsv = "UNCAnalysisResults.csv"
)

# Ensure input file exists
if (-Not (Test-Path $InputCsv)) {
    Write-Host "Input CSV file not found: $InputCsv"
    exit
}

# Read UNC paths from CSV
$uncPaths = Import-Csv -Path $InputCsv

# Prepare results array
$results = @()

foreach ($entry in $uncPaths) {
    $path = $entry.Path
    $totalSize = 0
    $totalFolders = 0
    $totalFiles = 0
    $rootContainsFiles = "No"

    if (Test-Path $path) {
        $items = Get-ChildItem -Path $path -Recurse -Force -ErrorAction SilentlyContinue

        foreach ($item in $items) {
            if ($item.PSIsContainer) {
                $totalFolders++
            } else {
                $totalFiles++
                $totalSize += $item.Length
            }
        }

        $rootFiles = Get-ChildItem -Path $path -File -Force -ErrorAction SilentlyContinue
        if ($rootFiles.Count -gt 0) {
            $rootContainsFiles = "Yes"
        }

        $sizeGB = [math]::Round($totalSize / 1GB, 4)
    } else {
        $sizeGB = "Path not found"
        $totalFolders = "-"
        $totalFiles = "-"
        $rootContainsFiles = "-"
    }

    $results += [PSCustomObject]@{
        Path              = $path
        SizeGB            = $sizeGB
        FileCount         = $totalFiles
        FolderCount       = $totalFolders
        RootHasFiles      = $rootContainsFiles
    }
}

# Export results to CSV
$results | Export-Csv -Path $OutputCsv -NoTypeInformation
Write-Host "Analysis complete. Results saved to $OutputCsv"