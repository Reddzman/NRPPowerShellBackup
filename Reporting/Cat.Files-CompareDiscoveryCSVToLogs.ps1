<#
.DESCRIPTION
  - Briefly states the purpose of the script
  - Describes the primary inputs and outputs
  - Lists major processing steps in execution order
  - Notes any performance or safety constraints (limits, defaults, assumptions)

.PARAMETER InputPath
  Path to the input file or folder used as the primary data source.

.PARAMETER OutputPath
  Destination path where results or reports are written.

.PARAMETER Filter
  Optional filter used to limit which items are processed.

.PARAMETER Recurse
  Switch indicating whether subfolders should be processed recursively.

.PARAMETER MaxItems
  Maximum number of items to process. Helps prevent excessive runtime.
  Default value is applied if not specified.

.EXAMPLE
cd C:\GitHub\CatPowerShell\Reporting
.\Cat.Files-CompareDiscoveryCSVToLogs.ps1 `
    -CsvPath "C:\Users\npatterson.CORP\OneDrive - Quisitive\Clients\Meritage\p220fsm010\2601291024 - p220fsm010\Cat.Files-NetworkShareDiscovery.csv" `
    -LogPath "C:\Users\npatterson.CORP\OneDrive - Quisitive\Clients\Meritage\p220fsm010\2601291024 - p220fsm010\UserDrives.Files-NetworkShareDiscovery-2601291024.log" `
    -OutputPath "C:\Users\npatterson.CORP\OneDrive - Quisitive\Clients\Meritage\p220fsm010\2601291024 - p220fsm010\Cat.Files-NetworkShareDiscovery_WithStatus.csv"
#>

param (
    [Parameter(Mandatory)]
    [string]$CsvPath,

    [Parameter(Mandatory)]
    [string]$LogPath,

    [Parameter(Mandatory)]
    [string]$OutputPath
)

# --- Load CSV ---
$csv = Import-Csv $CsvPath

<#$csv | Get-Member
$csv.Count
$csvPaths
#>

# Normalize paths once
$csvPaths = $csv.serverUNC | ForEach-Object { $_.Trim() }

# Initialize result table
$results = @{}
foreach ($path in $csvPaths) {
    $results[$path] = [pscustomobject]@{
        Path   = $path
        Status = 'no'     # default pessimistic
        Error  = 'Server block not found in log'
    }
}

# --- Streaming log parse ---
$currentPath = $null
$currentHadError = $false
$currentErrorText = $null

Get-Content $LogPath -ReadCount 500 | ForEach-Object {
    foreach ($line in $_) {

        # Detect new server block
        if ($line -match 'Scanning Server:\s+(.+?)\s+Share:') {

            # Finalize previous block
            if ($currentPath -and $results.ContainsKey($currentPath)) {
                if (-not $currentHadError) {
                    $results[$currentPath].Status = 'yes'
                    $results[$currentPath].Error  = ''
                }
                else {
                    $results[$currentPath].Status = 'no'
                    $results[$currentPath].Error  = $currentErrorText
                }
            }

            # Start new block
            $currentPath = $matches[1].Trim()
            $currentHadError = $false
            $currentErrorText = $null
            continue
        }

        # Capture first ERROR - inside a block
        if ($currentPath -and -not $currentHadError -and $line -match 'ERROR\s*-\s*(.+)$') {
            $currentHadError = $true
            $currentErrorText = $matches[1].Trim()
        }
    }
}

# Finalize last block
if ($currentPath -and $results.ContainsKey($currentPath)) {
    if (-not $currentHadError) {
        $results[$currentPath].Status = 'yes'
        $results[$currentPath].Error  = ''
    }
    else {
        $results[$currentPath].Status = 'no'
        $results[$currentPath].Error  = $currentErrorText
    }
}

# --- Write output ---
#$results.Values | Get-Member
$csvPaths |
    ForEach-Object { $results[$_] } |
    Select-Object Path, Status, Error |
    Export-Csv $OutputPath -NoTypeInformation -Encoding UTF8
