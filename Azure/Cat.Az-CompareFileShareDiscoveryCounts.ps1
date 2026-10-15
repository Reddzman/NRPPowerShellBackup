<#
.SYNOPSIS
    Validates CSV file counts between RawData subfolders and a consolidated folder
    in an Azure Storage Account using ADLS Gen2 / Blob operations.

.DESCRIPTION
    This script compares the total number of CSV files found under:

        RawData/*/<Category>/*.csv

    against the total number of CSV files found under:

        <ConsolidationFolder>/<Category>/*.csv

    It is intended to validate that consolidation or copy operations
    successfully moved all expected files.

    - No file contents are read
    - No modifications are performed
    - Validation is based purely on blob counts by prefix

.PARAMETER StorageAccountName
    Name of the Azure Storage Account.

.PARAMETER ContainerName
    Name of the blob container.

.PARAMETER RawDataRoot
    Root folder containing RawData subfolders.
    Default: 'RawData'

.PARAMETER ConsolidationFolder
    Destination folder name (typically a timestamp such as 2604010823).

.PARAMETER Categories
    One or more category folders to validate.
    Valid values typically include: directories, files, permissions

.EXAMPLE
    .\Cat.Az-CompareFileShareDiscoveryCounts.ps1 `
        -StorageAccountName stdatamigrationeus2 `
        -ContainerName datamigration `
        -ConsolidationFolder bronze/MeritageUserDrives/2604010823 `
        -RawDataRoot bronze/MeritageUserDrives/RawData

.NOTES
    Author: <Your Name>
    Requires: Az.Storage module
    Read-only validation script
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory)]
    [string]$StorageAccountName,

    [Parameter(Mandatory)]
    [string]$ContainerName,

    [Parameter(Mandatory)]
    [string]$RawDataRoot,

    [Parameter(Mandatory)]
    [string]$ConsolidationFolder,

    [string[]]$Categories = @("directories","files","permissions")
)

#region Initialization

Write-Information "Starting ADLS consolidation validation..." -InformationAction Continue

# Create storage context using the connected Azure identity
$ctx = New-AzStorageContext `
    -StorageAccountName $StorageAccountName `
    -UseConnectedAccount

#endregion Initialization

#region Helper Function

function Get-CsvBlobCount {
    param (
        [Parameter(Mandatory)]
        [string]$Prefix,

        [Parameter(Mandatory)]
        [string]$LikePattern
    )

    $blobs = Get-AzStorageBlob `
        -Container $ContainerName `
        -Context $ctx `
        -Prefix $Prefix

    ($blobs | Where-Object { $_.Name -like $LikePattern }).Count
}

#endregion Helper Function

#region Validation Logic

$results = foreach ($category in $Categories) {

    Write-Information "Validating category: $category" -InformationAction Continue

    # RawData: roll up ALL subfolders
    $rawPrefix  = "$RawDataRoot/"
    $rawPattern = "$RawDataRoot/*/$category/*.csv"

    $rawCount = Get-CsvBlobCount `
        -Prefix $rawPrefix `
        -LikePattern $rawPattern

    # Consolidated folder
    $destPrefix = "$ConsolidationFolder/$category/"
    $destPattern = "$destPrefix*.csv"

    $destCount = Get-CsvBlobCount `
        -Prefix $destPrefix `
        -LikePattern $destPattern

    [PSCustomObject]@{
        Category                = $category
        RawDataCsvCount         = $rawCount
        ConsolidatedCsvCount    = $destCount
        Match                   = ($rawCount -eq $destCount)
        Difference              = ($destCount - $rawCount)
    }
}

#endregion Validation Logic

#region Output

Write-Host ""
Write-Host "===== Consolidation Validation Summary =====" -ForegroundColor Cyan
$results | Format-Table -AutoSize

if ($results.Match -contains $false) {
    Write-Warning "One or more categories do NOT match expected file counts."
}
else {
    Write-Host "All categories validated successfully." -ForegroundColor Green
}

#endregion Output