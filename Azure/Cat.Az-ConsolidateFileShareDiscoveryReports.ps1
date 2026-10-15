<#
.SYNOPSIS
    Consolidates CSV discovery reports from distributed ADLS Gen2 source folders
    into a single destination folder structure using safe, incremental copy.

.DESCRIPTION
    Scans a source folder in an ADLS Gen2-enabled Azure Storage filesystem (container).
    The source folder contains multiple subfolders. Each subfolder may include the
    following logical directories:

        - directories
        - files
        - permissions

    Files from ONLY those logical directories are consolidated into a single destination
    folder structure:

        <DestinationPath>/directories
        <DestinationPath>/files
        <DestinationPath>/permissions

    KEY DESIGN NOTE (IMPORTANT)
    ---------------------------
    This script intentionally executes COPY operations SEQUENTIALLY.

    Reason:
      - The Az.Storage cmdlet Copy-AzDataLakeGen2Item is a dynamically generated
        proxy cmdlet and cannot be reliably invoked from PowerShell parallel
        runspaces (ForEach-Object -Parallel, ThreadJobs, Background Jobs).
      - This results in "cmdlet not recognized" errors despite correct module
        loading and versions.

    Sequential execution ensures:
      - Deterministic behavior
      - Full cmdlet availability
      - Correct handling of -WhatIf
      - Reliable reporting and rollup summaries

    Parallelism should be achieved via AzCopy or Azure-native tooling if required.

    Key behaviors:
      - Source files are NEVER deleted
      - Destination files are NEVER overwritten (skip if exists)
      - Script is incremental and safe to re-run
      - -WhatIf is supported for safe planning runs
      - Reports are always generated

.AZ MODULE PATH NOTE
    Your Az.Storage cmdlets require a NON-empty -Path.
    Root operations must use -Path "/".

.AUTHENTICATION
    Uses Azure AD authentication via Az PowerShell.
    SubscriptionId is mandatory for deterministic execution.

.PARAMETER SubscriptionId
    Azure subscription ID (mandatory).

.PARAMETER ResourceGroupName
    Resource group containing the storage account.

.PARAMETER StorageAccountName
    ADLS Gen2-enabled storage account.

.PARAMETER ContainerName
    ADLS Gen2 filesystem (container) name.

.PARAMETER SourcePath
    Root source folder (relative to filesystem).

.PARAMETER DestinationPath
    Root destination folder (relative to filesystem).

.PARAMETER DirectoriesToProcess
    Logical directories to consolidate.
    Default: directories, files, permissions

.PARAMETER OutputReportPath
    Local path for CSV reports.

.PARAMETER CreateDestinationFolders
    Create destination logical folders if missing.

.PARAMETER WhatIf
    Planning mode (no copies performed).

.EXAMPLE
    cd C:\GitHub\CatPowerShell\Azure\
    ./Cat.Az-ConsolidateFileShareDiscoveryReports.ps1 `
      -SubscriptionId "a01e0667-ed24-40ac-9e6e-a6a8793f3686" `
      -ResourceGroupName "rg-prac-prod-file-share-migration-discovery-01" `
      -StorageAccountName "stdatamigrationeus2" `
      -ContainerName "datamigration" `
      -SourcePath "bronze/MeritageUserDrives/RawData" `
      -DestinationPath "bronze/MeritageUserDrives/2604010823" `
      -OutputReportPath "C:\GitHub\CatPowerShell\_Output Files" `
      -CreateDestinationFolders `
      -WhatIf
#>

#region Initialization

using module Az.Accounts
using module Az.Storage

[CmdletBinding(SupportsShouldProcess)]
param (
    [Parameter(Mandatory)][string]$SubscriptionId,
    [Parameter(Mandatory)][string]$ResourceGroupName,
    [Parameter(Mandatory)][string]$StorageAccountName,
    [Parameter(Mandatory)][string]$ContainerName,
    [Parameter(Mandatory)][string]$SourcePath,
    [Parameter(Mandatory)][string]$DestinationPath,
    [string[]]$DirectoriesToProcess = @("directories","files","permissions"),
    [Parameter(Mandatory)][string]$OutputReportPath,
    [string]$tenantID = "d0659de2-684e-49bd-9b1f-1fd4cd0942d9", #Catapult Tenant
    [switch]$CreateDestinationFolders
)

$IsWhatIfRun = $PSBoundParameters.ContainsKey('WhatIf')

Write-Host
Write-Host "====================== Consolidation Started =======================" -ForegroundColor Cyan
Write-Host "✔ SubscriptionId    : $SubscriptionId"
Write-Host "✔ ResourceGroupName : $ResourceGroupName"
Write-Host "✔ StorageAccount    : $StorageAccountName"
Write-Host "✔ Container         : $ContainerName"
Write-Host "✔ SourcePath        : $SourcePath"
Write-Host "✔ DestinationPath   : $DestinationPath"
Write-Host "✔ WhatIf            : $IsWhatIfRun"
Write-Host "====================================================================" -ForegroundColor Cyan
Write-Host

#endregion

#region Helper Functions

function Normalize-AdlsPath {
    param([string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path)) { return "/" }
    ($Path -replace '\\','/').Trim('/')
}

function Get-LeafFromAdlsPath {
    param([string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path)) { return $null }
    ($Path -replace '\\','/').TrimEnd('/').Split('/')[-1]
}

#endregion

#region Azure Context Setup

$SourcePath      = Normalize-AdlsPath $SourcePath
$DestinationPath = Normalize-AdlsPath $DestinationPath

if (-not (Get-AzContext)) {
    Connect-AzAccount | Out-Null
}

Set-AzContext -SubscriptionId $SubscriptionId -WhatIf:$false | Out-Null

$storageAccount = Get-AzStorageAccount `
    -ResourceGroupName $ResourceGroupName `
    -Name $StorageAccountName `
    -ErrorAction Stop

$ctx = New-AzStorageContext `
    -StorageAccountName $StorageAccountName `
    -UseConnectedAccount

New-Item -ItemType Directory -Path $OutputReportPath -Force -WhatIf:$false | Out-Null

$SkippedFiles = [System.Collections.Generic.List[object]]::new()
$Summary      = [System.Collections.Generic.List[object]]::new()

#endregion

#region Azcopy
# Ensure azcopy is available
Write-Host "Authenticating AzCopy..." -ForegroundColor Yellow

# Add this before any AzCopy calls
azcopy login --tenant-id $tenantID
if ($LASTEXITCODE -ne 0) {
    throw "AzCopy login failed. Cannot continue."
}

Write-Host "AzCopy authentication successful." -ForegroundColor Green
Write-Host
#endregion

#region Preflight Validation

Get-AzDataLakeGen2ChildItem -Context $ctx -FileSystem $ContainerName -Path "/" -ErrorAction Stop | Out-Null

if (-not (Get-AzDataLakeGen2Item -Context $ctx -FileSystem $ContainerName -Path $SourcePath -ErrorAction SilentlyContinue)) {
    throw "SourcePath '$SourcePath' does not exist in filesystem '$ContainerName'."
}

if ($CreateDestinationFolders) {
    foreach ($dir in $DirectoriesToProcess) {
        $path = Normalize-AdlsPath "$DestinationPath/$dir"
        New-AzDataLakeGen2Item -Context $ctx -FileSystem $ContainerName -Path $path -Directory -WhatIf:$false -ErrorAction SilentlyContinue | Out-Null
    }
}

#endregion

#region Processing (SEQUENTIAL COPY)

$sourceSubFolders = Get-AzDataLakeGen2ChildItem -Context $ctx -FileSystem $ContainerName -Path $SourcePath | Where-Object IsDirectory

foreach ($sub in $sourceSubFolders) {

    $subPath = $sub.Path
    $subLeaf = Get-LeafFromAdlsPath $subPath
    if (-not $subLeaf) { continue }

    foreach ($logicalDir in $DirectoriesToProcess) {

        $srcDir = Normalize-AdlsPath "$SourcePath/$subLeaf/$logicalDir"
        $dstDir = Normalize-AdlsPath "$DestinationPath/$logicalDir"

        try {
            $files = Get-AzDataLakeGen2ChildItem -Context $ctx -FileSystem $ContainerName -Path $srcDir | Where-Object { -not $_.IsDirectory }
        } catch {
            continue
        }

        $copied = $skipped = $errors = 0
        $totalFiles = $files.Count
        $index = 1

        foreach ($f in $files) {
            $filePath = $f.Path
            $name = Get-LeafFromAdlsPath $filePath
            $destFile = "$dstDir/$name"

            $percent = [math]::Round(($index / $totalFiles) * 100, 1)

            Write-Progress `
                -Activity "Processing $subLeaf" `
                -Status "$logicalDir ($index of $totalFiles)" `
                -PercentComplete $percent
            $index++

            if ($index -eq $totalFiles) {Write-Host ("Processing {0} [{1} ({2} of {3})]" -f $subLeaf, $logicalDir, $index, $totalFiles)}

            if (Get-AzDataLakeGen2Item -Context $ctx -FileSystem $ContainerName -Path $destFile -ErrorAction SilentlyContinue) {
                $skipped++
                $SkippedFiles.Add([pscustomobject]@{
                    SourceSubFolder=$subPath
                    LogicalFolder=$logicalDir
                    FileName=$destFile
                    Reason="File already exists"
                })
                continue
            }

            if ($IsWhatIfRun) {
                Write-Host "WHATIF: Would copy '$filePath' -> '$destFile'"
                $copied++
                continue
            }

            try {
                $srcUrl  = "https://$StorageAccountName.dfs.core.windows.net/$ContainerName/$filePath"
                $destUrl = "https://$StorageAccountName.dfs.core.windows.net/$ContainerName/$destFile"

                if ($IsWhatIfRun) {

                    # WHATIF MODE:
                    # Do NOT run AzCopy. Only show the planned action.
                    Write-Host "WHATIF: Would AzCopy $srcUrl -> $destUrl"
                    $copied++
                }
                else {

                    # EXECUTION MODE:
                    # Run AzCopy and explicitly inspect exit code + output
                    $azOutput = & azcopy copy `
                        $srcUrl `
                        $destUrl `
                        --overwrite=false `
                        --log-level=ERROR `
                        --output-type=text `
                        2>&1

                    $exitCode = $LASTEXITCODE

                    if ($exitCode -eq 0) {

                        # Exit code 0 = either copied OR skipped.
                        # We must inspect output to distinguish.
                        if ($azOutput -match '(?i)skipped') {
                            $skipped++
                            $SkippedFiles.Add([pscustomobject]@{
                                SourceSubFolder = $subPath
                                LogicalFolder   = $logicalDir
                                FileName        = $filePath
                                Reason          = "File already exists (AzCopy overwrite=false)"
                            })
                        }
                        else {
                            $copied++
                        }

                    }
                    else {
                        throw "AzCopy failed (exit code $exitCode). Output: $azOutput"
                    }
                }
            }
            catch {
                $errors++
                $SkippedFiles.Add([pscustomobject]@{
                    SourceSubFolder = $subPath
                    LogicalFolder   = $logicalDir
                    FileName        = $filePath
                    Reason          = $_.Exception.Message
                })
            }
        }

        $Summary.Add([pscustomobject]@{
            SourceSubFolder=$subPath
            LogicalFolder=$logicalDir
            SourceFiles=$files.Count
            CopiedFiles=$copied
            SkippedFiles=$skipped
            ErrorFiles=$errors
        })
    }
}

#endregion

#region Reporting

$Summary | Export-Csv -NoTypeInformation (Join-Path $OutputReportPath "CopySummary.csv") -WhatIf:$false
$SkippedFiles | Export-Csv -NoTypeInformation (Join-Path $OutputReportPath "SkippedFiles.csv") -WhatIf:$false

#endregion

#region Rollup Summary

Write-Host
Write-Host "============================ Consolidation Rollup Summary ============================" -ForegroundColor Cyan

foreach ($dir in $DirectoriesToProcess) {
    $rows = $Summary | Where-Object LogicalFolder -eq $dir
    Write-Host ("{0,-15} Source:{1,6} Copied:{2,6} Skipped:{3,6} Errors:{4,6}" -f `
        $dir,
        ($rows | Measure-Object SourceFiles -Sum).Sum,
        ($rows | Measure-Object CopiedFiles -Sum).Sum,
        ($rows | Measure-Object SkippedFiles -Sum).Sum,
        ($rows | Measure-Object ErrorFiles -Sum).Sum
    )
}

Write-Host "======================================================================================" -ForegroundColor Cyan
Write-Host "✔ Reports:"
Write-Host "  $(Join-Path $OutputReportPath 'CopySummary.csv')"
Write-Host "  $(Join-Path $OutputReportPath 'SkippedFiles.csv')"
Write-Host

#endregion