<#
.SYNOPSIS
    Compares source network folder item counts against migrated OneDrive content under the destination folder and exports an Excel validation report.

.DESCRIPTION
    Reads a migration mapping CSV that contains the same columns used by the ShareGate/PowerShell migration process.
    For each row, the script validates:
        - Source network folder counts and size
        - Destination OneDrive folder counts and size under the specified destFolder value
        - Item, folder, file, and size discrepancies
        - Missing source items that are not found in the destination OneDrive folder, when -IncludeMissingItems is used

    The workbook includes:
        1. Dashboard
        2. ValidationResults
        3. DiscrepancyUsers
        4. MissingItems
        5. Errors
        6. RunInfo

.OUTPUTS

    Excel Workbook:
        Dashboard
        ValidationResults
        DiscrepancyUsers
        MissingItems
        Errors
        RunInfo

    Supporting Files:
        *.checkpoint.csv
            Cached validation results used by -Resume

        *.errors.csv
            Detailed validation and connection failures

        *.missing-items.csv
            Source items not found in OneDrive

        *.runinfo.csv
            Execution settings and runtime information

        *.log
            Detailed execution log

.NOTES
    Used for the Sanford Health N Drive to OneDrive migration validation. Can be used for other network folder to OneDrive migrations with the same CSV format.
    PowerShell: Tested syntax target is PowerShell 7.x / Windows PowerShell 5.1 with PnP.PowerShell and ImportExcel installed.

.REQUIRED INPUT CSV COLUMNS
    Wave,DisplayName,srcTitle,srcUPN,TargetUPN,Email,srcUserName,Network Folder Size (GB),srcUNCPath,destSiteTitle,destSiteURL,destGroupAlias,destListName,destListPath,destFolder,siteTemplate,CommSiteDesign,siteOwner,modernTeamSiteIsPublic,isHubSite,hubParentURL,classicSubSiteParentURL,destChannel,destChannelType,customSPTemplate,onWarning,onError,onSiteObjectExists,onContentItemExists,noCustomPermissions,normalMode,waitForCompletion,filterStartDate

.EXAMPLE
    # Full validation with missing item path tab
    .\Validate--NetworkFolderMigration.ps1 `
        -InputCsv .\Pilot-Wave.csv `
        -OutputXlsx .\Pilot_Validation_Report.xlsx `
        -TargetFolderName "Migrated_NDrive"
        -Interactive `
        -IncludeMissingItems

        Performs a comprehensive migration validation against the target folder in OneDrive
        locations listed in the migration mapping CSV.

        The script:
        - Scans the source network folder structure.
        - Scans the destination OneDrive folder.
        - Compares total items, folders, files, and content size.
        - Identifies discrepancies between source and destination.
        - Generates a MissingItems worksheet containing source files and folders
            that were not found in the destination inventory.

        Recommended for pilot migrations, migration sign-off activities,
        post-wave validation, and troubleshooting data loss concerns.

.EXAMPLE
    # Quicker validation - counts only, no missing path comparison
    .\Validate--NetworkFolderMigration.ps1 `
        -InputCsv .\Pilot-Wave.csv `
        -OutputXlsx .\Pilot_Validation_Report_Fast.xlsx `
        -TargetFolderName "Migrated_NDrive"
        -Interactive `
        -FastCountsOnly

        Performs a high-speed validation that only compares aggregate metrics
        between the source and destination locations.

        The script:
        - Compares item counts.
        - Compares folder counts.
        - Compares file counts.
        - Compares total size.
        - Does not build detailed item inventories.
        - Does not generate a MissingItems worksheet.

        Recommended when validating hundreds or thousands of users where a quick
        confidence check is needed before running detailed discrepancy analysis.

.EXAMPLE
    # App-only auth, useful for unattended wave validation if app permissions are approved
    .\Validate--NetworkFolderMigration.ps1 `
        -InputCsv .\Batch1.csv `
        -OutputXlsx .\Batch1_Validation.xlsx `
        -TargetFolderName "Migrated_NDrive"
        -Tenant "CLIENTDOMAIN.onmicrosoft.com" `
        -ClientId "YYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYY" `
        -CertificateThumbprint "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
        -IncludeMissingItems

        Performs a full migration validation using certificate-based
        authentication instead of interactive sign-in.

        This method is intended for:
        - Scheduled validation jobs
        - Azure Automation
        - Overnight validation of large migration waves
        - Service accounts operating without user interaction

        Generates the same validation outputs as interactive mode while allowing
        the process to run unattended.

.EXAMPLE
    # Resume a previously interrupted validation
    .\Validate--NetworkFolderMigration.ps1 `
        -InputCsv .\Batch2.csv `
        -OutputXlsx .\Batch2_Validation.xlsx `
        -TargetFolderName "Migrated_NDrive"
        -Interactive `
        -Resume

    Resumes a previously started validation run using the checkpoint file
    generated in the output folder.

    Any users that were successfully processed during earlier executions
    are skipped, allowing validation to continue from the last completed row.

    Useful when:
      - A workstation reboots
      - Authentication expires
      - The script is stopped accidentally
      - Large validation waves take multiple sessions to complete


.EXAMPLE
    # Generate a discrepancy-focused report
    .\Validate--NetworkFolderMigration.ps1 `
        -InputCsv .\MFLD_Batch3.csv `
        -OutputXlsx .\MFLD_Batch3_Discrepancies.xlsx `
        -TargetFolderName "Migrated_NDrive"
        -Interactive `
        -DiscrepancyRerunFilter AnyDiscrepancy

    Produces a report where the DiscrepancyUsers worksheet contains only
    users with validation differences or processing errors.

    Helps migration administrators quickly review users requiring
    investigation without manually filtering the full ValidationResults
    worksheet.

.EXAMPLE
    # Focus only on item count differences
    .\Validate--NetworkFolderMigration.ps1 `
        -InputCsv .\MFLD_Batch3.csv `
        -OutputXlsx .\MFLD_CountDiscrepancies.xlsx `
        -TargetFolderName "Migrated_NDrive"
        -Interactive `
        -DiscrepancyRerunFilter CountDiscrepancy

    Filters the DiscrepancyUsers worksheet to only include users where the
    total source item count differs from the destination item count.

    Useful for identifying migrations where content may be missing without
    reviewing size or folder differences.

.EXAMPLE
    # Automatically install required PowerShell modules
    .\Validate--NetworkFolderMigration.ps1 `
        -InputCsv .\Pilot.csv `
        -OutputXlsx .\PilotValidation.xlsx `
        -TargetFolderName "Migrated_NDrive"
        -Interactive `
        -InstallMissingModules

    Checks for required modules and automatically installs:
      - PnP.PowerShell
      - ImportExcel

    Useful for first-time execution on a new workstation, migration server,
    or administrator workstation.

.EXAMPLE
    # Validate against a custom target folder
    .\Validate--NetworkFolderMigration.ps1 `
        -InputCsv .\Batch1.csv `
        -OutputXlsx .\Batch1Validation.xlsx `
        -Interactive `
        -TargetFolderName "Migrated_NDrive"

    Useful when migration waves use different folder naming conventions
    or when validating test migrations.

.EXAMPLE
    # Validate content stored in a non-standard library
    .\Validate--NetworkFolderMigration.ps1 `
        -InputCsv .\Batch1.csv `
        -OutputXlsx .\Batch1Validation.xlsx `
        -Interactive `
        -DefaultDocumentLibraryName "Shared Documents"

    Directs validation to a document library other than the default
    "Documents" library.

    Useful for tenants with customized library names or legacy site
    configurations.

#>

[CmdletBinding(DefaultParameterSetName = 'Interactive')]
param(
    [Parameter(Mandatory)]
    [ValidateScript({ Test-Path -LiteralPath $_ })]
    [string]$InputCsv,

    [Parameter(Mandatory)]
    [string]$OutputXlsx,

    [Parameter(ParameterSetName = 'Interactive')]
    [switch]$Interactive,

    [Parameter(ParameterSetName = 'AppOnly', Mandatory)]
    [string]$Tenant,

    [Parameter(ParameterSetName = 'AppOnly', Mandatory)]
    [string]$ClientId,

    [Parameter(ParameterSetName = 'AppOnly', Mandatory)]
    [string]$CertificateThumbprint,

    [string]$TargetFolderName = 'MFLD_N_Drive',

    [string]$DefaultDocumentLibraryName = 'Documents',

    [ValidateRange(100,5000)]
    [int]$PnPPageSize = 2000,

    [ValidateRange(1,3)]
    [int]$RetryCount = 2,

    [ValidateRange(1,120)]
    [int]$RetryDelaySeconds = 10,

    [ValidateSet('All','CountDiscrepancy','AnyDiscrepancy','None')]
    [string]$DiscrepancyRerunFilter = 'AnyDiscrepancy',

    [switch]$IncludeMissingItems,

    [switch]$FastCountsOnly,

    [switch]$Resume,

    [switch]$InstallMissingModules
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

#region Output paths
$runStarted = Get-Date
$outputFolder = Split-Path -Path $OutputXlsx -Parent
if ([string]::IsNullOrWhiteSpace($outputFolder)) { $outputFolder = (Get-Location).Path }
if (-not (Test-Path -LiteralPath $outputFolder)) { New-Item -ItemType Directory -Path $outputFolder -Force | Out-Null }
$baseName = [System.IO.Path]::GetFileNameWithoutExtension($OutputXlsx)
$checkpointCsv = Join-Path $outputFolder "$baseName.checkpoint.csv"
$errorCsv      = Join-Path $outputFolder "$baseName.errors.csv"
$missingCsv    = Join-Path $outputFolder "$baseName.missing-items.csv"
$runInfoCsv    = Join-Path $outputFolder "$baseName.runinfo.csv"
$logFile       = Join-Path $outputFolder "$baseName.log"
#endregion Output paths

#region Logging
function Write-Log {
    param(
        [ValidateSet('INFO','WARN','ERROR','DEBUG')][string]$Level,
        [string]$Message
    )
    $line = "{0} [{1}] {2}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Message
    Add-Content -Path $logFile -Value $line -Encoding UTF8
    switch ($Level) {
        'ERROR' { Write-Host $line -ForegroundColor Red }
        'WARN'  { Write-Host $line -ForegroundColor Yellow }
        default { Write-Host $line }
    }
}

function Add-ErrorRecord {
    param(
        [string]$Wave,
        [string]$DisplayName,
        [string]$srcUPN,
        [string]$TargetUPN,
        [string]$srcUNCPath,
        [string]$destSiteURL,
        [string]$Stage,
        [string]$Message
    )
    [PSCustomObject]@{
        Timestamp   = Get-Date
        Wave        = $Wave
        DisplayName = $DisplayName
        srcUPN      = $srcUPN
        TargetUPN   = $TargetUPN
        srcUNCPath  = $srcUNCPath
        destSiteURL = $destSiteURL
        Stage       = $Stage
        Error       = $Message
    } | Export-Csv -Path $errorCsv -NoTypeInformation -Append -Encoding UTF8
}
#endregion Logging

#region Module checks
function Ensure-Module {
    param([string]$Name)
    if (-not (Get-Module -ListAvailable -Name $Name)) {
        if ($InstallMissingModules) {
            Write-Log -Level INFO -Message "Installing missing module $Name for current user."
            Install-Module -Name $Name -Scope CurrentUser -Force -AllowClobber
        }
        else {
            throw "Required module '$Name' is not installed. Install it first or rerun with -InstallMissingModules."
        }
    }
    Import-Module $Name -ErrorAction Stop
}

Ensure-Module -Name PnP.PowerShell
Ensure-Module -Name ImportExcel
#endregion Module checks

#region Helpers
function Validate-WithRetry {
    param(
        [scriptblock]$ScriptBlock,
        [string]$OperationName,
        [object]$Row
    )
    for ($attempt = 1; $attempt -le $RetryCount; $attempt++) {
        try { return & $ScriptBlock }
        catch {
            $msg = $_.Exception.Message
            Write-Log -Level WARN -Message "$OperationName failed on attempt $attempt of $RetryCount for $($Row.DisplayName). $msg"
            if ($attempt -eq $RetryCount) {
                Add-ErrorRecord -Wave $Row.Wave -DisplayName $Row.DisplayName -srcUPN $Row.srcUPN -TargetUPN $Row.TargetUPN -srcUNCPath $Row.srcUNCPath -destSiteURL $Row.destSiteURL -Stage $OperationName -Message $msg
                throw
            }
            Start-Sleep -Seconds $RetryDelaySeconds
        }
    }
}

function Normalize-RelativePath {
    param([string]$Path)
    if ($null -eq $Path) { return '' }
    $p = $Path -replace '\\','/'
    $p = $p.Trim('/')
    return $p.ToLowerInvariant()
}

function ConvertTo-ServerRelativeUrl {
    param([string]$SiteUrl, [string]$LibraryName, [string]$FolderName)
    $uri = [Uri]$SiteUrl
    $sitePath = [System.Uri]::UnescapeDataString($uri.AbsolutePath.TrimEnd('/'))
    $lib = if ([string]::IsNullOrWhiteSpace($LibraryName)) { $DefaultDocumentLibraryName } else { $LibraryName }
    $folder = if ([string]::IsNullOrWhiteSpace($FolderName)) { $TargetFolderName } else { $FolderName }
    return ("{0}/{1}/{2}" -f $sitePath, $lib.Trim('/'), $folder.Trim('/'))
}

function Escape-CamlValue {
    param([string]$Value)
    return [System.Security.SecurityElement]::Escape($Value)
}

function Get-LibraryNameFromRow {
    param([object]$Row)
    if ($Row.PSObject.Properties.Name -contains 'destListName' -and -not [string]::IsNullOrWhiteSpace($Row.destListName)) { return $Row.destListName }
    if ($Row.PSObject.Properties.Name -contains 'destListPath' -and -not [string]::IsNullOrWhiteSpace($Row.destListPath)) {
        $leaf = ($Row.destListPath -replace '\\','/').Trim('/').Split('/')[-1]
        if ($leaf) { return $leaf }
    }
    return $DefaultDocumentLibraryName
}

function Get-FolderNameFromRow {
    param([object]$Row)
    if ($Row.PSObject.Properties.Name -contains 'destFolder' -and -not [string]::IsNullOrWhiteSpace($Row.destFolder)) { return $Row.destFolder }
    return $TargetFolderName
}
#endregion Helpers

#region Source scan
function Get-SourceInventory {
    param(
        [string]$RootPath,
        [switch]$CollectItems
    )

    $result = [ordered]@{
        Exists          = $false
        ItemCount       = 0L
        FileCount       = 0L
        FolderCount     = 0L
        TotalSizeBytes  = 0L
        AccessErrors    = 0L
        ErrorMessages   = New-Object System.Collections.Generic.List[string]
        Items           = if ($CollectItems) { New-Object 'System.Collections.Generic.HashSet[string]' } else { $null }
        ItemRows        = if ($CollectItems) { New-Object System.Collections.Generic.List[object] } else { $null }
    }

    if (-not (Test-Path -LiteralPath $RootPath)) {
        $result.ErrorMessages.Add('Source path not found') | Out-Null
        return [PSCustomObject]$result
    }

    $result.Exists = $true
    $rootFull = [System.IO.Path]::GetFullPath($RootPath).TrimEnd('\')
    $stack = New-Object System.Collections.Generic.Stack[string]
    $stack.Push($rootFull)

    while (@($stack).Count -gt 0) {
        $current = $stack.Pop()

        try {
            foreach ($dir in [System.IO.Directory]::EnumerateDirectories($current)) {
                $result.FolderCount++
                $result.ItemCount++
                $rel = $dir.Substring($rootFull.Length).TrimStart('\')
                if ($CollectItems) {
                    $norm = Normalize-RelativePath $rel
                    [void]$result.Items.Add(('Folder|' + $norm))
                    $result.ItemRows.Add([PSCustomObject]@{ RelativePath=$rel; ObjectType='Folder'; SourceFullPath=$dir; SizeBytes=0 }) | Out-Null
                }
                $stack.Push($dir)
            }
        }
        catch {
            $result.AccessErrors++
            $result.ErrorMessages.Add("Folder enumerate error: $current :: $($_.Exception.Message)") | Out-Null
        }

        try {
            foreach ($file in [System.IO.Directory]::EnumerateFiles($current)) {
                $result.FileCount++
                $result.ItemCount++
                $size = 0L
                try { $size = ([System.IO.FileInfo]$file).Length } catch { $size = 0L }
                $result.TotalSizeBytes += $size
                $rel = $file.Substring($rootFull.Length).TrimStart('\')
                if ($CollectItems) {
                    $norm = Normalize-RelativePath $rel
                    [void]$result.Items.Add(('File|' + $norm))
                    $result.ItemRows.Add([PSCustomObject]@{ RelativePath=$rel; ObjectType='File'; SourceFullPath=$file; SizeBytes=$size }) | Out-Null
                }
            }
        }
        catch {
            $result.AccessErrors++
            $result.ErrorMessages.Add("File enumerate error: $current :: $($_.Exception.Message)") | Out-Null
        }
    }

    return [PSCustomObject]$result
}
#endregion Source scan

#region Destination scan
function Connect-ToOneDriveSite {
    param([string]$Url)
    if ($PSCmdlet.ParameterSetName -eq 'AppOnly') {
        return Connect-PnPOnline -Url $Url -Tenant $Tenant -ClientId $ClientId -Thumbprint $CertificateThumbprint -ReturnConnection
    }
    else {
        return Connect-PnPOnline -Url $Url -Interactive -ReturnConnection
    }
}

function Get-OneDriveFolderInventory {
    param(
        [object]$Row,
        [string]$SiteUrl,
        [string]$ListName,
        [string]$FolderName,
        [switch]$CollectItems
    )

    $result = [ordered]@{
        Exists          = $false
        ItemCount       = 0L
        FileCount       = 0L
        FolderCount     = 0L
        TotalSizeBytes  = 0L
        FolderUrl       = $null
        Items           = if ($CollectItems) { New-Object 'System.Collections.Generic.HashSet[string]' } else { $null }
        ErrorMessages   = New-Object System.Collections.Generic.List[string]
    }

    $folderServerRelativeUrl = ConvertTo-ServerRelativeUrl -SiteUrl $SiteUrl -LibraryName $ListName -FolderName $FolderName
    $result.FolderUrl = $folderServerRelativeUrl

    $connection = Validate-WithRetry -Row $Row -OperationName 'Connect-PnPOnline' -ScriptBlock { Connect-ToOneDriveSite -Url $SiteUrl }

    $escapedUrl = Escape-CamlValue $folderServerRelativeUrl
    $query = @"
<View Scope='RecursiveAll'>
  <Query>
    <Where>
      <BeginsWith>
        <FieldRef Name='FileRef' />
        <Value Type='Text'>$escapedUrl</Value>
      </BeginsWith>
    </Where>
  </Query>
  <ViewFields>
    <FieldRef Name='FileRef' />
    <FieldRef Name='FileLeafRef' />
    <FieldRef Name='FSObjType' />
    <FieldRef Name='File_x0020_Size' />
  </ViewFields>
  <RowLimit Paged='TRUE'>$PnPPageSize</RowLimit>
</View>
"@

    $items = Validate-WithRetry -Row $Row -OperationName 'Get-PnPListItem' -ScriptBlock {
        Get-PnPListItem -List $ListName -Query $query -PageSize $PnPPageSize -Connection $connection
    }

    foreach ($item in $items) {
        $fileRef = [string]$item.FieldValues.FileRef
        if ([string]::IsNullOrWhiteSpace($fileRef)) { continue }
        if ($fileRef.TrimEnd('/') -ieq $folderServerRelativeUrl.TrimEnd('/')) {
            $result.Exists = $true
            continue
        }

        $result.Exists = $true
        $fsObjType = [int]$item.FieldValues.FSObjType
        $isFolder = ($fsObjType -eq 1)
        $result.ItemCount++
        if ($isFolder) { $result.FolderCount++ } else { $result.FileCount++ }
        $size = 0L
        if (-not $isFolder -and $item.FieldValues.ContainsKey('File_x0020_Size') -and $null -ne $item.FieldValues.File_x0020_Size) {
            [void][int64]::TryParse([string]$item.FieldValues.File_x0020_Size, [ref]$size)
            $result.TotalSizeBytes += $size
        }

        if ($CollectItems) {
            $rel = $fileRef.Substring($folderServerRelativeUrl.TrimEnd('/').Length).TrimStart('/')
            $norm = Normalize-RelativePath $rel
            $keyType = if ($isFolder) { 'Folder' } else { 'File' }
            [void]$result.Items.Add(($keyType + '|' + $norm))
        }
    }

    return [PSCustomObject]$result
}
#endregion Destination scan

#region Load inputs and resume
$requiredColumns = @('Wave','DisplayName','srcTitle','srcUPN','TargetUPN','Email','srcUserName','Network Folder Size (GB)','srcUNCPath','destSiteTitle','destSiteURL','destGroupAlias','destListName','destListPath','destFolder','siteTemplate','CommSiteDesign','siteOwner','modernTeamSiteIsPublic','isHubSite','hubParentURL','classicSubSiteParentURL','destChannel','destChannelType','customSPTemplate','onWarning','onError','onSiteObjectExists','onContentItemExists','noCustomPermissions','normalMode','waitForCompletion','filterStartDate')
$rows = @(Import-Csv -Path $InputCsv)
if (@($rows).Count -eq 0) { throw 'Input CSV contains no rows.' }
$actualColumns = @($rows[0].PSObject.Properties.Name)
$missingRequired = @($requiredColumns | Where-Object { $_ -notin $actualColumns })
if (@($missingRequired).Count -gt 0) { throw "Input CSV is missing required columns: $($missingRequired -join ', ')" }

$processedKeys = New-Object 'System.Collections.Generic.HashSet[string]'
if ($Resume -and (Test-Path -LiteralPath $checkpointCsv)) {
    foreach ($done in Import-Csv -Path $checkpointCsv) {
        [void]$processedKeys.Add(($done.Wave + '|' + $done.srcUPN + '|' + $done.srcUNCPath))
    }
    Write-Log -Level INFO -Message "Resume enabled. Found $(@($processedKeys).Count) completed rows in checkpoint."
}
elseif (-not $Resume) {
    Remove-Item -LiteralPath $checkpointCsv,$errorCsv,$missingCsv,$runInfoCsv,$OutputXlsx -Force -ErrorAction SilentlyContinue
}
#endregion Load inputs and resume

#region Main
Write-Log -Level INFO -Message "Starting validation. Rows: $(@($rows).Count). IncludeMissingItems: $IncludeMissingItems. FastCountsOnly: $FastCountsOnly."

$rowNumber = 0
foreach ($row in $rows) {
    $rowNumber++
    $key = $row.Wave + '|' + $row.srcUPN + '|' + $row.srcUNCPath
    if ($processedKeys.Contains($key)) {
        Write-Log -Level INFO -Message "Skipping completed row $rowNumber/$(@($rows).Count): $($row.DisplayName)"
        continue
    }

    Write-Log -Level INFO -Message "Processing row $rowNumber/$(@($rows).Count): $($row.DisplayName) <$($row.TargetUPN)>"

    $listName = Get-LibraryNameFromRow -Row $row
    $folderName = Get-FolderNameFromRow -Row $row
    $collectItems = ($IncludeMissingItems -and -not $FastCountsOnly)

    $source = $null
    $dest = $null
    $status = 'Completed'
    $statusDetails = ''

    try {
        $source = Get-SourceInventory -RootPath $row.srcUNCPath -CollectItems:$collectItems
    }
    catch {
        $status = 'SourceScanFailed'
        $statusDetails = $_.Exception.Message
        Add-ErrorRecord -Wave $row.Wave -DisplayName $row.DisplayName -srcUPN $row.srcUPN -TargetUPN $row.TargetUPN -srcUNCPath $row.srcUNCPath -destSiteURL $row.destSiteURL -Stage 'SourceScan' -Message $statusDetails
        $source = [PSCustomObject]@{ Exists=$false; ItemCount=$null; FileCount=$null; FolderCount=$null; TotalSizeBytes=$null; AccessErrors=1; ErrorMessages=@($statusDetails); Items=$null; ItemRows=$null }
    }

    try {
        $dest = Get-OneDriveFolderInventory -Row $row -SiteUrl $row.destSiteURL -ListName $listName -FolderName $folderName -CollectItems:$collectItems
    }
    catch {
        if ($status -eq 'Completed') { $status = 'DestinationScanFailed' } else { $status += ';DestinationScanFailed' }
        $statusDetails = $_.Exception.Message
        Add-ErrorRecord -Wave $row.Wave -DisplayName $row.DisplayName -srcUPN $row.srcUPN -TargetUPN $row.TargetUPN -srcUNCPath $row.srcUNCPath -destSiteURL $row.destSiteURL -Stage 'DestinationScan' -Message $statusDetails
        $dest = [PSCustomObject]@{ Exists=$false; ItemCount=$null; FileCount=$null; FolderCount=$null; TotalSizeBytes=$null; FolderUrl=$null; Items=$null; ErrorMessages=@($statusDetails) }
    }

    $itemDiff   = if ($null -ne $source.ItemCount -and $null -ne $dest.ItemCount) { [int64]$source.ItemCount - [int64]$dest.ItemCount } else { $null }
    $fileDiff   = if ($null -ne $source.FileCount -and $null -ne $dest.FileCount) { [int64]$source.FileCount - [int64]$dest.FileCount } else { $null }
    $folderDiff = if ($null -ne $source.FolderCount -and $null -ne $dest.FolderCount) { [int64]$source.FolderCount - [int64]$dest.FolderCount } else { $null }
    $sizeDiff   = if ($null -ne $source.TotalSizeBytes -and $null -ne $dest.TotalSizeBytes) { [int64]$source.TotalSizeBytes - [int64]$dest.TotalSizeBytes } else { $null }

    if ($status -eq 'Completed') {
        if ($itemDiff -ne 0 -or $fileDiff -ne 0 -or $folderDiff -ne 0 -or $sizeDiff -ne 0) { $status = 'Discrepancy' }
        else { $status = 'Matched' }
    }

    $sourceErrors = if ($source.ErrorMessages) { ($source.ErrorMessages -join ' | ') } else { '' }
    $destErrors = if ($dest.ErrorMessages) { ($dest.ErrorMessages -join ' | ') } else { '' }

    $result = [ordered]@{}
    foreach ($col in $requiredColumns) { $result[$col] = $row.PSObject.Properties[$col].Value }
    $result['NDriveItemCount'] = $source.ItemCount
    $result['OneDriveItemCount'] = $dest.ItemCount
    $result['ItemCountDiscrepancy'] = $itemDiff
    $result['NDriveFolderCount'] = $source.FolderCount
    $result['OneDriveFolderCount'] = $dest.FolderCount
    $result['FolderCountDiscrepancy'] = $folderDiff
    $result['NDriveFileCount'] = $source.FileCount
    $result['OneDriveFileCount'] = $dest.FileCount
    $result['FileCountDiscrepancy'] = $fileDiff
    $result['NDriveTotalSizeBytes'] = $source.TotalSizeBytes
    $result['OneDriveTotalSizeBytes'] = $dest.TotalSizeBytes
    $result['SizeDiscrepancyBytes'] = $sizeDiff
    $result['OneDriveValidationFolderUrl'] = $dest.FolderUrl
    $result['SourcePathExists'] = $source.Exists
    $result['OneDriveFolderExists'] = $dest.Exists
    $result['SourceAccessErrors'] = $source.AccessErrors
    $result['ValidationStatus'] = $status
    $result['ValidationDetails'] = ($statusDetails, $sourceErrors, $destErrors | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }) -join ' | '
    $result['ValidatedAt'] = Get-Date

    [PSCustomObject]$result | Export-Csv -Path $checkpointCsv -NoTypeInformation -Append -Encoding UTF8

    if ($collectItems -and $source.Items -and $dest.Items) {
        foreach ($sourceItem in $source.ItemRows) {
            $keyType = $sourceItem.ObjectType + '|' + (Normalize-RelativePath $sourceItem.RelativePath)
            if (-not $dest.Items.Contains($keyType)) {
                [PSCustomObject]@{
                    Wave                        = $row.Wave
                    DisplayName                 = $row.DisplayName
                    srcUPN                      = $row.srcUPN
                    TargetUPN                   = $row.TargetUPN
                    Email                       = $row.Email
                    srcUNCPath                  = $row.srcUNCPath
                    destSiteURL                 = $row.destSiteURL
                    OneDriveValidationFolderUrl = $dest.FolderUrl
                    ObjectType                  = $sourceItem.ObjectType
                    RelativePath                = $sourceItem.RelativePath
                    SourceFullPath              = $sourceItem.SourceFullPath
                    ExpectedOneDrivePath        = ($dest.FolderUrl.TrimEnd('/') + '/' + (($sourceItem.RelativePath -replace '\\','/').TrimStart('/')))
                    SizeBytes                   = $sourceItem.SizeBytes
                    Reason                      = 'Source item not found in OneDrive target folder inventory'
                } | Export-Csv -Path $missingCsv -NoTypeInformation -Append -Encoding UTF8
            }
        }
    }
}
#endregion Main

#region Export workbook
Write-Log -Level INFO -Message 'Building Excel workbook.'
#$results = @(Import-Csv -Path $checkpointCsv)
if (Test-Path $checkpointCsv) {
    $results = @(Import-Csv -Path $checkpointCsv)
}
else {
    throw "Checkpoint file was never created: $checkpointCsv"
}
$errors = if (Test-Path -LiteralPath $errorCsv) { @(Import-Csv -Path $errorCsv) } else { @() }
$missing = if (Test-Path -LiteralPath $missingCsv) { @(Import-Csv -Path $missingCsv) } else { @() }

function To-Int64Safe { param($Value) $n = 0L; if ([int64]::TryParse([string]$Value, [ref]$n)) { return $n } return 0L }

$totalUsers = @($results).Count
$matchedUsers = @($results | Where-Object { $_.ValidationStatus -eq 'Matched' }).Count
$discrepancyUsers = @($results | Where-Object { $_.ValidationStatus -eq 'Discrepancy' }).Count
$errorUsers = @($results | Where-Object { $_.ValidationStatus -notin @('Matched','Discrepancy') }).Count
$totalNDriveItems = ($results | ForEach-Object { To-Int64Safe $_.NDriveItemCount } | Measure-Object -Sum).Sum
$totalOneDriveItems = ($results | ForEach-Object { To-Int64Safe $_.OneDriveItemCount } | Measure-Object -Sum).Sum
$totalItemDiff = ($results | ForEach-Object { To-Int64Safe $_.ItemCountDiscrepancy } | Measure-Object -Sum).Sum
$totalFolderDiff = ($results | ForEach-Object { To-Int64Safe $_.FolderCountDiscrepancy } | Measure-Object -Sum).Sum
$totalFileDiff = ($results | ForEach-Object { To-Int64Safe $_.FileCountDiscrepancy } | Measure-Object -Sum).Sum
$totalSizeDiff = ($results | ForEach-Object { To-Int64Safe $_.SizeDiscrepancyBytes } | Measure-Object -Sum).Sum

$dashboard = @(
    [PSCustomObject]@{ Metric='Total Users Processed'; Value=$totalUsers; Notes='' },
    [PSCustomObject]@{ Metric='Matched Users'; Value=$matchedUsers; Notes='No item, file, folder, or size discrepancy' },
    [PSCustomObject]@{ Metric='Discrepancy Users'; Value=$discrepancyUsers; Notes='Included in DiscrepancyUsers worksheet based on filter' },
    [PSCustomObject]@{ Metric='Errored Users'; Value=$errorUsers; Notes='Review Errors worksheet' },
    [PSCustomObject]@{ Metric='Total NDrive Items'; Value=$totalNDriveItems; Notes='' },
    [PSCustomObject]@{ Metric='Total OneDrive Items'; Value=$totalOneDriveItems; Notes='Only under target OneDrive folder' },
    [PSCustomObject]@{ Metric='Total Item Count Discrepancy'; Value=$totalItemDiff; Notes='NDriveItemCount - OneDriveItemCount' },
    [PSCustomObject]@{ Metric='Total Folder Count Discrepancy'; Value=$totalFolderDiff; Notes='NDriveFolderCount - OneDriveFolderCount' },
    [PSCustomObject]@{ Metric='Total File Count Discrepancy'; Value=$totalFileDiff; Notes='NDriveFileCount - OneDriveFileCount' },
    [PSCustomObject]@{ Metric='Total Size Discrepancy Bytes'; Value=$totalSizeDiff; Notes='NDriveTotalSizeBytes - OneDriveTotalSizeBytes' },
    [PSCustomObject]@{ Metric='Missing Item Rows'; Value=@($missing).Count; Notes='Only populated when -IncludeMissingItems is used' }
)

$discrepancyRows = switch ($DiscrepancyRerunFilter) {
    'All' { $results }
    'CountDiscrepancy' { @($results | Where-Object { (To-Int64Safe $_.ItemCountDiscrepancy) -ne 0 }) }
    'AnyDiscrepancy' { @($results | Where-Object { (To-Int64Safe $_.ItemCountDiscrepancy) -ne 0 -or (To-Int64Safe $_.FileCountDiscrepancy) -ne 0 -or (To-Int64Safe $_.FolderCountDiscrepancy) -ne 0 -or (To-Int64Safe $_.SizeDiscrepancyBytes) -ne 0 -or $_.ValidationStatus -notin @('Matched') }) }
    'None' { @() }
}

$runInfo = @(
    [PSCustomObject]@{ Name='RunStarted'; Value=$runStarted },
    [PSCustomObject]@{ Name='RunCompleted'; Value=(Get-Date) },
    [PSCustomObject]@{ Name='InputCsv'; Value=(Resolve-Path -LiteralPath $InputCsv).Path },
    [PSCustomObject]@{ Name='OutputXlsx'; Value=$OutputXlsx },
    [PSCustomObject]@{ Name='TargetFolderName'; Value=$TargetFolderName },
    [PSCustomObject]@{ Name='DefaultDocumentLibraryName'; Value=$DefaultDocumentLibraryName },
    [PSCustomObject]@{ Name='IncludeMissingItems'; Value=$IncludeMissingItems },
    [PSCustomObject]@{ Name='FastCountsOnly'; Value=$FastCountsOnly },
    [PSCustomObject]@{ Name='PnPPageSize'; Value=$PnPPageSize },
    [PSCustomObject]@{ Name='DiscrepancyRerunFilter'; Value=$DiscrepancyRerunFilter },
    [PSCustomObject]@{ Name='CheckpointCsv'; Value=$checkpointCsv },
    [PSCustomObject]@{ Name='ErrorCsv'; Value=$errorCsv },
    [PSCustomObject]@{ Name='MissingCsv'; Value=$missingCsv },
    [PSCustomObject]@{ Name='LogFile'; Value=$logFile }
)
$runInfo | Export-Csv -Path $runInfoCsv -NoTypeInformation -Encoding UTF8

$commonExcelParams = @{ Path=$OutputXlsx; AutoSize=$true; FreezeTopRow=$true; BoldTopRow=$true; TableStyle='Medium2' }
$dashboard | Export-Excel @commonExcelParams -WorksheetName 'Dashboard' -ClearSheet
$results | Export-Excel @commonExcelParams -WorksheetName 'ValidationResults' -ClearSheet
$discrepancyRows | Export-Excel @commonExcelParams -WorksheetName 'DiscrepancyUsers' -ClearSheet
$missing | Export-Excel @commonExcelParams -WorksheetName 'MissingItems' -ClearSheet
$errors | Export-Excel @commonExcelParams -WorksheetName 'Errors' -ClearSheet
$runInfo | Export-Excel @commonExcelParams -WorksheetName 'RunInfo' -ClearSheet

# Add charts and basic formatting
<# Dashboard chart creation temporarily disabled
$pkg = Open-ExcelPackage -Path $OutputXlsx
try {
    $wsDash = $pkg.Workbook.Worksheets['Dashboard']
    if ($wsDash) {
        $wsDash.Cells['A1:C1'].Style.Font.Bold = $true
        $wsDash.Cells['A1:C1'].Style.Fill.PatternType = 'Solid'
        $wsDash.Cells['A1:C1'].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::FromArgb(31,78,121))
        $wsDash.Cells['A1:C1'].Style.Font.Color.SetColor([System.Drawing.Color]::White)
        $chart = $wsDash.Drawings.AddChart('UserStatusChart','Pie')
        $chart.Title.Text = 'User Validation Status'
        $statusDataStart = 15
        $wsDash.Cells[$statusDataStart,1].Value = 'Status'
        $wsDash.Cells[$statusDataStart,2].Value = 'Count'
        $wsDash.Cells[$statusDataStart+1,1].Value = 'Matched'
        $wsDash.Cells[$statusDataStart+1,2].Value = $matchedUsers
        $wsDash.Cells[$statusDataStart+2,1].Value = 'Discrepancy'
        $wsDash.Cells[$statusDataStart+2,2].Value = $discrepancyUsers
        $wsDash.Cells[$statusDataStart+3,1].Value = 'Error'
        $wsDash.Cells[$statusDataStart+3,2].Value = $errorUsers
        #$chart.Series.Add($wsDash.Cells[$statusDataStart+1,2,$statusDataStart+3,2],$wsDash.Cells[$statusDataStart+1,1,$statusDataStart+3,1]) | Out-Null
        $chart.SetPosition(1,0,4,0)
        $chart.SetSize(520,320)
    }
}
finally {
    Close-ExcelPackage $pkg
}
#>
Write-Log -Level INFO -Message "Report complete: $OutputXlsx"
Write-Host "`nDone. Report created: $OutputXlsx" -ForegroundColor Green
Write-Host "Checkpoint CSV: $checkpointCsv"
Write-Host "Log: $logFile"
if (Test-Path -LiteralPath $missingCsv) { Write-Host "Missing items CSV: $missingCsv" }
if (Test-Path -LiteralPath $errorCsv) { Write-Host "Errors CSV: $errorCsv" }
#endregion Export workbook
