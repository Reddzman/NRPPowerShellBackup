#Requires -Version 5.1

<#
.SYNOPSIS
    Identifies discovery output reports associated with directories listed in an errors CSV and evaluates
    the related scan log sections for failure indicators.

.DESCRIPTION
    This script reviews discovery scan folders beneath a provided root path. For each CSV in the
    directories, files, and permissions report folders, it reads up to a configurable number of valid
    records and verifies that the records contain one consistent ServerUNC and ShareName combination.

    A report is relevant only when its consistent ServerUNC and ShareName combination matches a row in
    the provided errors CSV. Relevant reports are then evaluated against the scan folder's log file.
    A log section begins with a "Scanning Server:" line and ends immediately before the next such line.
    Configurable failure text, initially "Error - " and "Get-ACL", determines whether a section failed.

    The script is report-only. It never deletes discovery output. Matching reports are written as
    individual rows to a UTF-8 findings CSV. Reports with inconsistent identities or processing problems
    are written to a separate UTF-8 exceptions CSV.

.AUTHENTICATION
    No authentication is required. The account running the script must have read access to the discovery
    root, error CSV, report CSV files, and scan log files, plus write access to the output and log folders.

.PARAMETER DiscoveryRootPath
    Root folder containing one or more scan folders. Each scan folder may contain directories, files,
    permissions, details, and errors subfolders. Only directories, files, and permissions are inspected.

.PARAMETER ErrorCsvPath
    Path to the aggregate error CSV. The CSV must contain serverUNC and sharename columns.

.PARAMETER OutputPath
    Folder where the findings and exceptions CSV files are created.

.PARAMETER LogPath
    Folder where the PowerShell transcript is created.

.PARAMETER ValidRecordsToCheck
    Maximum number of valid records to inspect in each report. A valid record contains nonblank
    ServerUNC and ShareName values. Blank and malformed records do not count toward this limit.

.EXAMPLE
    cd "C:\github\CatPowerShell\Reporting"
    .\Cat.Files-FindIncompleteDiscoveryReports.ps1 `
        -DiscoveryRootPath "..\_Output Files\strongtie" `
        -ErrorCsvPath "..\_Input Files\discovery-errors-all-260922.csv" `
        -ValidRecordsToCheck 10 `
        -OutputPath "..\_Output Files" `
        -LogPath "..\_Log Files"

    Checks only the first valid record in each report.
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$DiscoveryRootPath,

    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$ErrorCsvPath,

    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [string]$OutputPath = "..\_Output Files",

    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [string]$LogPath = "..\_Log Files",

    [Parameter(Mandatory = $false)]
    [ValidateRange(1, 1000)]
    [int]$ValidRecordsToCheck = 10
)

#region 1. Functions

function Get-NormalizedDiscoveryIdentity {
    <#
        Normalization is used repeatedly for the errors CSV, output reports, and log sections. Keeping
        this logic in one function prevents otherwise identical values from comparing differently due
        to case, leading slashes, slash direction, or trailing slashes.
    #>
    [CmdletBinding()]
    param (
        [AllowNull()]
        [AllowEmptyString()]
        [string]$ServerUNC,

        [AllowNull()]
        [AllowEmptyString()]
        [string]$ShareName
    )

    if ([string]::IsNullOrWhiteSpace($ServerUNC) -or [string]::IsNullOrWhiteSpace($ShareName)) {
        return $null
    }

    $normalizedServer = $ServerUNC.Trim().Replace('/', '\').TrimStart('\').TrimEnd('\')
    $normalizedShare = $ShareName.Trim().Replace('/', '\').TrimStart('\').TrimEnd('\')

    if ([string]::IsNullOrWhiteSpace($normalizedServer) -or [string]::IsNullOrWhiteSpace($normalizedShare)) {
        return $null
    }

    [pscustomobject]@{
        ServerUNC  = $normalizedServer
        ShareName  = $normalizedShare
        Key        = ("{0}|{1}" -f $normalizedServer, $normalizedShare)
        UNCPath    = ("\\{0}\{1}" -f $normalizedServer, $normalizedShare)
    }
}

function Get-ReportIdentitySample {
    <#
        Reads only enough of a report to classify its identity. TextFieldParser is used instead of
        splitting lines on commas so quoted commas, escaped quotation marks, and multiline fields are
        handled as CSV data. Blank or malformed records are skipped and do not count as valid records.
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [System.IO.FileInfo]$ReportFile,

        [Parameter(Mandatory = $true)]
        [int]$RecordsToCheck
    )

    $parser = $null
    $physicalRecordsRead = 0
    $validRecordsEvaluated = 0
    $invalidRecordsSkipped = 0
    $identities = New-Object 'System.Collections.Generic.Dictionary[string,object]' ([System.StringComparer]::OrdinalIgnoreCase)

    try {
        $parser = New-Object Microsoft.VisualBasic.FileIO.TextFieldParser($ReportFile.FullName)
        $parser.TextFieldType = [Microsoft.VisualBasic.FileIO.FieldType]::Delimited
        $parser.SetDelimiters(@(','))
        $parser.HasFieldsEnclosedInQuotes = $true
        $parser.TrimWhiteSpace = $false

        # Locate the first parseable, nonblank record and treat it as the CSV header.
        $headerFields = $null
        while (-not $parser.EndOfData -and $null -eq $headerFields) {
            try {
                $candidateHeader = $parser.ReadFields()
                $physicalRecordsRead++
            }
            catch [Microsoft.VisualBasic.FileIO.MalformedLineException] {
                $invalidRecordsSkipped++
                continue
            }

            if ($null -ne $candidateHeader -and ($candidateHeader -join '').Trim().Length -gt 0) {
                $headerFields = $candidateHeader
            }
        }

        if ($null -eq $headerFields) {
            return [pscustomobject]@{
                Status                 = 'NoHeader'
                Message                = 'No parseable, nonblank CSV header was found.'
                PhysicalRecordsRead    = $physicalRecordsRead
                ValidRecordsEvaluated  = $validRecordsEvaluated
                InvalidRecordsSkipped  = $invalidRecordsSkipped
                DistinctIdentities     = @()
                Identity               = $null
            }
        }

        # Remove a UTF-8 BOM if it appears in the first column name and locate required columns without
        # assuming a fixed column order.
        $headerFields[0] = $headerFields[0].TrimStart([char]0xFEFF)
        $serverIndex = -1
        $shareIndex = -1

        for ($columnIndex = 0; $columnIndex -lt $headerFields.Count; $columnIndex++) {
            if ($headerFields[$columnIndex].Trim() -ieq 'ServerUNC') {
                $serverIndex = $columnIndex
            }
            elseif ($headerFields[$columnIndex].Trim() -ieq 'ShareName') {
                $shareIndex = $columnIndex
            }
        }

        if ($serverIndex -lt 0 -or $shareIndex -lt 0) {
            return [pscustomobject]@{
                Status                 = 'MissingRequiredColumns'
                Message                = 'The report header does not contain both ServerUNC and ShareName.'
                PhysicalRecordsRead    = $physicalRecordsRead
                ValidRecordsEvaluated  = $validRecordsEvaluated
                InvalidRecordsSkipped  = $invalidRecordsSkipped
                DistinctIdentities     = @()
                Identity               = $null
            }
        }

        while (-not $parser.EndOfData -and $validRecordsEvaluated -lt $RecordsToCheck) {
            try {
                $fields = $parser.ReadFields()
                $physicalRecordsRead++
            }
            catch [Microsoft.VisualBasic.FileIO.MalformedLineException] {
                $invalidRecordsSkipped++
                continue
            }

            # A short record cannot contain both required indexed fields and is therefore skipped.
            if ($null -eq $fields -or $fields.Count -le [Math]::Max($serverIndex, $shareIndex)) {
                $invalidRecordsSkipped++
                continue
            }

            $identity = Get-NormalizedDiscoveryIdentity -ServerUNC $fields[$serverIndex] -ShareName $fields[$shareIndex]
            if ($null -eq $identity) {
                $invalidRecordsSkipped++
                continue
            }

            $validRecordsEvaluated++
            if (-not $identities.ContainsKey($identity.Key)) {
                $identities.Add($identity.Key, $identity)
            }

            # A second identity is sufficient to classify the report as inconsistent. Continuing to read
            # additional records would not change that decision.
            if ($identities.Count -gt 1) {
                break
            }
        }

        if ($validRecordsEvaluated -eq 0) {
            return [pscustomobject]@{
                Status                 = 'NoValidRecords'
                Message                = 'No record containing both ServerUNC and ShareName was found.'
                PhysicalRecordsRead    = $physicalRecordsRead
                ValidRecordsEvaluated  = $validRecordsEvaluated
                InvalidRecordsSkipped  = $invalidRecordsSkipped
                DistinctIdentities     = @()
                Identity               = $null
            }
        }

        if ($identities.Count -gt 1) {
            return [pscustomobject]@{
                Status                 = 'InconsistentIdentity'
                Message                = 'More than one ServerUNC and ShareName combination was found in the sampled valid records.'
                PhysicalRecordsRead    = $physicalRecordsRead
                ValidRecordsEvaluated  = $validRecordsEvaluated
                InvalidRecordsSkipped  = $invalidRecordsSkipped
                DistinctIdentities     = @($identities.Values)
                Identity               = $null
            }
        }

        return [pscustomobject]@{
            Status                 = 'Consistent'
            Message                = 'The sampled valid records contain one consistent ServerUNC and ShareName combination.'
            PhysicalRecordsRead    = $physicalRecordsRead
            ValidRecordsEvaluated  = $validRecordsEvaluated
            InvalidRecordsSkipped  = $invalidRecordsSkipped
            DistinctIdentities     = @($identities.Values)
            Identity               = @($identities.Values)[0]
        }
    }
    catch {
        return [pscustomobject]@{
            Status                 = 'ReadFailure'
            Message                = $_.Exception.Message
            PhysicalRecordsRead    = $physicalRecordsRead
            ValidRecordsEvaluated  = $validRecordsEvaluated
            InvalidRecordsSkipped  = $invalidRecordsSkipped
            DistinctIdentities     = @($identities.Values)
            Identity               = $null
        }
    }
    finally {
        if ($null -ne $parser) {
            $parser.Close()
            $parser.Dispose()
        }
    }
}

function Get-RelevantLogSections {
    <#
        Parses one potentially large log as a stream. Only sections whose normalized identity belongs
        to a matching report are retained. This limits memory use and ensures unrelated log errors do
        not become findings.
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [System.IO.FileInfo]$LogFile,

        [Parameter(Mandatory = $true)]
        [System.Collections.Generic.HashSet[string]]$TargetKeys,

        [Parameter(Mandatory = $true)]
        [string[]]$FailurePatterns
    )

    $sectionsByKey = New-Object 'System.Collections.Generic.Dictionary[string,object]' ([System.StringComparer]::OrdinalIgnoreCase)
    $reader = $null
    $lineNumber = 0
    $currentSection = $null

    # The share value is captured to the end of the line so spaces and special characters remain intact.
    $scanStartPattern = '^(?<Timestamp>\d{2}-\d{2}-\d{2}\s+\d{2}:\d{2}:\d{2})\s+-\s+Scanning Server:\s*(?<Server>.*?)\s+Share:\s*(?<Share>.+?)\s*$'

    try {
        $reader = New-Object System.IO.StreamReader($LogFile.FullName, $true)

        while (-not $reader.EndOfStream) {
            $line = $reader.ReadLine()
            $lineNumber++
            $match = [regex]::Match($line, $scanStartPattern, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)

            if ($match.Success) {
                # Complete the previous relevant section when the next scan marker is reached.
                if ($null -ne $currentSection) {
                    $currentSection.EndLine = $lineNumber - 1
                    if (-not $sectionsByKey.ContainsKey($currentSection.Key)) {
                        $sectionsByKey.Add($currentSection.Key, (New-Object System.Collections.ArrayList))
                    }
                    [void]$sectionsByKey[$currentSection.Key].Add([pscustomobject]$currentSection)
                    $currentSection = $null
                }

                $identity = Get-NormalizedDiscoveryIdentity -ServerUNC $match.Groups['Server'].Value -ShareName $match.Groups['Share'].Value

                # Ignore non-target sections entirely. Their errors are not relevant to report findings.
                if ($null -ne $identity -and $TargetKeys.Contains($identity.Key)) {
                    $currentSection = [ordered]@{
                        Key                  = $identity.Key
                        ServerUNC            = $identity.ServerUNC
                        ShareName            = $identity.ShareName
                        ScanStart            = $match.Groups['Timestamp'].Value
                        StartLine            = $lineNumber
                        EndLine              = $null
                        FailurePatternsFound = New-Object System.Collections.ArrayList
                    }
                }
                continue
            }

            if ($null -ne $currentSection) {
                foreach ($failurePattern in $FailurePatterns) {
                    if ($line.IndexOf($failurePattern, [System.StringComparison]::OrdinalIgnoreCase) -ge 0 -and
                        -not $currentSection.FailurePatternsFound.Contains($failurePattern)) {
                        [void]$currentSection.FailurePatternsFound.Add($failurePattern)
                    }
                }
            }
        }

        # The final section ends at the end of the file because there is no following scan marker.
        if ($null -ne $currentSection) {
            $currentSection.EndLine = $lineNumber
            if (-not $sectionsByKey.ContainsKey($currentSection.Key)) {
                $sectionsByKey.Add($currentSection.Key, (New-Object System.Collections.ArrayList))
            }
            [void]$sectionsByKey[$currentSection.Key].Add([pscustomobject]$currentSection)
        }

        return [pscustomobject]@{
            Status        = 'Success'
            Message       = 'The log was parsed successfully.'
            SectionsByKey = $sectionsByKey
            LinesRead     = $lineNumber
        }
    }
    catch {
        return [pscustomobject]@{
            Status        = 'ReadFailure'
            Message       = $_.Exception.Message
            SectionsByKey = $sectionsByKey
            LinesRead     = $lineNumber
        }
    }
    finally {
        if ($null -ne $reader) {
            $reader.Close()
            $reader.Dispose()
        }
    }
}

function Write-Utf8Csv {
    <#
        Export-Csv does not create a useful header-only file when the collection is empty. This helper is
        used for both requested CSV outputs and guarantees that each output exists with a stable schema.
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [object[]]$Rows,

        [Parameter(Mandatory = $true)]
        [string[]]$Columns,

        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    if ($Rows.Count -gt 0) {
        $Rows | Select-Object -Property $Columns | Export-Csv -Path $Path -NoTypeInformation -Encoding UTF8
    }
    else {
        # ConvertTo-Csv provides correct CSV escaping for column names. Selecting the header line creates
        # an empty but structurally valid result file.
        $emptyRow = [ordered]@{}
        foreach ($column in $Columns) {
            $emptyRow[$column] = $null
        }
        $headerLine = ([pscustomobject]$emptyRow | ConvertTo-Csv -NoTypeInformation)[0]
        Set-Content -Path $Path -Value $headerLine -Encoding UTF8
    }
}

#endregion 1. Functions

#region 2. Script Start and Configuration

$scriptStart = Get-Date
$runTimestamp = $scriptStart.ToString('yyMMdd-HHmmss')
$transcriptStarted = $false

# These are intentionally simple literal strings. Add another array entry when discovery logs introduce
# another failure message that should cause a scan section to be evaluated as failed.
$scanFailurePatterns = @('Error - ','Get-ACL')
$reportFolderNames = @('directories', 'files', 'permissions')
$findingsColumns = @('ScanFolderName', 'ScanFolderPath', 'ReportType', 'ReportFileName', 'ReportFilePath',
    'ReportFileSizeBytes', 'ReportLastWriteTime', 'RecordsRequested', 'ValidRecordsEvaluated',
    'InvalidRecordsSkipped', 'ErrorServerUNC', 'ErrorShareName', 'ReportServerUNC', 'ReportShareName',
    'MatchedUNCPath', 'LogFilePath', 'MatchingLogSectionCount', 'FailedLogSectionCount',
    'SuccessfulLogSectionCount', 'LogScanStartTimes', 'LogLineRanges', 'FailurePatternsFound',
    'EvaluationResult', 'Confidence', 'ConfidenceReason', 'RecommendedAction', 'Notes'
)

$exceptionsColumns = @(
    'ScanFolderName', 'ScanFolderPath', 'ReportType', 'ReportFileName', 'ReportFilePath',
    'ExceptionType', 'ExceptionMessage', 'PhysicalRecordsRead', 'ValidRecordsEvaluated',
    'InvalidRecordsSkipped', 'DistinctIdentityCount', 'DistinctIdentitiesFound', 'RecommendedAction'
)

#endregion 2. Script Start and Configuration

#region 3. Input Validation and Output Preparation

Write-Host
Write-Host ("{0} - Starting Find Incomplete Discovery Reports" -f (Get-Date -Format 'yy-MM-dd HH:mm:ss')) -ForegroundColor Cyan
Write-Host
Write-Host 'Validating Script Inputs' -ForegroundColor Magenta

if (-not (Test-Path -LiteralPath $DiscoveryRootPath -PathType Container)) {
    throw ("Discovery root path was not found: {0}" -f $DiscoveryRootPath)
}
if (-not (Test-Path -LiteralPath $ErrorCsvPath -PathType Leaf)) {
    throw ("Errors CSV was not found: {0}" -f $ErrorCsvPath)
}

foreach ($requiredDirectory in @($OutputPath, $LogPath)) {
    if (-not (Test-Path -LiteralPath $requiredDirectory -PathType Container)) {
        Write-Host ("   {0} - Creating directory: {1}" -f (Get-Date -Format 'yy-MM-dd HH:mm:ss'), $requiredDirectory) -ForegroundColor DarkGray
        New-Item -Path $requiredDirectory -ItemType Directory -Force | Out-Null
    }
}

$findingsPath = Join-Path -Path $OutputPath -ChildPath ("Cat.Files-CompareDiscoveryIssuesToOutput-Findings-{0}.csv" -f $runTimestamp)
$exceptionsPath = Join-Path -Path $OutputPath -ChildPath ("Cat.Files-CompareDiscoveryIssuesToOutput-Exceptions-{0}.csv" -f $runTimestamp)
$transcriptPath = Join-Path -Path $LogPath -ChildPath ("Cat.Files-CompareDiscoveryIssuesToOutput-{0}.log" -f $runTimestamp)

try {
    Start-Transcript -Path $transcriptPath -Force | Out-Null
    $transcriptStarted = $true
}
catch {
    Write-Warning ("The transcript could not be started. Processing will continue. {0}" -f $_.Exception.Message)
}

Write-Host ("   {0} - Importing errors CSV" -f (Get-Date -Format 'yy-MM-dd HH:mm:ss'))
$errorRows = @(Import-Csv -LiteralPath $ErrorCsvPath)

if ($errorRows.Count -eq 0) {
    throw 'The errors CSV contains no data rows.'
}

$errorPropertyNames = @($errorRows[0].PSObject.Properties.Name)
if (-not ($errorPropertyNames -icontains 'serverUNC') -or -not ($errorPropertyNames -icontains 'sharename')) {
    throw 'The errors CSV must contain serverUNC and sharename columns.'
}

$errorLookup = New-Object 'System.Collections.Generic.Dictionary[string,object]' ([System.StringComparer]::OrdinalIgnoreCase)
$invalidErrorRows = 0

foreach ($errorRow in $errorRows) {
    $identity = Get-NormalizedDiscoveryIdentity -ServerUNC $errorRow.serverUNC -ShareName $errorRow.sharename
    if ($null -eq $identity) {
        $invalidErrorRows++
        continue
    }

    if (-not $errorLookup.ContainsKey($identity.Key)) {
        $errorLookup.Add($identity.Key, [pscustomobject]@{
            ServerUNC = $errorRow.serverUNC.Trim()
            ShareName = $errorRow.sharename.Trim()
            Identity  = $identity
        })
    }
}

if ($errorLookup.Count -eq 0) {
    throw 'The errors CSV contains no valid serverUNC and sharename combinations.'
}

Write-Host
Write-Host 'Script Parameters' -ForegroundColor Cyan
Write-Host ("   Discovery Root Path: {0}" -f $DiscoveryRootPath)
Write-Host ("   Error CSV Path: {0}" -f $ErrorCsvPath)
Write-Host ("   Valid Records to Check: {0}" -f $ValidRecordsToCheck)
Write-Host ("   Failure Patterns: {0}" -f ($scanFailurePatterns -join '; '))
Write-Host
Write-Host 'Output Variables' -ForegroundColor Cyan
Write-Host ("   Findings CSV: {0}" -f $findingsPath)
Write-Host ("   Exceptions CSV: {0}" -f $exceptionsPath)
Write-Host ("   Transcript: {0}" -f $transcriptPath)
Write-Host
Write-Host 'Errors CSV Summary' -ForegroundColor Cyan
Write-Host ("   Source Rows: {0}" -f $errorRows.Count)
Write-Host ("   Unique Valid Identities: {0}" -f $errorLookup.Count)
Write-Host ("   Invalid Rows Skipped: {0}" -f $invalidErrorRows)

#endregion 3. Input Validation and Output Preparation

#region 4. Main Script

$findings = New-Object System.Collections.ArrayList
$exceptions = New-Object System.Collections.ArrayList

$summary = [ordered]@{
    ScanFoldersEvaluated        = 0
    ReportFilesDiscovered       = 0
    ReportsWithConsistentID     = 0
    ReportsWithNoErrorMatch     = 0
    ReportsMatchingErrorCsv     = 0
    ReportsEvaluatedFailed      = 0
    ReportsEvaluatedSuccessful  = 0
    ReportsEvaluatedIndeterminate = 0
    ReportsWithInconsistentID   = 0
    ReportsWithNoValidRecords   = 0
    ReportsWithExceptions       = 0
    LogsParsed                  = 0
}

$scanFolders = @(Get-ChildItem -LiteralPath $DiscoveryRootPath -Directory -ErrorAction Stop | Sort-Object Name)

Write-Host
Write-Host 'Evaluating Discovery Scan Folders' -ForegroundColor Magenta
Write-Host ("   {0} - Scan folders discovered: {1}" -f (Get-Date -Format 'yy-MM-dd HH:mm:ss'), $scanFolders.Count)

foreach ($scanFolder in $scanFolders) {
    $summary.ScanFoldersEvaluated++
    Write-Host
    Write-Host ("   {0} - Scan Folder: {1}" -f (Get-Date -Format 'yy-MM-dd HH:mm:ss'), $scanFolder.Name) -ForegroundColor Cyan

    # First identify the reports that match the errors CSV. The log is intentionally not parsed unless
    # at least one relevant report exists in this scan folder.
    $matchingReports = New-Object System.Collections.ArrayList

    foreach ($reportFolderName in $reportFolderNames) {
        $reportFolderPath = Join-Path -Path $scanFolder.FullName -ChildPath $reportFolderName
        if (-not (Test-Path -LiteralPath $reportFolderPath -PathType Container)) {
            Write-Host ("      {0} folder not found; skipping." -f $reportFolderName) -ForegroundColor DarkGray
            continue
        }

        $reportFiles = @(Get-ChildItem -LiteralPath $reportFolderPath -File -Filter '*.csv' -ErrorAction SilentlyContinue | Sort-Object Name)
        Write-Host ("      {0}: {1} CSV report(s)" -f $reportFolderName, $reportFiles.Count) -ForegroundColor DarkGray

        foreach ($reportFile in $reportFiles) {
            $summary.ReportFilesDiscovered++
            $sample = Get-ReportIdentitySample -ReportFile $reportFile -RecordsToCheck $ValidRecordsToCheck

            if ($sample.Status -ne 'Consistent') {
                $summary.ReportsWithExceptions++
                if ($sample.Status -eq 'InconsistentIdentity') {
                    $summary.ReportsWithInconsistentID++
                }
                elseif ($sample.Status -eq 'NoValidRecords') {
                    $summary.ReportsWithNoValidRecords++
                }

                $distinctIdentityText = @($sample.DistinctIdentities | ForEach-Object { $_.UNCPath }) -join '; '
                [void]$exceptions.Add([pscustomobject]@{
                    ScanFolderName        = $scanFolder.Name
                    ScanFolderPath        = $scanFolder.FullName
                    ReportType           = $reportFolderName
                    ReportFileName       = $reportFile.Name
                    ReportFilePath       = $reportFile.FullName
                    ExceptionType        = $sample.Status
                    ExceptionMessage     = $sample.Message
                    PhysicalRecordsRead  = $sample.PhysicalRecordsRead
                    ValidRecordsEvaluated = $sample.ValidRecordsEvaluated
                    InvalidRecordsSkipped = $sample.InvalidRecordsSkipped
                    DistinctIdentityCount = @($sample.DistinctIdentities).Count
                    DistinctIdentitiesFound = $distinctIdentityText
                    RecommendedAction    = 'Manual review'
                })
                continue
            }

            $summary.ReportsWithConsistentID++
            if (-not $errorLookup.ContainsKey($sample.Identity.Key)) {
                $summary.ReportsWithNoErrorMatch++
                continue
            }

            $summary.ReportsMatchingErrorCsv++
            [void]$matchingReports.Add([pscustomobject]@{
                ReportType = $reportFolderName
                ReportFile = $reportFile
                Sample     = $sample
                ErrorEntry = $errorLookup[$sample.Identity.Key]
            })
        }
    }

    if ($matchingReports.Count -eq 0) {
        Write-Host '      No reports matched the errors CSV. Log parsing was not required.' -ForegroundColor DarkGray
        continue
    }

    Write-Host ("      Matching reports found: {0}" -f $matchingReports.Count) -ForegroundColor Yellow

    # Reports in this scan folder define the only log identities that matter. This prevents unrelated log
    # errors from influencing findings and reduces the amount of section data retained in memory.
    $targetKeys = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($matchingReport in $matchingReports) {
        [void]$targetKeys.Add($matchingReport.Sample.Identity.Key)
    }

    $logFiles = @(Get-ChildItem -LiteralPath $scanFolder.FullName -File -Filter '*.log' -ErrorAction SilentlyContinue | Sort-Object Name)
    $logFile = $null
    $logParse = $null
    $logAvailabilityMessage = $null

    if ($logFiles.Count -eq 1) {
        $logFile = $logFiles[0]
        Write-Host ("      {0} - Parsing Log: {1}" -f (Get-Date -Format 'yy-MM-dd HH:mm:ss'), $logFile.Name)
        $logParse = Get-RelevantLogSections -LogFile $logFile -TargetKeys $targetKeys -FailurePatterns $scanFailurePatterns
        if ($logParse.Status -eq 'Success') {
            $summary.LogsParsed++
        }
        else {
            $logAvailabilityMessage = ("The log could not be parsed: {0}" -f $logParse.Message)
        }
    }
    elseif ($logFiles.Count -eq 0) {
        $logAvailabilityMessage = 'No .log file was found directly in the scan folder.'
    }
    else {
        $logAvailabilityMessage = ("Multiple .log files were found directly in the scan folder: {0}" -f (($logFiles.Name) -join '; '))
    }

    foreach ($matchingReport in $matchingReports) {
        $identity = $matchingReport.Sample.Identity
        $errorEntry = $matchingReport.ErrorEntry
        $matchingSections = @()

        if ($null -ne $logParse -and $logParse.Status -eq 'Success' -and $logParse.SectionsByKey.ContainsKey($identity.Key)) {
            $matchingSections = @($logParse.SectionsByKey[$identity.Key])
        }

        $failedSections = @($matchingSections | Where-Object { $_.FailurePatternsFound.Count -gt 0 })
        $successfulSections = @($matchingSections | Where-Object { $_.FailurePatternsFound.Count -eq 0 })
        $allPatterns = @($failedSections | ForEach-Object { $_.FailurePatternsFound } | Sort-Object -Unique)

        $evaluationResult = 'Indeterminate'
        $confidence = 'Low'
        $confidenceReason = ''
        $recommendedAction = 'Manual review'
        $notes = ''

        if ($null -ne $logAvailabilityMessage) {
            $confidenceReason = 'The report identity matches the errors CSV, but the scan log is unavailable or ambiguous.'
            $notes = $logAvailabilityMessage
        }
        elseif ($null -ne $logParse -and $logParse.Status -ne 'Success') {
            $confidenceReason = 'The report identity matches the errors CSV, but the scan log could not be parsed.'
            $notes = $logParse.Message
        }
        elseif ($matchingSections.Count -eq 0) {
            $confidenceReason = 'The report identity matches the errors CSV, but no matching scan section was found in the log.'
            $notes = 'No matching log section was found.'
        }
        elseif ($failedSections.Count -eq $matchingSections.Count) {
            $evaluationResult = 'Failed'
            $confidence = 'High'
            $confidenceReason = 'The sampled report identity is consistent, matches the errors CSV, and every matching log section contains a configured failure pattern.'
            $recommendedAction = 'Review for deletion'
        }
        elseif ($successfulSections.Count -eq $matchingSections.Count) {
            $evaluationResult = 'Successful'
            $confidence = 'High'
            $confidenceReason = 'The sampled report identity is consistent, matches the errors CSV, and no matching log section contains a configured failure pattern.'
            $recommendedAction = 'Retain'
        }
        else {
            $confidenceReason = 'Matching log sections have conflicting successful and failed evaluations.'
            $notes = 'The report cannot be reliably associated with one of the conflicting log sections.'
        }

        switch ($evaluationResult) {
            'Failed' { $summary.ReportsEvaluatedFailed++ }
            'Successful' { $summary.ReportsEvaluatedSuccessful++ }
            default { $summary.ReportsEvaluatedIndeterminate++ }
        }

        $reportFile = $matchingReport.ReportFile
        [void]$findings.Add([pscustomobject]@{
            ScanFolderName          = $scanFolder.Name
            ScanFolderPath          = $scanFolder.FullName
            ReportType             = $matchingReport.ReportType
            ReportFileName         = $reportFile.Name
            ReportFilePath         = $reportFile.FullName
            ReportFileSizeBytes    = $reportFile.Length
            ReportLastWriteTime    = $reportFile.LastWriteTime.ToString('yyyy-MM-dd HH:mm:ss')
            RecordsRequested       = $ValidRecordsToCheck
            ValidRecordsEvaluated  = $matchingReport.Sample.ValidRecordsEvaluated
            InvalidRecordsSkipped  = $matchingReport.Sample.InvalidRecordsSkipped
            ErrorServerUNC         = $errorEntry.ServerUNC
            ErrorShareName         = $errorEntry.ShareName
            ReportServerUNC        = $identity.ServerUNC
            ReportShareName        = $identity.ShareName
            MatchedUNCPath         = $identity.UNCPath
            LogFilePath            = if ($null -ne $logFile) { $logFile.FullName } else { $null }
            MatchingLogSectionCount = $matchingSections.Count
            FailedLogSectionCount  = $failedSections.Count
            SuccessfulLogSectionCount = $successfulSections.Count
            LogScanStartTimes      = (@($matchingSections | ForEach-Object { $_.ScanStart }) -join '; ')
            LogLineRanges          = (@($matchingSections | ForEach-Object { "{0}-{1}" -f $_.StartLine, $_.EndLine }) -join '; ')
            FailurePatternsFound   = ($allPatterns -join '; ')
            EvaluationResult       = $evaluationResult
            Confidence             = $confidence
            ConfidenceReason       = $confidenceReason
            RecommendedAction      = $recommendedAction
            Notes                  = $notes
        })
    }
}

#endregion 4. Main Script

#region 5. Export Results and Complete Script

Write-Host
Write-Host ("{0} - Exporting Results" -f (Get-Date -Format 'yy-MM-dd HH:mm:ss')) -ForegroundColor Cyan
Write-Utf8Csv -Rows @($findings) -Columns $findingsColumns -Path $findingsPath
Write-Utf8Csv -Rows @($exceptions) -Columns $exceptionsColumns -Path $exceptionsPath

$scriptFinish = Get-Date
$elapsed = $scriptFinish - $scriptStart

Write-Host
Write-Host 'Processing Summary' -ForegroundColor Magenta
Write-Host ("   Scan Folders Evaluated: {0}" -f $summary.ScanFoldersEvaluated)
Write-Host ("   Report Files Discovered: {0}" -f $summary.ReportFilesDiscovered)
Write-Host ("   Reports with Consistent Identity: {0}" -f $summary.ReportsWithConsistentID)
Write-Host ("   Reports with No Errors CSV Match: {0}" -f $summary.ReportsWithNoErrorMatch)
Write-Host ("   Reports Matching Errors CSV: {0}" -f $summary.ReportsMatchingErrorCsv)
Write-Host ("   Reports Evaluated as Failed: {0}" -f $summary.ReportsEvaluatedFailed) -ForegroundColor Red
Write-Host ("   Reports Evaluated as Successful: {0}" -f $summary.ReportsEvaluatedSuccessful) -ForegroundColor Green
Write-Host ("   Reports Evaluated as Indeterminate: {0}" -f $summary.ReportsEvaluatedIndeterminate) -ForegroundColor Yellow
Write-Host ("   Reports with Inconsistent Identity: {0}" -f $summary.ReportsWithInconsistentID)
Write-Host ("   Reports with No Valid Records: {0}" -f $summary.ReportsWithNoValidRecords)
Write-Host ("   Reports with Processing Exceptions: {0}" -f $summary.ReportsWithExceptions)
Write-Host ("   Logs Parsed: {0}" -f $summary.LogsParsed)
Write-Host
Write-Host 'Output Files' -ForegroundColor Cyan
Write-Host ("   Findings CSV: {0}" -f $findingsPath)
Write-Host ("   Exceptions CSV: {0}" -f $exceptionsPath)
Write-Host ("   Transcript: {0}" -f $transcriptPath)
Write-Host
Write-Host ("{0} - Script Complete" -f (Get-Date -Format 'yy-MM-dd HH:mm:ss')) -ForegroundColor Cyan
Write-Host ("   Elapsed Time: {0:hh\:mm\:ss}" -f $elapsed)
Write-Host

if ($transcriptStarted) {
    Stop-Transcript | Out-Null
}

#endregion 5. Export Results and Complete Script