#Requires -Version 7.0
#Requires -PSEdition Core

<#
.SYNOPSIS
    Moves reviewed, high-confidence failed discovery report files to the Windows Recycle Bin.

.DESCRIPTION
    This script reads a reviewed findings CSV created by the discovery report comparison script. The
    reviewed CSV should be copied to the standard "_Input Files" directory before this script is run.

    A row is eligible only when all of the following conditions are true:

        * EvaluationResult equals Failed.
        * Confidence equals High.
        * ReportType equals directories, files, or permissions.
        * ReportFilePath contains a valid path beneath the approved output root.

    Duplicate ReportFilePath values are processed only once. Before any file is moved, the script displays
    a candidate summary and requires an explicit Yes or No response. Eligible files are sent to the Windows
    Recycle Bin. The script does not use Remove-Item and does not intentionally perform permanent deletion.

    Every input row is represented in the results CSV. Ineligible rows, duplicate rows, missing files,
    cancelled candidates, successful Recycle Bin operations, and errors are recorded for audit purposes.

.AUTHENTICATION
    No service authentication is required. The interactive account running PowerShell must have:

        * Read access to the reviewed findings CSV.
        * Read and delete access to eligible report files.
        * Write access to the standard "_Output Files" and "_Log Files" directories.

.PARAMETER InputFile
    Full path to the human-reviewed findings CSV copied to the standard "_Input Files" directory. The CSV
    must contain ReportFilePath, ReportFileName, ReportType, EvaluationResult, and Confidence columns.

.PARAMETER ApprovedOutputRoot
    Root directory beneath which report files are allowed to be moved to the Recycle Bin. Files outside
    this directory are skipped. The default is E:\Quisitive\PowerShell\_Output Files.

.PARAMETER OutputPath
    Directory where the UTF-8 cleanup results CSV is created. The default is the standard
    E:\Quisitive\PowerShell\_Output Files directory.

.PARAMETER LogPath
    Directory where the PowerShell transcript is created. The default is the standard
    E:\Quisitive\PowerShell\_Log Files directory.

.EXAMPLE
    .\Cat.Files-DeleteDiscoveryIssueReports.ps1 `
        -InputFile "..\_Input Files\FindIncompleteDiscoveryReports-260922-101123.csv"

    Reviews the candidate summary, prompts for Yes or No, and sends approved files to the Windows Recycle
    Bin when the user enters Yes.

.EXAMPLE
    .\Cat.Files-DeleteDiscoveryIssueReports.ps1 `
        -InputFile "..\_Input Files\ReviewedDiscoveryFindings.csv" `
        -ApprovedOutputRoot "..\_Output Files\strongtie" `
        -OutputPath "..\_Output Files" `
        -LogPath "..\_Log Files"

    Restricts eligible report files to the specified strongtie discovery output root.
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$InputFile,

    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [string]$ApprovedOutputRoot = "..\_Output Files",

    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [string]$OutputPath = "..\_Output Files",

    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [string]$LogPath = "..\_Log Files"
)

#region 1. Functions

function Test-ApprovedDeletionCandidate {
    <#
        The same eligibility checks are needed during initial candidate selection and immediately before
        the Recycle Bin operation. Centralizing those checks prevents the two stages from applying
        different safety rules.
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [psobject]$Row,

        [Parameter(Mandatory = $true)]
        [string]$ApprovedRoot
    )

    if ([string]::IsNullOrWhiteSpace([string]$Row.ReportFilePath)) {
        return [pscustomobject]@{
            Eligible       = $false
            Reason         = 'Missing ReportFilePath'
            NormalizedPath = $null
        }
    }

    if ([string]$Row.EvaluationResult -ine 'Failed') {
        return [pscustomobject]@{
            Eligible       = $false
            Reason         = 'EvaluationResult is not Failed'
            NormalizedPath = $null
        }
    }

    if ([string]$Row.Confidence -ine 'High') {
        return [pscustomobject]@{
            Eligible       = $false
            Reason         = 'Confidence is not High'
            NormalizedPath = $null
        }
    }

    $approvedReportTypes = @('directories', 'files', 'permissions')
    if ([string]$Row.ReportType -inotin $approvedReportTypes) {
        return [pscustomobject]@{
            Eligible       = $false
            Reason         = 'ReportType is not directories, files, or permissions'
            NormalizedPath = $null
        }
    }

    try {
        $normalizedPath = [System.IO.Path]::GetFullPath([string]$Row.ReportFilePath)
        $normalizedRoot = [System.IO.Path]::GetFullPath($ApprovedRoot)

        # Adding the directory separator prevents a similarly named sibling path, such as
        # "_Output Files-Old", from passing a simple StartsWith comparison for "_Output Files".
        $rootWithSeparator = $normalizedRoot.TrimEnd(
            [System.IO.Path]::DirectorySeparatorChar,
            [System.IO.Path]::AltDirectorySeparatorChar
        ) + [System.IO.Path]::DirectorySeparatorChar

        if (-not $normalizedPath.StartsWith($rootWithSeparator, [System.StringComparison]::OrdinalIgnoreCase)) {
            return [pscustomobject]@{
                Eligible       = $false
                Reason         = 'ReportFilePath is outside the approved output root'
                NormalizedPath = $normalizedPath
            }
        }
    }
    catch {
        return [pscustomobject]@{
            Eligible       = $false
            Reason         = ("ReportFilePath is invalid: {0}" -f $_.Exception.Message)
            NormalizedPath = $null
        }
    }

    return [pscustomobject]@{
        Eligible       = $true
        Reason         = 'Approved'
        NormalizedPath = $normalizedPath
    }
}

function Move-FileToRecycleBin {
    <#
        Microsoft.VisualBasic.FileIO provides the Windows Recycle Bin behavior required by this script.
        The assembly is loaded once during script initialization. This function exists because the same
        strongly typed call is made for every eligible report file.
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$Path
    )

    [Microsoft.VisualBasic.FileIO.FileSystem]::DeleteFile(
        $Path,
        [Microsoft.VisualBasic.FileIO.UIOption]::OnlyErrorDialogs,
        [Microsoft.VisualBasic.FileIO.RecycleOption]::SendToRecycleBin
    )
}

function Write-Utf8Csv {
    <#
        Writes a UTF-8 CSV with a stable column order. AllowEmptyCollection prevents an empty results
        collection from causing parameter binding errors. The function also creates a header-only CSV if
        there are no result rows.
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $false)]
        [AllowEmptyCollection()]
        [object[]]$Rows = @(),

        [Parameter(Mandatory = $true)]
        [string[]]$Columns,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$Path
    )

    if ($null -ne $Rows -and $Rows.Count -gt 0) {
        $Rows |
            Select-Object -Property $Columns |
            Export-Csv -LiteralPath $Path -NoTypeInformation -Encoding utf8
    }
    else {
        $emptyRow = [ordered]@{}
        foreach ($column in $Columns) {
            $emptyRow[$column] = $null
        }

        $headerLine = ([pscustomobject]$emptyRow | ConvertTo-Csv -NoTypeInformation)[0]
        Set-Content -LiteralPath $Path -Value $headerLine -Encoding utf8
    }
}

#endregion 1. Functions

#region 2. Script Start and Variables

$scriptStart = Get-Date
$runTimestamp = $scriptStart.ToString('yyMMdd-HHmmss')
$transcriptStarted = $false
$operationCancelled = $false

$resultsPath = Join-Path -Path $OutputPath -ChildPath ("DeleteDiscoveryIssueReports-Results-{0}.csv" -f $runTimestamp)
$transcriptPath = Join-Path -Path $LogPath -ChildPath ("DeleteDiscoveryIssueReports-{0}.log" -f $runTimestamp)

$requiredColumns = @(
    'ReportFilePath',
    'ReportFileName',
    'ReportType',
    'EvaluationResult',
    'Confidence'
)

$resultsColumns = @(
    'InputRowNumber',
    'ReportFileName',
    'ReportFilePath',
    'NormalizedReportFilePath',
    'ReportType',
    'EvaluationResult',
    'Confidence',
    'DeletionEligible',
    'FileExisted',
    'FileSizeBytes',
    'FileLastWriteTime',
    'DeletionStatus',
    'ProcessedDateTime',
    'Notes'
)

$results = New-Object System.Collections.ArrayList
$candidates = New-Object System.Collections.ArrayList

$summary = [ordered]@{
    InputRowsReviewed        = 0
    EligibleRows             = 0
    UniqueCandidateFiles     = 0
    DuplicateRowsSkipped     = 0
    IneligibleRowsSkipped    = 0
    ExistingCandidateFiles   = 0
    MissingCandidateFiles    = 0
    FilesMovedToRecycleBin   = 0
    RecycleBinErrors         = 0
    CandidatesCancelled      = 0
}

Write-Host
Write-Host ("{0} - Starting Discovery Report Cleanup" -f (Get-Date -Format 'yy-MM-dd HH:mm:ss')) -ForegroundColor Cyan

#endregion 2. Script Start and Variables

#region 3. Input Validation and Output Preparation

try {
    Write-Host
    Write-Host 'Validating Script Inputs' -ForegroundColor Magenta

    # The Recycle Bin implementation is Windows-specific. Stop before reading or acting on the review
    # report if the script is launched from a non-Windows PowerShell Core environment.
    if (-not $IsWindows) {
        throw 'This script requires PowerShell Core running interactively on Windows.'
    }

    if (-not (Test-Path -LiteralPath $InputFile -PathType Leaf)) {
        throw ("Input findings CSV was not found: {0}" -f $InputFile)
    }

    if ([System.IO.Path]::GetExtension($InputFile) -ine '.csv') {
        throw ("Input file must be a CSV file: {0}" -f $InputFile)
    }

    if (-not (Test-Path -LiteralPath $ApprovedOutputRoot -PathType Container)) {
        throw ("Approved output root was not found: {0}" -f $ApprovedOutputRoot)
    }

    # Follow the standardized directory pattern. Missing output or log directories are created before
    # transcript and CSV creation so all operational artifacts use the expected locations.
    foreach ($requiredDirectory in @($OutputPath, $LogPath)) {
        if (-not (Test-Path -LiteralPath $requiredDirectory -PathType Container)) {
            Write-Host ("   {0} - Creating directory: {1}" -f (Get-Date -Format 'yy-MM-dd HH:mm:ss'), $requiredDirectory) -ForegroundColor DarkGray
            New-Item -Path $requiredDirectory -ItemType Directory -Force -ErrorAction Stop | Out-Null
        }
    }

    try {
        Start-Transcript -Path $transcriptPath -Force -ErrorAction Stop | Out-Null
        $transcriptStarted = $true
    }
    catch {
        Write-Warning ("The transcript could not be started. Processing will continue. {0}" -f $_.Exception.Message)
    }

    # Load the assembly once rather than once per file. If it is unavailable, the script stops before the
    # confirmation prompt because no approved Recycle Bin implementation would be available.
    Add-Type -AssemblyName Microsoft.VisualBasic -ErrorAction Stop

    Write-Host
    Write-Host 'Script Parameters' -ForegroundColor Cyan
    Write-Host ("   Input File: {0}" -f $InputFile)
    Write-Host ("   Approved Output Root: {0}" -f $ApprovedOutputRoot)
    Write-Host
    Write-Host 'Output Variables' -ForegroundColor Cyan
    Write-Host ("   Results CSV: {0}" -f $resultsPath)
    Write-Host ("   Transcript: {0}" -f $transcriptPath)
}
catch {
    Write-Host
    Write-Host ("{0} - Input Validation Failed" -f (Get-Date -Format 'yy-MM-dd HH:mm:ss')) -ForegroundColor Red
    Write-Host ("   {0}" -f $_.Exception.Message) -ForegroundColor Red

    if ($transcriptStarted) {
        Stop-Transcript | Out-Null
    }

    throw
}

#endregion 3. Input Validation and Output Preparation

#region 4. Import and Validate Reviewed Findings CSV

try {
    Write-Host
    Write-Host ("{0} - Importing Reviewed Findings CSV" -f (Get-Date -Format 'yy-MM-dd HH:mm:ss')) -ForegroundColor Cyan

    $reportRows = @(Import-Csv -LiteralPath $InputFile -ErrorAction Stop)
    if ($reportRows.Count -eq 0) {
        throw 'The reviewed findings CSV contains no data rows.'
    }

    $sourceColumns = @($reportRows[0].PSObject.Properties.Name)
    $missingColumns = @($requiredColumns | Where-Object { $sourceColumns -inotcontains $_ })

    if ($missingColumns.Count -gt 0) {
        throw ("The reviewed findings CSV is missing required columns: {0}" -f ($missingColumns -join ', '))
    }

    $summary.InputRowsReviewed = $reportRows.Count
    Write-Host ("   {0} - Reviewed rows imported: {1}" -f (Get-Date -Format 'yy-MM-dd HH:mm:ss'), $reportRows.Count)
}
catch {
    Write-Host ("   {0} - Findings CSV validation failed: {1}" -f (Get-Date -Format 'yy-MM-dd HH:mm:ss'), $_.Exception.Message) -ForegroundColor Red

    if ($transcriptStarted) {
        Stop-Transcript | Out-Null
    }

    throw
}

#endregion 4. Import and Validate Reviewed Findings CSV

#region 5. Evaluate and Deduplicate Candidates

Write-Host
Write-Host ("{0} - Evaluating Candidate Files" -f (Get-Date -Format 'yy-MM-dd HH:mm:ss')) -ForegroundColor Cyan

# A case-insensitive dictionary ensures paths that differ only by character casing are processed once.
$uniquePathLookup = New-Object 'System.Collections.Generic.Dictionary[string,object]' ([System.StringComparer]::OrdinalIgnoreCase)
$inputRowNumber = 0

foreach ($row in $reportRows) {
    $inputRowNumber++
    $validation = Test-ApprovedDeletionCandidate -Row $row -ApprovedRoot $ApprovedOutputRoot

    if (-not $validation.Eligible) {
        $summary.IneligibleRowsSkipped++
        [void]$results.Add([pscustomobject]@{
            InputRowNumber           = $inputRowNumber
            ReportFileName          = $row.ReportFileName
            ReportFilePath          = $row.ReportFilePath
            NormalizedReportFilePath = $validation.NormalizedPath
            ReportType              = $row.ReportType
            EvaluationResult        = $row.EvaluationResult
            Confidence              = $row.Confidence
            DeletionEligible        = $false
            FileExisted             = $null
            FileSizeBytes           = $null
            FileLastWriteTime       = $null
            DeletionStatus          = 'Skipped - Ineligible'
            ProcessedDateTime       = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
            Notes                    = $validation.Reason
        })
        continue
    }

    $summary.EligibleRows++

    if ($uniquePathLookup.ContainsKey($validation.NormalizedPath)) {
        $summary.DuplicateRowsSkipped++
        [void]$results.Add([pscustomobject]@{
            InputRowNumber           = $inputRowNumber
            ReportFileName          = $row.ReportFileName
            ReportFilePath          = $row.ReportFilePath
            NormalizedReportFilePath = $validation.NormalizedPath
            ReportType              = $row.ReportType
            EvaluationResult        = $row.EvaluationResult
            Confidence              = $row.Confidence
            DeletionEligible        = $false
            FileExisted             = $null
            FileSizeBytes           = $null
            FileLastWriteTime       = $null
            DeletionStatus          = 'Skipped - Duplicate Path'
            ProcessedDateTime       = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
            Notes                    = 'Another eligible row already references this ReportFilePath.'
        })
        continue
    }

    $candidate = [pscustomobject]@{
        InputRowNumber = $inputRowNumber
        SourceRow      = $row
        NormalizedPath = $validation.NormalizedPath
    }

    $uniquePathLookup.Add($validation.NormalizedPath, $candidate)
    [void]$candidates.Add($candidate)
}

$summary.UniqueCandidateFiles = $candidates.Count

# File existence is summarized before confirmation so the user knows how many files can actually be acted
# upon. Existence is checked again immediately before each Recycle Bin operation.
foreach ($candidate in $candidates) {
    if (Test-Path -LiteralPath $candidate.NormalizedPath -PathType Leaf) {
        $summary.ExistingCandidateFiles++
    }
    else {
        $summary.MissingCandidateFiles++
    }
}

#endregion 5. Evaluate and Deduplicate Candidates

#region 6. User Confirmation

Write-Host
Write-Host 'Delete Candidate Summary' -ForegroundColor Magenta
Write-Host ("   Reviewed Report Rows: {0}" -f $summary.InputRowsReviewed)
Write-Host ("   Eligible Failed/High Rows: {0}" -f $summary.EligibleRows)
Write-Host ("   Unique Candidate Files: {0}" -f $summary.UniqueCandidateFiles)
Write-Host ("   Duplicate Eligible Rows Skipped: {0}" -f $summary.DuplicateRowsSkipped)
Write-Host ("   Ineligible Rows Skipped: {0}" -f $summary.IneligibleRowsSkipped)
Write-Host ("   Existing Candidate Files: {0}" -f $summary.ExistingCandidateFiles)
Write-Host ("   Missing Candidate Files: {0}" -f $summary.MissingCandidateFiles)
Write-Host
Write-Host 'WARNING:' -ForegroundColor Red
Write-Host '   Eligible files will be moved to the Windows Recycle Bin.' -ForegroundColor Red
Write-Host '   Only enter Yes after reviewing the candidate summary.' -ForegroundColor Red
Write-Host

if ($candidates.Count -eq 0) {
    Write-Host 'No eligible unique candidate files were found. No confirmation is required.' -ForegroundColor Yellow
    $operationCancelled = $true
}
else {
    do {
        $proceed = Read-Host 'Do you want to continue? (Yes or No)'

        if ($proceed -inotmatch '^(Yes|No)$') {
            Write-Host '   Enter Yes or No.' -ForegroundColor Yellow
        }
    }
    until ($proceed -imatch '^(Yes|No)$')

    if ($proceed -ine 'Yes') {
        $operationCancelled = $true
        Write-Host
        Write-Host ("{0} - Operation Cancelled by User" -f (Get-Date -Format 'yy-MM-dd HH:mm:ss')) -ForegroundColor Yellow
    }
}

#endregion 6. User Confirmation

#region 7. Recycle Bin Processing

if ($operationCancelled) {
    # Preserve an audit row for each candidate that was eligible but not processed because the user
    # cancelled or no actionable candidates existed.
    foreach ($candidate in $candidates) {
        $summary.CandidatesCancelled++
        $row = $candidate.SourceRow
        $fileExists = Test-Path -LiteralPath $candidate.NormalizedPath -PathType Leaf

        [void]$results.Add([pscustomobject]@{
            InputRowNumber           = $candidate.InputRowNumber
            ReportFileName          = $row.ReportFileName
            ReportFilePath          = $row.ReportFilePath
            NormalizedReportFilePath = $candidate.NormalizedPath
            ReportType              = $row.ReportType
            EvaluationResult        = $row.EvaluationResult
            Confidence              = $row.Confidence
            DeletionEligible        = $true
            FileExisted             = $fileExists
            FileSizeBytes           = $null
            FileLastWriteTime       = $null
            DeletionStatus          = 'Cancelled - Not Processed'
            ProcessedDateTime       = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
            Notes                    = 'The user did not approve Recycle Bin processing.'
        })
    }
}
else {
    Write-Host
    Write-Host ("{0} - Moving Approved Files to the Windows Recycle Bin" -f (Get-Date -Format 'yy-MM-dd HH:mm:ss')) -ForegroundColor Cyan

    foreach ($candidate in $candidates) {
        $row = $candidate.SourceRow

        # Reapply all eligibility rules immediately before processing. This protects against in-memory
        # changes and keeps the final operation governed by the same rules used during candidate selection.
        $finalValidation = Test-ApprovedDeletionCandidate -Row $row -ApprovedRoot $ApprovedOutputRoot
        if (-not $finalValidation.Eligible) {
            [void]$results.Add([pscustomobject]@{
                InputRowNumber           = $candidate.InputRowNumber
                ReportFileName          = $row.ReportFileName
                ReportFilePath          = $row.ReportFilePath
                NormalizedReportFilePath = $finalValidation.NormalizedPath
                ReportType              = $row.ReportType
                EvaluationResult        = $row.EvaluationResult
                Confidence              = $row.Confidence
                DeletionEligible        = $false
                FileExisted             = $null
                FileSizeBytes           = $null
                FileLastWriteTime       = $null
                DeletionStatus          = 'Skipped - Final Validation Failed'
                ProcessedDateTime       = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
                Notes                    = $finalValidation.Reason
            })
            continue
        }

        if (-not (Test-Path -LiteralPath $finalValidation.NormalizedPath -PathType Leaf)) {
            [void]$results.Add([pscustomobject]@{
                InputRowNumber           = $candidate.InputRowNumber
                ReportFileName          = $row.ReportFileName
                ReportFilePath          = $row.ReportFilePath
                NormalizedReportFilePath = $finalValidation.NormalizedPath
                ReportType              = $row.ReportType
                EvaluationResult        = $row.EvaluationResult
                Confidence              = $row.Confidence
                DeletionEligible        = $true
                FileExisted             = $false
                FileSizeBytes           = $null
                FileLastWriteTime       = $null
                DeletionStatus          = 'Skipped - File Not Found'
                ProcessedDateTime       = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
                Notes                    = 'The file did not exist when processing was attempted.'
            })
            continue
        }

        $fileSizeBytes = $null
        $fileLastWriteTime = $null

        try {
            $file = Get-Item -LiteralPath $finalValidation.NormalizedPath -Force -ErrorAction Stop
            $fileSizeBytes = $file.Length
            $fileLastWriteTime = $file.LastWriteTime.ToString('yyyy-MM-dd HH:mm:ss')

            Write-Host ("   {0} - Moving to Recycle Bin: {1}" -f (Get-Date -Format 'yy-MM-dd HH:mm:ss'), $file.Name)
            Move-FileToRecycleBin -Path $file.FullName

            # The API does not return an operation result. Confirm that the source path no longer exists
            # before recording the action as successful.
            if (Test-Path -LiteralPath $file.FullName) {
                throw 'The Recycle Bin operation returned, but the source file still exists.'
            }

            $summary.FilesMovedToRecycleBin++
            [void]$results.Add([pscustomobject]@{
                InputRowNumber           = $candidate.InputRowNumber
                ReportFileName          = $row.ReportFileName
                ReportFilePath          = $row.ReportFilePath
                NormalizedReportFilePath = $finalValidation.NormalizedPath
                ReportType              = $row.ReportType
                EvaluationResult        = $row.EvaluationResult
                Confidence              = $row.Confidence
                DeletionEligible        = $true
                FileExisted             = $true
                FileSizeBytes           = $fileSizeBytes
                FileLastWriteTime       = $fileLastWriteTime
                DeletionStatus          = 'Moved to Recycle Bin'
                ProcessedDateTime       = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
                Notes                    = ''
            })

            Write-Host '      Recycle Bin operation completed.' -ForegroundColor Green
        }
        catch {
            $summary.RecycleBinErrors++
            [void]$results.Add([pscustomobject]@{
                InputRowNumber           = $candidate.InputRowNumber
                ReportFileName          = $row.ReportFileName
                ReportFilePath          = $row.ReportFilePath
                NormalizedReportFilePath = $finalValidation.NormalizedPath
                ReportType              = $row.ReportType
                EvaluationResult        = $row.EvaluationResult
                Confidence              = $row.Confidence
                DeletionEligible        = $true
                FileExisted             = $true
                FileSizeBytes           = if ($null -ne $fileSizeBytes) { $fileSizeBytes } else { $null }
                FileLastWriteTime       = if ($null -ne $fileLastWriteTime) { $fileLastWriteTime } else { $null }
                DeletionStatus          = 'Recycle Bin Error'
                ProcessedDateTime       = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
                Notes                    = $_.Exception.Message
            })

            Write-Host ("      Recycle Bin error: {0}" -f $_.Exception.Message) -ForegroundColor Red
        }
    }
}

#endregion 7. Recycle Bin Processing

#region 8. Export Results and Complete Script

Write-Host
Write-Host ("{0} - Exporting Cleanup Results" -f (Get-Date -Format 'yy-MM-dd HH:mm:ss')) -ForegroundColor Cyan

try {
    Write-Utf8Csv -Rows @($results) -Columns $resultsColumns -Path $resultsPath
}
catch {
    Write-Host ("   Results CSV export failed: {0}" -f $_.Exception.Message) -ForegroundColor Red

    if ($transcriptStarted) {
        Stop-Transcript | Out-Null
    }

    throw
}

$scriptFinish = Get-Date
$elapsed = $scriptFinish - $scriptStart

Write-Host
Write-Host 'Processing Summary' -ForegroundColor Magenta
Write-Host ("   Input Rows Reviewed: {0}" -f $summary.InputRowsReviewed)
Write-Host ("   Eligible Failed/High Rows: {0}" -f $summary.EligibleRows)
Write-Host ("   Unique Candidate Files: {0}" -f $summary.UniqueCandidateFiles)
Write-Host ("   Duplicate Rows Skipped: {0}" -f $summary.DuplicateRowsSkipped)
Write-Host ("   Ineligible Rows Skipped: {0}" -f $summary.IneligibleRowsSkipped)
Write-Host ("   Existing Candidate Files: {0}" -f $summary.ExistingCandidateFiles)
Write-Host ("   Missing Candidate Files: {0}" -f $summary.MissingCandidateFiles)
Write-Host ("   Files Moved to Recycle Bin: {0}" -f $summary.FilesMovedToRecycleBin) -ForegroundColor Green
Write-Host ("   Recycle Bin Errors: {0}" -f $summary.RecycleBinErrors) -ForegroundColor Red
Write-Host ("   Candidates Cancelled: {0}" -f $summary.CandidatesCancelled) -ForegroundColor Yellow
Write-Host
Write-Host 'Output Files' -ForegroundColor Cyan
Write-Host ("   Results CSV: {0}" -f $resultsPath)
Write-Host ("   Transcript: {0}" -f $transcriptPath)
Write-Host
Write-Host ("{0} - Script Complete" -f (Get-Date -Format 'yy-MM-dd HH:mm:ss')) -ForegroundColor Cyan
Write-Host ("   Elapsed Time: {0:hh\:mm\:ss}" -f $elapsed)
Write-Host

if ($transcriptStarted) {
    Stop-Transcript | Out-Null
}

#endregion 8. Export Results and Complete Script