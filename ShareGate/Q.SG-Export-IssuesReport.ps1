<#
.SYNOPSIS
    Export ShareGate migration logs with errors/warnings and builds consolidated summary reports.

.DESCRIPTION
    This script uses the ShareGate PowerShell module to:
      1. Find migration copy sessions within a date range using Find-CopySessions.
      2. Export each session report using Export-Report.
      3. Inspect exported report rows for Errors and/or Warnings.
      4. Optionally filter the issue rows by specific users and/or source folders.
      5. Keep only session reports that contain matching issue rows.
      6. Generate consolidated CSV reports for remediation tracking.

    Run this on each migration server that has log files you wish to report on.

.REQUIREMENTS
    - ShareGate Migrate installed on the machine running the script.
    - ShareGate PowerShell module available in the PowerShell session.
    - ShareGate Migrate Pro or Enterprise subscription for PowerShell integration.
    - PowerShell 5.1+ recommended.

.OUTPUTS
    The script creates the following files under ReportFolder:
      - Individual session CSV reports for sessions with matching issues.
      - SessionsWithIssues.csv
      - ConsolidatedIssueDetails.csv
      - ConsolidatedErrorSummary.csv
      - ConsolidatedWarningSummary.csv
      - ExecutiveIssueSummary.csv
      - UsersWithMostIssues.csv
      - FilterSummary.txt

.PARAMETER StartDate
    Start date for ShareGate copy session search.

.PARAMETER EndDate
    Optional end date. Defaults to current date/time.

.PARAMETER ReportFolder
    Output folder for exported reports and consolidated summaries.

.PARAMETER IncludeErrors
    Include error records. If neither IncludeErrors nor IncludeWarnings is supplied, both are included by default.

.PARAMETER IncludeWarnings
    Include warning records. If neither IncludeErrors nor IncludeWarnings is supplied, both are included by default.

.PARAMETER Users
    Optional list of home folder/user names to include. Example: -Users abellz,jsmith

.PARAMETER UserCsv
    Optional CSV file containing users to include. Default column name: User.

.PARAMETER UserCsvColumn
    Column name in UserCsv. Default: User.

.PARAMETER SourceFolders
    Optional list of source path/folder text filters to include. Example: -SourceFolders Finance,Legal
    Supports plain contains matching and wildcard matching if the value contains * or ?.

.PARAMETER SourceFolderCsv
    Optional CSV file containing source folder filters to include. Default column name: SourceFolder.

.PARAMETER SourceFolderCsvColumn
    Column name in SourceFolderCsv. Default: SourceFolder.

.PARAMETER ExcludeUsers
    Optional list of user/home folder names to exclude.

.PARAMETER ExcludeUserCsv
    Optional CSV file containing users to exclude. Default column name: User.

.PARAMETER ExcludeFolders
    Optional list of source path/folder text filters to exclude.

.PARAMETER ExcludeFolderCsv
    Optional CSV file containing source folder filters to exclude. Default column name: SourceFolder.

.PARAMETER UserPathSegmentIndex
    Zero-based index used to extract the user/home folder from UNC paths.
    For \\Server\Share\User\Folder, use the default value 2:
      0 = Server
      1 = Share
      2 = User
      3 = Folder

.PARAMETER KeepTempReports
    Keeps exported temporary reports even when no matching issues are found. Disabled by default.

.PARAMETER KeepAllSessionReports
    Keeps all exported session reports, even sessions with no matching issues. Disabled by default.

.PARAMETER DefaultColumns
    Uses Export-Report -DefaultColumns.

.PARAMETER NoItems
    Uses Export-Report -NoItems. Not recommended if you need item-level error detail.

.PARAMETER NoItemVersions
    Uses Export-Report -NoItemVersions.

.PARAMETER FileShareOnly
    When enabled, attempts to keep only Import from File Share rows if the exported report includes operation/task/type columns.
    If no usable operation/task/type fields exist, rows are not excluded because the script assumes sessions are file-share imports.
    Enabled by default.

.PARAMETER DisableFileShareOnly
    Disables FileShareOnly behavior.

.PARAMETER MinIssuesCount
    Minimum number of matching issues required for a ShareGate session
    report to be retained and included in the consolidated outputs.

.EXAMPLE
    .\Q.SG-Export-IssuesReport.ps1 `
        -StartDate "2026-07-01" `
        -ReportFolder "D:\ShareGateIssueReports"

.EXAMPLE
    Export errors and warnings for specific pilot users only:

    .\Q.SG-Export-IssuesReport.ps1 `
        -StartDate "2026-07-01" `
        -ReportFolder "D:\ShareGateIssueReports" `
        -Users abellz,jsmith,adoe

.EXAMPLE
    Use a CSV of users:

    .\Q.SG-Export-IssuesReport.ps1 `
        -StartDate "2026-07-01" `
        -ReportFolder "D:\ShareGateIssueReports" `
        -UserCsv ".\PilotUsers.csv"

    PilotUsers.csv example:
        User
        abellz
        jsmith
        adoe

.EXAMPLE
    Filter to Finance and Legal folders but exclude test users:

    .\Q.SG-Export-IssuesReport.ps1 `
        -StartDate "2026-07-01" `
        -ReportFolder "D:\ShareGateIssueReports" `
        -SourceFolders Finance,Legal `
        -ExcludeUsers testuser,svc_migration

.EXAMPLE
    Errors only:

    .\Q.SG-Export-IssuesReport.ps1 `
        -StartDate "2026-07-01" `
        -ReportFolder "D:\ShareGateIssueReports" `
        -IncludeErrors


.EXAMPLE
    Export only sessions with 5 or more errors.

    .\Q.SG-Export-IssuesReport.ps1 `
        -StartDate "2026-07-01" `
        -ReportFolder "D:\ShareGateIssueReports" `
        -IncludeErrors `
        -MinIssuesCount 5


.NOTES
    Recommended workflow:
      1. Run the script after a migration batch.
      2. Start with ExecutiveIssueSummary.csv for leadership/stakeholder status.
      3. Use ConsolidatedErrorSummary.csv for remediation categories.
      4. Use ConsolidatedIssueDetails.csv for source path, destination path, and item-level troubleshooting.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [datetime]$StartDate,

    [datetime]$EndDate = (Get-Date),

    [Parameter(Mandatory = $true)]
    [string]$ReportFolder,

    [switch]$IncludeErrors,
    [switch]$IncludeWarnings,

    [string[]]$Users,
    [string]$UserCsv,
    [string]$UserCsvColumn = "User",

    [string[]]$SourceFolders,
    [string]$SourceFolderCsv,
    [string]$SourceFolderCsvColumn = "SourceFolder",

    [string[]]$ExcludeUsers,
    [string]$ExcludeUserCsv,
    [string]$ExcludeUserCsvColumn = "User",

    [string[]]$ExcludeFolders,
    [string]$ExcludeFolderCsv,
    [string]$ExcludeFolderCsvColumn = "SourceFolder",

    [int]$UserPathSegmentIndex = 2,

    [switch]$KeepTempReports,
    [switch]$KeepAllSessionReports,

    [switch]$DefaultColumns,
    [switch]$NoItems,
    [switch]$NoItemVersions,

    [switch]$DisableFileShareOnly,

    [int]$MinIssuesCount = 0
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = "Stop"

function Write-Info {
    param([string]$Message)
    Write-Host $Message -ForegroundColor Cyan
}

function Write-Good {
    param([string]$Message)
    Write-Host $Message -ForegroundColor Green
}

function Write-Warn {
    param([string]$Message)
    Write-Warning $Message
}

function Get-PropertyValue {
    param(
        [Parameter(Mandatory = $true)]$Object,
        [Parameter(Mandatory = $true)][string[]]$Names
    )

    foreach ($Name in $Names) {
        $Prop = $Object.PSObject.Properties | Where-Object { $_.Name -eq $Name } | Select-Object -First 1
        if ($Prop) {
            return "$($Prop.Value)"
        }
    }

    return ""
}

function Import-FilterValuesFromCsv {
    param(
        [string]$Path,
        [string]$ColumnName
    )

    if ([string]::IsNullOrWhiteSpace($Path)) {
        return @()
    }

    if (-not (Test-Path -Path $Path)) {
        throw "CSV file not found: $Path"
    }

    $CsvRows = Import-Csv -Path $Path
    if (-not $CsvRows) {
        return @()
    }

    $Values = foreach ($CsvRow in $CsvRows) {
        $Value = Get-PropertyValue -Object $CsvRow -Names @($ColumnName)
        if (-not [string]::IsNullOrWhiteSpace($Value)) {
            $Value.Trim()
        }
    }

    return @($Values | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique)
}

function Test-TextMatchesAnyFilter {
    param(
        [string]$Text,
        [string[]]$Filters
    )

    if (-not $Filters -or $Filters.Count -eq 0) {
        return $false
    }

    if ($null -eq $Text) {
        $Text = ""
    }

    foreach ($Filter in $Filters) {
        if ([string]::IsNullOrWhiteSpace($Filter)) {
            continue
        }

        $F = $Filter.Trim()

        if ($F.Contains("*") -or $F.Contains("?")) {
            if ($Text -like $F) {
                return $true
            }
        }
        else {
            if ($Text.IndexOf($F, [System.StringComparison]::OrdinalIgnoreCase) -ge 0) {
                return $true
            }
        }
    }

    return $false
}

function Get-FileShareUserFromPath {
    param(
        [string]$SourcePath,
        [int]$SegmentIndex = 2
    )

    if ([string]::IsNullOrWhiteSpace($SourcePath)) {
        return ""
    }

    $Parts = $SourcePath -split "[\\/]" | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }

    if ($Parts.Count -gt $SegmentIndex) {
        return "$($Parts[$SegmentIndex])"
    }

    return (Split-Path -Path $SourcePath -Leaf -ErrorAction SilentlyContinue)
}

function Test-UserFilterMatch {
    param(
        [string]$SourcePath,
        [string]$ExtractedUser,
        [string[]]$UserFilters
    )

    if (-not $UserFilters -or $UserFilters.Count -eq 0) {
        return $false
    }

    foreach ($User in $UserFilters) {
        if ([string]::IsNullOrWhiteSpace($User)) {
            continue
        }

        $U = $User.Trim()

        if ($ExtractedUser -and ($ExtractedUser.Equals($U, [System.StringComparison]::OrdinalIgnoreCase))) {
            return $true
        }

        $Escaped = [regex]::Escape($U)
        if ($SourcePath -match "(?i)(^|[\\/])$Escaped($|[\\/])") {
            return $true
        }
    }

    return $false
}

function Get-IssueCategory {
    param([string]$IssueText)

    if ([string]::IsNullOrWhiteSpace($IssueText)) {
        return "Other"
    }

    switch -Regex ($IssueText) {
        "personal site|onedrive.*not found|does not have a personal site|user.*onedrive.*not.*provision" { return "OneDrive Not Provisioned" }
        "path too long|maximum path|exceeds.*path|URL.*too long|too many characters" { return "Path Too Long" }
        "invalid character|illegal character|unsupported character|character.*not.*allowed" { return "Invalid Characters" }
        "access denied|permission|unauthorized|not authorized|forbidden" { return "Access Denied / Permissions" }
        "file is locked|locked by|checked out|being used by another process" { return "Locked or Checked-Out File" }
        "unsupported.*file type|file type.*unsupported|blocked file type|extension.*blocked" { return "Unsupported or Blocked File Type" }
        "already exists|duplicate|same name" { return "Already Exists / Duplicate" }
        "quota|storage limit|not enough space|insufficient storage" { return "Quota or Storage Limit" }
        "throttl|too many requests|429|server busy|temporarily unavailable" { return "Throttling / Temporary Service Issue" }
        "network|connection|timeout|timed out|closed unexpectedly|could not connect" { return "Network or Connectivity" }
        "malware|virus|infected" { return "Malware or Security Block" }
        "too large|file size|maximum size" { return "File Size Limit" }
        "not found|cannot find|could not find|missing" { return "Source or Destination Not Found" }
        default { return "Other" }
    }
}

function Get-NormalizedIssueText {
    param([string]$IssueText)

    if ([string]::IsNullOrWhiteSpace($IssueText)) {
        return ""
    }

    $Text = $IssueText.Trim()

    # Reduce noise so identical remediation issues group together more cleanly.
    $Text = $Text -replace "https?://\S+", "<url>"
    $Text = $Text -replace '\\\\\S+', '[UNC_PATH]'
    $Text = $Text -replace "[A-Fa-f0-9]{8}-([A-Fa-f0-9]{4}-){3}[A-Fa-f0-9]{12}", "<guid>"
    $Text = $Text -replace "\b\d{4,}\b", "<number>"
    $Text = $Text -replace "\s+", " "

    return $Text
}

function Test-IsFileShareImportRow {
    param($Row)

    $CandidateText = @(
        (Get-PropertyValue -Object $Row -Names @("Operation", "Task", "Task name", "TaskName", "Type", "Migration type", "MigrationType", "Action"))
        (Get-PropertyValue -Object $Row -Names @("Source", "Source path", "SourcePath", "Source Url", "Source URL"))
    ) -join " "

    if ([string]::IsNullOrWhiteSpace($CandidateText)) {
        return $true
    }

    if ($CandidateText -match "(?i)file share|import from file|file system|\\\\") {
        return $true
    }

    # If the report does not expose a clear operation field, do not exclude rows.
    return $true
}

# Default behavior: if neither switch is provided, include both errors and warnings.
if (-not $IncludeErrors -and -not $IncludeWarnings) {
    $IncludeErrors = $true
    $IncludeWarnings = $true
}

$FileShareOnly = -not $DisableFileShareOnly

# Load filters from CSVs and merge with inline filters.
$Users = @($Users) + @(Import-FilterValuesFromCsv -Path $UserCsv -ColumnName $UserCsvColumn)
$Users = @($Users | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | ForEach-Object { $_.Trim() } | Select-Object -Unique)

$SourceFolders = @($SourceFolders) + @(Import-FilterValuesFromCsv -Path $SourceFolderCsv -ColumnName $SourceFolderCsvColumn)
$SourceFolders = @($SourceFolders | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | ForEach-Object { $_.Trim() } | Select-Object -Unique)

$ExcludeUsers = @($ExcludeUsers) + @(Import-FilterValuesFromCsv -Path $ExcludeUserCsv -ColumnName $ExcludeUserCsvColumn)
$ExcludeUsers = @($ExcludeUsers | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | ForEach-Object { $_.Trim() } | Select-Object -Unique)

$ExcludeFolders = @($ExcludeFolders) + @(Import-FilterValuesFromCsv -Path $ExcludeFolderCsv -ColumnName $ExcludeFolderCsvColumn)
$ExcludeFolders = @($ExcludeFolders | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | ForEach-Object { $_.Trim() } | Select-Object -Unique)

# Create output folders.
if (-not (Test-Path -Path $ReportFolder)) {
    New-Item -Path $ReportFolder -ItemType Directory -Force | Out-Null
}

$TempFolder = Join-Path $ReportFolder "_Temp"
if (-not (Test-Path -Path $TempFolder)) {
    New-Item -Path $TempFolder -ItemType Directory -Force | Out-Null
}

$SessionIssueSummary = New-Object System.Collections.Generic.List[object]
$ConsolidatedIssues = New-Object System.Collections.Generic.List[object]
$ProcessingErrors = New-Object System.Collections.Generic.List[object]

Write-Host ""
Write-Info "ShareGate Import from File Share Issue Export"
Write-Info "Start Date     : $StartDate"
Write-Info "End Date       : $EndDate"
Write-Info "Report Folder  : $ReportFolder"
Write-Info "Include Errors : $IncludeErrors"
Write-Info "Include Warnings: $IncludeWarnings"
Write-Info "FileShareOnly  : $FileShareOnly"
Write-Info "Min Issues    : $MinIssuesCount"
Write-Host ""

# Confirm ShareGate cmdlets exist.
$FindSessionsCommand = Get-Command -Name Find-CopySessions -ErrorAction SilentlyContinue
$ExportReportCommand = Get-Command -Name Export-Report -ErrorAction SilentlyContinue

if (-not $FindSessionsCommand -or -not $ExportReportCommand) {
    throw "ShareGate cmdlets Find-CopySessions and/or Export-Report were not found. Open the ShareGate PowerShell console or import the ShareGate module before running this script."
}

try {
    $Sessions = Find-CopySessions -From $StartDate -To $EndDate
}
catch {
    throw "Failed to retrieve ShareGate copy sessions. $($_.Exception.Message)"
}

if (-not $Sessions) {
    Write-Warn "No ShareGate copy sessions were found for the selected date range."
    return
}

$SessionCount = @($Sessions).Count
Write-Good "Sessions found: $SessionCount"
Write-Host ""

$Counter = 0

foreach ($Session in $Sessions) {
    $Counter++

    $SessionId = Get-PropertyValue -Object $Session -Names @("SessionId", "Id", "Session ID")

    if ([string]::IsNullOrWhiteSpace($SessionId)) {
        Write-Warn "Could not determine SessionId for session object. Skipping."
        continue
    }

    Write-Progress -Activity "Processing ShareGate sessions" -Status "$Counter of $SessionCount - Session $SessionId" -PercentComplete (($Counter / $SessionCount) * 100)

    $TempReportPath = Join-Path $TempFolder "$SessionId.csv"
    $FinalReportPath = Join-Path $ReportFolder "$SessionId.csv"

    try {
        $ExportParams = @{
            SessionId = $SessionId
            Path      = $TempReportPath
            Overwrite = $true
        }

        if ($DefaultColumns) { $ExportParams.DefaultColumns = $true }
        if ($NoItems) { $ExportParams.NoItems = $true }
        if ($NoItemVersions) { $ExportParams.NoItemVersions = $true }

        Export-Report @ExportParams
    }
    catch {
        $ProcessingErrors.Add([PSCustomObject]@{
            SessionId = $SessionId
            Stage     = "Export-Report"
            Error     = $_.Exception.Message
        })
        Write-Warn "Failed to export report for session $SessionId. $($_.Exception.Message)"
        continue
    }

    if (-not (Test-Path -Path $TempReportPath)) {
        $ProcessingErrors.Add([PSCustomObject]@{
            SessionId = $SessionId
            Stage     = "Export Validation"
            Error     = "Expected report was not created: $TempReportPath"
        })
        Write-Warn "Expected report was not created for session $SessionId."
        continue
    }

    try {
        $Rows = Import-Csv -Path $TempReportPath
    }
    catch {
        $ProcessingErrors.Add([PSCustomObject]@{
            SessionId = $SessionId
            Stage     = "Import-Csv"
            Error     = $_.Exception.Message
        })
        Write-Warn "Failed to import exported CSV for session $SessionId. $($_.Exception.Message)"
        continue
    }

    if (-not $Rows) {
        if (-not $KeepTempReports -and -not $KeepAllSessionReports) {
            Remove-Item -Path $TempReportPath -Force -ErrorAction SilentlyContinue
        }
        continue
    }

    $IssueRows = foreach ($Row in $Rows) {
        if ($FileShareOnly -and -not (Test-IsFileShareImportRow -Row $Row)) {
            continue
        }

        $StatusValue = Get-PropertyValue -Object $Row -Names @("Status", "Result", "State")
        $MessageValue = Get-PropertyValue -Object $Row -Names @("Message", "Messages", "Details", "Detail", "Description")
        $ErrorValue = Get-PropertyValue -Object $Row -Names @("Error", "Errors", "Error message", "ErrorMessage")
        $WarningValue = Get-PropertyValue -Object $Row -Names @("Warning", "Warnings", "Warning message", "WarningMessage")

        $SourcePath = Get-PropertyValue -Object $Row -Names @("Source path", "SourcePath", "Source", "Source Url", "Source URL", "SourcePathUrl")
        $DestinationPath = Get-PropertyValue -Object $Row -Names @("Destination path", "DestinationPath", "Destination", "Destination Url", "Destination URL", "DestinationPathUrl")
        $ItemName = Get-PropertyValue -Object $Row -Names @("Name", "Item name", "ItemName", "Title", "File name", "Folder name")
        $TaskName = Get-PropertyValue -Object $Row -Names @("Task", "Task name", "TaskName", "Operation", "Action", "Type")
        $SourceSite = Get-PropertyValue -Object $Row -Names @("Source site", "SourceSite")
        $DestinationSite = Get-PropertyValue -Object $Row -Names @("Destination site", "DestinationSite")

        $IsError = (($StatusValue -match "(?i)Error|Failed|Failure") -or (-not [string]::IsNullOrWhiteSpace($ErrorValue)))
        $IsWarning = (($StatusValue -match "(?i)Warning") -or (-not [string]::IsNullOrWhiteSpace($WarningValue)))

        if (-not (($IncludeErrors -and $IsError) -or ($IncludeWarnings -and $IsWarning))) {
            continue
        }

        $ExtractedUser = Get-FileShareUserFromPath -SourcePath $SourcePath -SegmentIndex $UserPathSegmentIndex

        $IncludeRecord = $true

        if ($Users -and $Users.Count -gt 0) {
            $IncludeRecord = Test-UserFilterMatch -SourcePath $SourcePath -ExtractedUser $ExtractedUser -UserFilters $Users
        }

        if ($IncludeRecord -and $SourceFolders -and $SourceFolders.Count -gt 0) {
            $IncludeRecord = Test-TextMatchesAnyFilter -Text $SourcePath -Filters $SourceFolders
        }

        if ($IncludeRecord -and $ExcludeUsers -and $ExcludeUsers.Count -gt 0) {
            if (Test-UserFilterMatch -SourcePath $SourcePath -ExtractedUser $ExtractedUser -UserFilters $ExcludeUsers) {
                $IncludeRecord = $false
            }
        }

        if ($IncludeRecord -and $ExcludeFolders -and $ExcludeFolders.Count -gt 0) {
            if (Test-TextMatchesAnyFilter -Text $SourcePath -Filters $ExcludeFolders) {
                $IncludeRecord = $false
            }
        }

        if (-not $IncludeRecord) {
            continue
        }

        $IssueType = if ($IsError) { "Error" } elseif ($IsWarning) { "Warning" } else { "Issue" }

        $RawIssueText = @($StatusValue, $ErrorValue, $WarningValue, $MessageValue) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
        $IssueText = ($RawIssueText -join " | ").Trim()
        $NormalizedIssueText = Get-NormalizedIssueText -IssueText $IssueText
        $Category = Get-IssueCategory -IssueText $IssueText

        [PSCustomObject]@{
            SessionId            = $SessionId
            IssueType            = $IssueType
            Category             = $Category
            ExtractedUser        = $ExtractedUser
            Status               = $StatusValue
            IssueText            = $IssueText
            NormalizedIssueText  = $NormalizedIssueText
            Error                = $ErrorValue
            Warning              = $WarningValue
            Message              = $MessageValue
            SourcePath           = $SourcePath
            DestinationPath      = $DestinationPath
            ItemName             = $ItemName
            TaskName             = $TaskName
            SourceSite           = $SourceSite
            DestinationSite      = $DestinationSite
            ReportFile           = ""
        }
    }

    $IssueRows = @($IssueRows)

    $ErrorCount = @(
        $IssueRows | Where-Object { $_.IssueType -eq "Error" }
    ).Count

    $WarningCount = @(
        $IssueRows | Where-Object { $_.IssueType -eq "Warning" }
    ).Count

    # Determine session issue count based on selected filters
    if ($IncludeErrors -and $IncludeWarnings) {
        $MatchingIssueCount = $IssueRows.Count
    }
    elseif ($IncludeErrors) {
        $MatchingIssueCount = $ErrorCount
    }
    elseif ($IncludeWarnings) {
        $MatchingIssueCount = $WarningCount
    }
    else {
        $MatchingIssueCount = $IssueRows.Count
    }

    if ($MatchingIssueCount -lt $MinIssuesCount) {

        Write-Host (
            "Skipping session {0}. Matching issues ({1}) " +
            "below MinIssuesCount ({2})." -f
            $SessionId,
            $MatchingIssueCount,
            $MinIssuesCount
        ) -ForegroundColor DarkGray

        if ($KeepAllSessionReports) {
            Move-Item -Path $TempReportPath -Destination $FinalReportPath -Force
        }
        elseif (-not $KeepTempReports) {
            Remove-Item -Path $TempReportPath -Force -ErrorAction SilentlyContinue
        }

        continue
    }

    if ($IssueRows.Count -gt 0) {
        Move-Item -Path $TempReportPath -Destination $FinalReportPath -Force

        <#
        $ErrorCount = @($IssueRows | Where-Object { $_.IssueType -eq "Error" }).Count
        $WarningCount = @($IssueRows | Where-Object { $_.IssueType -eq "Warning" }).Count
        #>

        foreach ($Issue in $IssueRows) {
            $Issue.ReportFile = $FinalReportPath
            $ConsolidatedIssues.Add($Issue)
        }

        $SessionIssueSummary.Add([PSCustomObject]@{
            SessionId     = $SessionId
            Errors        = $ErrorCount
            Warnings      = $WarningCount
            TotalIssues   = $IssueRows.Count
            UniqueUsers   = @($IssueRows | Where-Object { $_.ExtractedUser } | Select-Object -ExpandProperty ExtractedUser -Unique).Count
            ReportFile    = $FinalReportPath
        })

        Write-Host "Issues found in session $SessionId - Errors: $ErrorCount Warnings: $WarningCount" -ForegroundColor Yellow
    }
    else {
        if ($KeepAllSessionReports) {
            Move-Item -Path $TempReportPath -Destination $FinalReportPath -Force
        }
        elseif (-not $KeepTempReports) {
            Remove-Item -Path $TempReportPath -Force -ErrorAction SilentlyContinue
        }
    }
}

Write-Progress -Activity "Processing ShareGate sessions" -Completed

# Output files.
$SessionsWithIssuesPath = Join-Path $ReportFolder "SessionsWithIssues.csv"
$ConsolidatedDetailsPath = Join-Path $ReportFolder "ConsolidatedIssueDetails.csv"
$ErrorSummaryPath = Join-Path $ReportFolder "ConsolidatedErrorSummary.csv"
$WarningSummaryPath = Join-Path $ReportFolder "ConsolidatedWarningSummary.csv"
$ExecutiveSummaryPath = Join-Path $ReportFolder "ExecutiveIssueSummary.csv"
$UsersWithMostIssuesPath = Join-Path $ReportFolder "UsersWithMostIssues.csv"
$ProcessingErrorsPath = Join-Path $ReportFolder "ProcessingErrors.csv"
$FilterSummaryPath = Join-Path $ReportFolder "FilterSummary.txt"

$SessionIssueSummary |
    Sort-Object TotalIssues -Descending |
    Export-Csv -Path $SessionsWithIssuesPath -NoTypeInformation -Encoding UTF8

$ConsolidatedIssues |
    Sort-Object IssueType, Category, ExtractedUser, SourcePath |
    Export-Csv -Path $ConsolidatedDetailsPath -NoTypeInformation -Encoding UTF8

$ConsolidatedIssues |
    Where-Object { $_.IssueType -eq "Error" } |
    Group-Object NormalizedIssueText |
    ForEach-Object {
        $Group = $_.Group
        [PSCustomObject]@{
            IssueType        = "Error"
            Category         = ($Group | Group-Object Category | Sort-Object Count -Descending | Select-Object -First 1 -ExpandProperty Name)
            Count            = $_.Count
            UniqueUsers      = @($Group | Where-Object { $_.ExtractedUser } | Select-Object -ExpandProperty ExtractedUser -Unique).Count
            IssueText        = ($Group | Select-Object -First 1 -ExpandProperty IssueText)
            NormalizedIssue  = $_.Name
            SessionsImpacted = (($Group | Select-Object -ExpandProperty SessionId -Unique) -join "; ")
            SampleUser       = ($Group | Where-Object { $_.ExtractedUser } | Select-Object -First 1 -ExpandProperty ExtractedUser)
            SampleSourcePath = ($Group | Where-Object { $_.SourcePath } | Select-Object -First 1 -ExpandProperty SourcePath)
            SampleDestPath   = ($Group | Where-Object { $_.DestinationPath } | Select-Object -First 1 -ExpandProperty DestinationPath)
            SampleItemName   = ($Group | Where-Object { $_.ItemName } | Select-Object -First 1 -ExpandProperty ItemName)
            SampleReportFile = ($Group | Select-Object -First 1 -ExpandProperty ReportFile)
        }
    } |
    Sort-Object Count -Descending |
    Export-Csv -Path $ErrorSummaryPath -NoTypeInformation -Encoding UTF8

$ConsolidatedIssues |
    Where-Object { $_.IssueType -eq "Warning" } |
    Group-Object NormalizedIssueText |
    ForEach-Object {
        $Group = $_.Group
        [PSCustomObject]@{
            IssueType        = "Warning"
            Category         = ($Group | Group-Object Category | Sort-Object Count -Descending | Select-Object -First 1 -ExpandProperty Name)
            Count            = $_.Count
            UniqueUsers      = @($Group | Where-Object { $_.ExtractedUser } | Select-Object -ExpandProperty ExtractedUser -Unique).Count
            IssueText        = ($Group | Select-Object -First 1 -ExpandProperty IssueText)
            NormalizedIssue  = $_.Name
            SessionsImpacted = (($Group | Select-Object -ExpandProperty SessionId -Unique) -join "; ")
            SampleUser       = ($Group | Where-Object { $_.ExtractedUser } | Select-Object -First 1 -ExpandProperty ExtractedUser)
            SampleSourcePath = ($Group | Where-Object { $_.SourcePath } | Select-Object -First 1 -ExpandProperty SourcePath)
            SampleDestPath   = ($Group | Where-Object { $_.DestinationPath } | Select-Object -First 1 -ExpandProperty DestinationPath)
            SampleItemName   = ($Group | Where-Object { $_.ItemName } | Select-Object -First 1 -ExpandProperty ItemName)
            SampleReportFile = ($Group | Select-Object -First 1 -ExpandProperty ReportFile)
        }
    } |
    Sort-Object Count -Descending |
    Export-Csv -Path $WarningSummaryPath -NoTypeInformation -Encoding UTF8

$ConsolidatedIssues |
    Group-Object Category |
    ForEach-Object {
        $Group = $_.Group
        [PSCustomObject]@{
            Category       = $_.Name
            TotalIssues    = $_.Count
            Errors         = @($Group | Where-Object { $_.IssueType -eq "Error" }).Count
            Warnings       = @($Group | Where-Object { $_.IssueType -eq "Warning" }).Count
            UniqueUsers    = @($Group | Where-Object { $_.ExtractedUser } | Select-Object -ExpandProperty ExtractedUser -Unique).Count
            UniqueSessions = @($Group | Select-Object -ExpandProperty SessionId -Unique).Count
            ExampleIssue   = ($Group | Select-Object -First 1 -ExpandProperty IssueText)
            ExamplePath    = ($Group | Where-Object { $_.SourcePath } | Select-Object -First 1 -ExpandProperty SourcePath)
        }
    } |
    Sort-Object TotalIssues -Descending |
    Export-Csv -Path $ExecutiveSummaryPath -NoTypeInformation -Encoding UTF8

$ConsolidatedIssues |
    Where-Object { $_.ExtractedUser } |
    Group-Object ExtractedUser |
    ForEach-Object {
        $Group = $_.Group
        [PSCustomObject]@{
            User          = $_.Name
            TotalIssues   = $_.Count
            Errors        = @($Group | Where-Object { $_.IssueType -eq "Error" }).Count
            Warnings      = @($Group | Where-Object { $_.IssueType -eq "Warning" }).Count
            TopCategory   = ($Group | Group-Object Category | Sort-Object Count -Descending | Select-Object -First 1 -ExpandProperty Name)
            Sessions      = (($Group | Select-Object -ExpandProperty SessionId -Unique) -join "; ")
            ExamplePath   = ($Group | Where-Object { $_.SourcePath } | Select-Object -First 1 -ExpandProperty SourcePath)
        }
    } |
    Sort-Object TotalIssues -Descending |
    Export-Csv -Path $UsersWithMostIssuesPath -NoTypeInformation -Encoding UTF8

if ($ProcessingErrors.Count -gt 0) {
    $ProcessingErrors |
        Export-Csv -Path $ProcessingErrorsPath -NoTypeInformation -Encoding UTF8
}

$FilterSummary = @"
ShareGate Import from File Share Issue Export - Filter Summary
Generated: $(Get-Date)

Date Range:
  StartDate: $StartDate
  EndDate  : $EndDate

Output Folder:
  $ReportFolder

Included Issue Types:
  IncludeErrors  : $IncludeErrors
  IncludeWarnings: $IncludeWarnings
  MinIssuesCount : $MinIssuesCount

User Filters:
  Users        : $($Users -join ', ')
  ExcludeUsers : $($ExcludeUsers -join ', ')

Folder Filters:
  SourceFolders  : $($SourceFolders -join ', ')
  ExcludeFolders : $($ExcludeFolders -join ', ')

User Path Segment Index:
  $UserPathSegmentIndex

Result Counts:
  SessionsFound           : $SessionCount
  SessionsWithIssues      : $($SessionIssueSummary.Count)
  ConsolidatedIssueRecords: $($ConsolidatedIssues.Count)
  ProcessingErrors        : $($ProcessingErrors.Count)
"@

$FilterSummary | Out-File -FilePath $FilterSummaryPath -Encoding UTF8

try {
    if ((Get-ChildItem -Path $TempFolder -Force -ErrorAction SilentlyContinue | Measure-Object).Count -eq 0) {
        Remove-Item -Path $TempFolder -Force -ErrorAction SilentlyContinue
    }
}
catch {}

Write-Host ""
Write-Good "Completed ShareGate issue export."
Write-Host ""
Write-Info "Generated files:"
Write-Host "  Sessions with issues        : $SessionsWithIssuesPath"
Write-Host "  Consolidated issue details  : $ConsolidatedDetailsPath"
Write-Host "  Consolidated error summary  : $ErrorSummaryPath"
Write-Host "  Consolidated warning summary: $WarningSummaryPath"
Write-Host "  Executive issue summary     : $ExecutiveSummaryPath"
Write-Host "  Users with most issues      : $UsersWithMostIssuesPath"
Write-Host "  Filter summary              : $FilterSummaryPath"
if ($ProcessingErrors.Count -gt 0) {
    Write-Host "  Processing errors           : $ProcessingErrorsPath"
}
Write-Host ""
Write-Host "Total sessions with issues: $($SessionIssueSummary.Count)"
Write-Host "Total issue records       : $($ConsolidatedIssues.Count)"
Write-Host ""
