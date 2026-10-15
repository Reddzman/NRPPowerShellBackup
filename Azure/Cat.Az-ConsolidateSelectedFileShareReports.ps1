<#
.SYNOPSIS
    Consolidates selected ADLS Gen2 file-share discovery reports after validating
    their names and report contents against an input CSV.

.DESCRIPTION
    Reads reportNaming, Full UNC Path, serverUNC, and shareName from an input CSV.
    It scans the directories, files, and permissions folders under each discovery
    execution folder, selects reports matching the naming templates, and validates
    the first ValidationRowLimit data rows against the corresponding target share.

    A report with a missing or mismatched server, share, or path is not copied.
    Multiple source reports that would use the same destination filename are all
    skipped for review. Existing destination files are never overwritten. The
    remaining rows are assumed to match when the sampled rows pass validation.

    Content validation reads each filename-matched report as a stream and stops
    after the configured number of data rows, without saving the full report locally.
    -WhatIf authenticates and performs read-only report discovery, then lists reports
    eligible for copying based on filename, folder, destination-collision, and
    destination-existence checks. It skips content downloads and performs no Azure
    directory creation or copy.

.AUTHENTICATION
        The script uses two separate authentication flows:

            - Az PowerShell: If the current session has no Az context, Connect-AzAccount
                uses device-code authentication for TenantId and SubscriptionId. Follow the
                URL and code shown in the terminal in any browser. If an Az context already
                exists, the script reuses it and selects SubscriptionId without prompting.
            - AzCopy: On normal copy runs, the script separately runs `azcopy login`
                before scanning reports. Complete its device-code prompt too. The Az
                Storage token message does not mean AzCopy is authenticated.

        Complete each device-code sign-in promptly before discovery starts; if login
        fails or the code expires, the script stops before scanning. With -WhatIf,
        Azure authentication is still required for read-only discovery, but the script
        skips the copy-only Storage token and AzCopy login because no copies will run.
        Fully unattended runs require a non-interactive identity configured for both
        Az PowerShell and AzCopy.

    Dependencies:
      - PowerShell 7.2 or later
      - Az.Accounts and Az.Storage PowerShell modules
      - AzCopy available on PATH
      - General functions module at ..\General\Cat.GeneralFunctions.psm1

        Known limitation:
            - Only the configured sample is checked. The script assumes later rows follow
                the same target-share rule and records that assumption in its summary.

.PARAMETER SubscriptionId
    Azure subscription used for the storage account.

.PARAMETER ResourceGroupName
    Resource group containing the storage account.

.PARAMETER StorageAccountName
    ADLS Gen2-enabled storage account.

.PARAMETER ContainerName
    ADLS Gen2 filesystem (container) name.

.PARAMETER SourcePath
    Root folder containing discovery execution folders.

.PARAMETER DestinationPath
    Root folder for consolidated directories, files, and permissions reports.

.PARAMETER InputCsvPath
    CSV containing Full UNC Path, serverUNC, shareName, and reportNaming columns.

.PARAMETER OutputReportPath
    Local folder for CopySummary.csv and ReviewItems.csv.

.PARAMETER ValidationRowLimit
    Number of data rows to validate in each report. Defaults to 100; the header is
    not counted. A report with fewer rows is validated to the end.

.PARAMETER TenantId
    Tenant used by Azure device-code authentication and AzCopy authentication.

.EXAMPLE
    CD D:\GitHub\CatPowerShell\Azure
    .\Cat.Az-ConsolidateSelectedFileShareReports.ps1 `
      -SubscriptionId "a01e0667-ed24-40ac-9e6e-a6a8793f3686" `
      -ResourceGroupName "rg-prac-prod-file-share-migration-discovery-01" `
      -StorageAccountName "stdatamigrationeus2" `
      -ContainerName "datamigration" `
      -SourcePath "bronze/SimpsonStrongTie-BranchDiscovery/RawData" `
      -DestinationPath "bronze/SST-Canada/2610081105" `
      -InputCsvPath "..\_Input Files\sst-discoveryreports-canada-261008.csv" `
      -OutputReportPath "..\_Output Files" `
      -ValidationRowLimit 100 `
      -TenantId "d0659de2-684e-49bd-9b1f-1fd4cd0942d9" `
      -WhatIf

Table of Contents
    1. Script Parameters
    2. Start Date / Time
    3. Snapin(s), Modules & Environment Connection(s)
    4. Input & Output Files
    5. Variables
    6. Functions
    7. Main Script
    8. Completed Script
#>

#region 1. Script Parameters
[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
param (
    [string]$SubscriptionId = "a01e0667-ed24-40ac-9e6e-a6a8793f3686",
    [string]$ResourceGroupName = "rg-prac-prod-file-share-migration-discovery-01",
    [string]$StorageAccountName = "stdatamigrationeus2",
    [string]$ContainerName = "datamigration",
    [string]$SourcePath = "bronze/SimpsonStrongTie-BranchDiscovery/RawData",
    [string]$DestinationPath = "bronze/SST-Canada/2610081105",
    [string]$InputCsvPath = "..\_Input Files\sst-discoveryreports-canada-261008.csv",
    [string]$OutputReportPath = "..\_Output Files",
    [ValidateRange(1, 2147483647)]
    [int]$ValidationRowLimit = 100,
    [string]$TenantId = "d0659de2-684e-49bd-9b1f-1fd4cd0942d9"
)
#endregion 1. Script Parameters

#region 2. Start Date / Time

$scriptStart = $null
$transcriptStarted = $false
$runCompleted = $false
# Minimal uses a single-line ANSI progress renderer when the host supports VT; avoid Classic's multi-line updates.
if ($Host.UI.SupportsVirtualTerminal) {
    $PSStyle.Progress.View = 'Minimal'
}
Get-Module Cat.* | Remove-Module -Force -WhatIf:$false
Import-Module (Join-Path $PSScriptRoot "..\General\Cat.GeneralFunctions.psm1") -Force
$scriptStart = Start-CatScript
$reportTimestamp = Get-Date -Format "yyMMddHHmmss"
New-Item -ItemType Directory -Path (Join-Path $PSScriptRoot "..\_Log Files") -Force -WhatIf:$false | Out-Null
Start-Transcript -Path (Join-Path $PSScriptRoot "..\_Log Files\Cat.Az-ConsolidateSelectedFileShareReports-$($scriptStart[1]).log") -WhatIf:$false -ErrorAction Stop | Out-Null
$transcriptStarted = $true

#endregion 2. Start Date / Time

try {
    #region 3. Snapin(s), Modules & Environment Connection(s)

    Write-Host
    Write-Host ("   {0} - Authentication" -f (Get-Date -Format "yy-MM-dd HH:mm:ss")) -ForegroundColor Magenta
    Write-Host ("      {0} - Loading Azure modules and connecting to the selected subscription..." -f (Get-Date -Format "yy-MM-dd HH:mm:ss"))
    Import-Module Az.Accounts -ErrorAction Stop
    Import-Module Az.Storage -ErrorAction Stop

    if (-not (Get-AzContext)) {
        # Device-code login provides a URL and code that can be opened in any browser, including from terminal-only hosts.
        # WhatIf still needs Azure credentials for read-only inventory, so opt authentication out of ShouldProcess.
        Connect-AzAccount `
            -Tenant $TenantId `
            -Subscription $SubscriptionId `
            -UseDeviceAuthentication `
            -WhatIf:$false `
            -ErrorAction Stop | Out-Null
    }
    Set-AzContext -SubscriptionId $SubscriptionId -WhatIf:$false -ErrorAction Stop | Out-Null

    Get-AzStorageAccount `
        -ResourceGroupName $ResourceGroupName `
        -Name $StorageAccountName `
        -ErrorAction Stop | Out-Null

    $ctx = New-AzStorageContext `
        -StorageAccountName $StorageAccountName `
        -UseConnectedAccount

    # Cache a Storage data-plane token before scanning so authentication failures surface before a long inventory run.
    $storageTokenValue = $null
    $storageTokenExpiresOn = [DateTimeOffset]::MinValue
    if (-not $WhatIfPreference) {
        $storageToken = Get-AzAccessToken -ResourceUrl "https://storage.azure.com/" -ErrorAction Stop
        $storageTokenValue = if ($storageToken.Token -is [System.Security.SecureString]) {
            [System.Net.NetworkCredential]::new("", $storageToken.Token).Password
        } else {
            [string]$storageToken.Token
        }
        $storageTokenExpiresOn = $storageToken.ExpiresOn
        Write-Host ("      {0} - Azure Storage access token acquired." -f (Get-Date -Format "yy-MM-dd HH:mm:ss"))

        # Authenticate AzCopy now, before discovery can take a long time and let its device code expire.
        if (-not (Get-Command azcopy -ErrorAction SilentlyContinue)) {
            throw "AzCopy was not found on PATH. Install AzCopy or add it to PATH before running this script."
        }
        Write-Host ("      {0} - Authenticating AzCopy before report discovery..." -f (Get-Date -Format "yy-MM-dd HH:mm:ss"))
        & azcopy login --tenant-id $TenantId
        if ($LASTEXITCODE -ne 0) {
            throw "AzCopy login failed (exit code $LASTEXITCODE)."
        }
        Write-Host ("      {0} - AzCopy authentication successful." -f (Get-Date -Format "yy-MM-dd HH:mm:ss"))
    }

    Write-Host ("      {0} - Azure context ready for {1} / {2}." -f (Get-Date -Format "yy-MM-dd HH:mm:ss"), $StorageAccountName, $ContainerName)
    #endregion 3. Snapin(s), Modules & Environment Connection(s)

    #region 4. Input & Output Files

    Write-Host
    Write-Host ("   {0} - Input and output files" -f (Get-Date -Format "yy-MM-dd HH:mm:ss")) -ForegroundColor Magenta
    # Resolve relative paths against the script so execution is independent of the caller's current directory.
    $InputCsvPath = if ([System.IO.Path]::IsPathRooted($InputCsvPath)) {
        $InputCsvPath
    } else {
        Join-Path $PSScriptRoot $InputCsvPath
    }
    $OutputReportPath = if ([System.IO.Path]::IsPathRooted($OutputReportPath)) {
        $OutputReportPath
    } else {
        Join-Path $PSScriptRoot $OutputReportPath
    }

    if (-not (Test-Path -LiteralPath $InputCsvPath -PathType Leaf)) {
        throw "Input CSV was not found: $InputCsvPath"
    }

    $inputRows = @(Import-Csv -LiteralPath $InputCsvPath -ErrorAction Stop)
    if ($inputRows.Count -eq 0) {
        throw "Input CSV contains no data rows: $InputCsvPath"
    }

    $requiredColumns = @("Full UNC Path", "serverUNC", "shareName", "reportNaming")
    $availableColumns = @($inputRows[0].PSObject.Properties.Name)
    foreach ($column in $requiredColumns) {
        if ($availableColumns -notcontains $column) {
            throw "Input CSV is missing required column '$column'."
        }
    }

    New-Item -ItemType Directory -Path $OutputReportPath -Force -WhatIf:$false | Out-Null
    Write-Host ("      {0} - Loaded {1} input rows from '{2}'." -f (Get-Date -Format "yy-MM-dd HH:mm:ss"), $inputRows.Count, $InputCsvPath)
    Write-Host ("      {0} - Local reports will be written to '{1}'." -f (Get-Date -Format "yy-MM-dd HH:mm:ss"), $OutputReportPath)
    #endregion 4. Input & Output Files

    #region 5. Variables

    # Normalize ADLS paths once here; the storage commands and URL builder expect forward slashes without outer slashes.
    $SourcePath = ($SourcePath -replace '\\','/').Trim('/')
    $DestinationPath = ($DestinationPath -replace '\\','/').Trim('/')

    $CopySummary = [System.Collections.Generic.List[object]]::new()
    $ReviewItems = [System.Collections.Generic.List[object]]::new()
    $httpClient = $null
    $compiledPatterns = [System.Collections.Generic.List[object]]::new()

    #endregion 5. Variables

    #region 6. Functions

    # Write one concise, timestamped line per meaningful phase or result; detailed records belong in CSV outputs.
    #region 6.1 Write-RunStatus
    function Write-RunStatus {
        [CmdletBinding()]
        param (
            [Parameter(Mandatory)][string]$Message,
            [ValidateSet("Info", "Job", "Activity", "Detail", "Success", "Warning", "Error")]
            [string]$Level = "Info",
            [ValidateRange(0, 8)]
            [int]$Indent = 0
        )

        $color = switch ($Level) {
            "Info" { "Cyan" }
            "Job" { "Magenta" }
            "Activity" { $null }
            "Detail" { "DarkGray" }
            "Success" { "Green" }
            "Warning" { "Yellow" }
            "Error" { "Red" }
        }
        $indentText = " " * (3 * $Indent)
        $timestamp = Get-Date -Format "yy-MM-dd HH:mm:ss"
        if ($Level -in @("Info", "Job")) { Write-Host }
        $statusLine = "{0}{1} - {2}" -f $indentText, $timestamp, $Message
        if ($Level -eq "Activity") {
            Write-Host $statusLine
        }
        else {
            Write-Host $statusLine -ForegroundColor $color
        }
    }
    #endregion 6.1 Write-RunStatus

    #region 6.2 Test-PathWithinRoot
    # This check is intentionally a helper because share-root containment is reused for multiple report columns.
    function Test-PathWithinRoot {
        param(
            [string]$CandidatePath,
            [string]$RootPath
        )

        # Permission exports prepend a PowerShell provider name; remove it before comparing UNC paths.
        $candidate = (($CandidatePath -replace '^Microsoft\.PowerShell\.Core\\FileSystem', '') -replace '/', '\').TrimEnd('\')
        $root = ($RootPath -replace '/', '\').TrimEnd('\')
        if ([string]::IsNullOrWhiteSpace($candidate) -or [string]::IsNullOrWhiteSpace($root)) { return $false }
        $candidate.Equals($root, [System.StringComparison]::OrdinalIgnoreCase) -or
            $candidate.StartsWith("$root\", [System.StringComparison]::OrdinalIgnoreCase)
    }
    #endregion 6.2 Test-PathWithinRoot

    #endregion 6. Functions

    #region 7. Main Script

    Write-RunStatus "Variables" -Level Info
    Write-RunStatus "Script configuration" -Level Job -Indent 1
    Write-RunStatus "Subscription: $SubscriptionId" -Level Activity -Indent 2
    Write-RunStatus "Resource group: $ResourceGroupName" -Level Activity -Indent 2
    Write-RunStatus "Storage account / container: $StorageAccountName / $ContainerName" -Level Activity -Indent 2
    Write-RunStatus "Source: $SourcePath" -Level Activity -Indent 2
    Write-RunStatus "Destination: $DestinationPath" -Level Activity -Indent 2
    Write-RunStatus "Input CSV: $InputCsvPath" -Level Activity -Indent 2
    Write-RunStatus "Output folder: $OutputReportPath" -Level Activity -Indent 2
    Write-RunStatus "Validation rows: $ValidationRowLimit; WhatIf: $WhatIfPreference" -Level Activity -Indent 2

    Write-RunStatus "Main Script - scanning and copying files" -Level Info
    Write-RunStatus "Reviewing for selected reports" -Level Job -Indent 1
    Write-RunStatus "Validating ADLS filesystem access and source path..." -Level Activity -Indent 2
    Get-AzDataLakeGen2ChildItem -Context $ctx -FileSystem $ContainerName -Path "/" -ErrorAction Stop | Out-Null
    if (-not (Get-AzDataLakeGen2Item -Context $ctx -FileSystem $ContainerName -Path $SourcePath -ErrorAction SilentlyContinue)) {
        throw "SourcePath '$SourcePath' does not exist in filesystem '$ContainerName'."
    }
    Write-RunStatus "Source path is available." -Level Activity -Indent 2
    Write-RunStatus "Compiling report-name patterns..." -Level Activity -Indent 2

    # Compile the CSV naming templates once. Unknown or malformed placeholders are
    # surfaced for review rather than being interpreted as regular expressions.
    for ($rowIndex = 0; $rowIndex -lt $inputRows.Count; $rowIndex++) {
        $inputRow = $inputRows[$rowIndex]
        $template = [string]$inputRow.reportNaming
        $hasReportType = ([regex]::Matches($template, '\[File,Dir,Perm\]')).Count -eq 1
        $hasTimestamp = ([regex]::Matches($template, '\[DateTimeStamp\]')).Count -eq 1

        if (-not $hasReportType -or -not $hasTimestamp) {
            $ReviewItems.Add([pscustomobject]@{
                InputRow = $rowIndex + 1
                ReportNaming = $template
                FullUNCPath = $inputRow.'Full UNC Path'
                serverUNC = $inputRow.serverUNC
                shareName = $inputRow.shareName
                SourcePath = ""
                DestinationPath = ""
                ReportType = ""
                Reason = "Malformed reportNaming template"
                Details = "Expected exactly one [File,Dir,Perm] and one [DateTimeStamp] placeholder."
            })
            continue
        }

        $regexText = [regex]::Escape($template).
            Replace('\[File,Dir,Perm]', '(?<ReportType>File|Dir|Perm)').
            Replace('\[DateTimeStamp]', '(?<Timestamp>\d{12})')

        $compiledPatterns.Add([pscustomobject]@{
            InputIndex = $rowIndex
            InputRow = $inputRow
            Regex = [regex]::new("^$regexText$", [System.Text.RegularExpressions.RegexOptions]::IgnoreCase -bor [System.Text.RegularExpressions.RegexOptions]::CultureInvariant)
        })
    }

    Write-RunStatus "Compiled $($compiledPatterns.Count) valid report-name patterns." -Level Activity -Indent 2

    #region 7.1 Source Discovery and Filename Matching

    $sourceExecutions = @(Get-AzDataLakeGen2ChildItem `
        -Context $ctx `
        -FileSystem $ContainerName `
        -Path $SourcePath `
        -ErrorAction Stop | Where-Object IsDirectory)

    $candidates = [System.Collections.Generic.List[object]]::new()
    $matchedTypesByInput = @{}
    $typeToFolder = @{
        File = "files"
        Dir = "directories"
        Perm = "permissions"
    }
    Write-RunStatus "Scanning $($sourceExecutions.Count) discovery execution folders..." -Level Activity -Indent 2
    $executionNumber = 0
    foreach ($execution in $sourceExecutions) {
        $executionNumber++
        Write-Progress `
            -Id 1 `
            -Activity "Scanning discovery reports" `
            -Status ("Execution folder {0,4} of {1,4}" -f $executionNumber, $sourceExecutions.Count) `
            -PercentComplete (($executionNumber / [math]::Max(1, $sourceExecutions.Count)) * 100)

        $executionPath = ($execution.Path -replace '\\','/').Trim('/')
        foreach ($reportType in @("File", "Dir", "Perm")) {
            $sourceFolder = "$executionPath/$($typeToFolder[$reportType])"
            try {
                $sourceFiles = @(Get-AzDataLakeGen2ChildItem `
                    -Context $ctx `
                    -FileSystem $ContainerName `
                    -Path $sourceFolder `
                    -ErrorAction Stop | Where-Object { -not $_.IsDirectory })
            }
            catch {
                continue
            }

            foreach ($sourceFile in $sourceFiles) {
                $fileName = ($sourceFile.Path -replace '\\','/').TrimEnd('/').Split('/')[-1]
                $patternMatches = [System.Collections.Generic.List[object]]::new()
                foreach ($pattern in $compiledPatterns) {
                    $match = $pattern.Regex.Match($fileName)
                    if ($match.Success -and $typeToFolder[$match.Groups['ReportType'].Value] -eq $typeToFolder[$reportType]) {
                        $patternMatches.Add([pscustomobject]@{
                            Pattern = $pattern
                            Timestamp = $match.Groups['Timestamp'].Value
                            ReportType = $match.Groups['ReportType'].Value
                        })
                    }
                }

                if ($patternMatches.Count -eq 0) { continue }

                if ($patternMatches.Count -gt 1) {
                    $ReviewItems.Add([pscustomobject]@{
                        InputRow = ""
                        ReportNaming = (($patternMatches | ForEach-Object { $_.Pattern.InputRow.reportNaming }) -join " | ")
                        FullUNCPath = ""
                        serverUNC = (($patternMatches | ForEach-Object { $_.Pattern.InputRow.serverUNC } | Sort-Object -Unique) -join " | ")
                        shareName = (($patternMatches | ForEach-Object { $_.Pattern.InputRow.shareName } | Sort-Object -Unique) -join " | ")
                        SourcePath = $sourceFile.Path
                        DestinationPath = ""
                        ReportType = $reportType
                        Reason = "Filename matches multiple input rows"
                        Details = "No copy attempted because the report's target share is ambiguous."
                    })
                    continue
                }

                $match = $patternMatches[0]
                $inputIndex = $match.Pattern.InputIndex
                if (-not $matchedTypesByInput.ContainsKey($inputIndex)) {
                    $matchedTypesByInput[$inputIndex] = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
                }
                [void]$matchedTypesByInput[$inputIndex].Add($match.ReportType)

                $destinationFolder = "$DestinationPath/$($typeToFolder[$reportType])"
                $destinationFile = "$destinationFolder/$fileName"
                $candidates.Add([pscustomobject]@{
                    InputIndex = $inputIndex
                    InputRow = $match.Pattern.InputRow
                    ReportNaming = $match.Pattern.InputRow.reportNaming
                    serverUNC = $match.Pattern.InputRow.serverUNC
                    shareName = $match.Pattern.InputRow.shareName
                    SourcePath = ($sourceFile.Path -replace '\\','/').Trim('/')
                    SourceExecution = $executionPath
                    FileName = $fileName
                    DestinationPath = $destinationFile
                    ReportType = $match.ReportType
                    Timestamp = $match.Timestamp
                })
            }
        }
    }
    Write-Progress -Id 1 -Activity "Scanning discovery reports" -Completed

    foreach ($pattern in $compiledPatterns) {
        $foundTypes = $matchedTypesByInput[$pattern.InputIndex]
        foreach ($reportType in @("File", "Dir", "Perm")) {
            if (-not $foundTypes -or -not $foundTypes.Contains($reportType)) {
                $ReviewItems.Add([pscustomobject]@{
                    InputRow = $pattern.InputIndex + 1
                    ReportNaming = $pattern.InputRow.reportNaming
                    FullUNCPath = $pattern.InputRow.'Full UNC Path'
                    serverUNC = $pattern.InputRow.serverUNC
                    shareName = $pattern.InputRow.shareName
                    SourcePath = ""
                    DestinationPath = ""
                    ReportType = $reportType
                    Reason = "No matching source report found"
                    Details = "Expected a report with this type token and a 12-digit timestamp."
                })
            }
        }
    }

    # Never choose an arbitrary source when multiple reports map to one destination.
    $duplicateDestinations = @{}
    foreach ($candidate in $candidates) {
        if (-not $duplicateDestinations.ContainsKey($candidate.DestinationPath)) {
            $duplicateDestinations[$candidate.DestinationPath] = [System.Collections.Generic.List[object]]::new()
        }
        $duplicateDestinations[$candidate.DestinationPath].Add($candidate)
    }

    #endregion 7.1 Source Discovery and Filename Matching

    #region 7.2 Content Validation and Copy

    Write-RunStatus "Copying Reports" -Level Job -Indent 1
    Write-RunStatus "Found $($candidates.Count) filename-matched reports; checking destination collisions..." -Level Activity -Indent 2

    $destinationFolders = @(
        $DestinationPath,
        "$DestinationPath/directories",
        "$DestinationPath/files",
        "$DestinationPath/permissions"
    )
    Write-RunStatus "Checking destination folder readiness..." -Level Activity -Indent 2
    foreach ($destinationFolder in $destinationFolders) {
        $segments = $destinationFolder.Split('/')
        $currentPath = ""
        foreach ($segment in $segments) {
            $currentPath = if ($currentPath) { "$currentPath/$segment" } else { $segment }
            if (-not (Get-AzDataLakeGen2Item -Context $ctx -FileSystem $ContainerName -Path $currentPath -ErrorAction SilentlyContinue)) {
                # ShouldProcess keeps destination creation within the same WhatIf/Confirm contract as report copies.
                if ($PSCmdlet.ShouldProcess("$ContainerName/$currentPath", "Create ADLS Gen2 directory")) {
                    New-AzDataLakeGen2Item `
                        -Context $ctx `
                        -FileSystem $ContainerName `
                        -Path $currentPath `
                        -Directory `
                        -ErrorAction Stop | Out-Null
                }
            }
        }
    }

    $candidateGroups = @($duplicateDestinations.GetEnumerator())
    $copyCandidates = @($candidates | Where-Object { $duplicateDestinations[$_.DestinationPath].Count -eq 1 })
    Write-RunStatus "Reports eligible for copy: $($copyCandidates.Count); duplicate destination names: $($candidates.Count - $copyCandidates.Count)." -Level Activity -Indent 2
    if ($copyCandidates.Count -gt 0 -and -not $WhatIfPreference) {
        $httpClient = [System.Net.Http.HttpClient]::new()
    }

    $groupNumber = 0
    foreach ($group in $candidateGroups) {
        $groupNumber++
        Write-Progress `
            -Id 2 `
            -Activity "Validating and consolidating reports" `
            -Status ("Report {0,5} of {1,5}" -f $groupNumber, $candidateGroups.Count) `
            -PercentComplete (($groupNumber / [math]::Max(1, $candidateGroups.Count)) * 100)

        $groupCandidates = $group.Value
        if ($groupCandidates.Count -gt 1) {
            # Skip every source in a collision group; selecting one would silently lose the other report.
            foreach ($candidate in $groupCandidates) {
                $CopySummary.Add([pscustomobject]@{
                    InputRow = $candidate.InputIndex + 1
                    ReportNaming = $candidate.ReportNaming
                    ReportType = $candidate.ReportType
                    serverUNC = $candidate.serverUNC
                    shareName = $candidate.shareName
                    SourcePath = $candidate.SourcePath
                    DestinationPath = $candidate.DestinationPath
                    RowsChecked = 0
                    InvalidRows = 0
                    Status = "DuplicateDestinationName"
                    Details = "All sources with this destination name were skipped."
                })
            }
            $ReviewItems.Add([pscustomobject]@{
                InputRow = (($groupCandidates | ForEach-Object { $_.InputIndex + 1 }) -join " | ")
                ReportNaming = (($groupCandidates | ForEach-Object { $_.ReportNaming }) -join " | ")
                FullUNCPath = (($groupCandidates | ForEach-Object { $_.InputRow.'Full UNC Path' }) -join " | ")
                serverUNC = (($groupCandidates | ForEach-Object { $_.serverUNC } | Sort-Object -Unique) -join " | ")
                shareName = (($groupCandidates | ForEach-Object { $_.shareName } | Sort-Object -Unique) -join " | ")
                SourcePath = (($groupCandidates | ForEach-Object { $_.SourcePath }) -join " | ")
                DestinationPath = $groupCandidates[0].DestinationPath
                ReportType = $groupCandidates[0].ReportType
                Reason = "Duplicate destination filename"
                Details = "All matching source files were skipped for review."
            })
            continue
        }

        $candidate = $groupCandidates[0]
        $existingDestination = Get-AzDataLakeGen2Item `
            -Context $ctx `
            -FileSystem $ContainerName `
            -Path $candidate.DestinationPath `
            -ErrorAction SilentlyContinue
        if ($existingDestination) {
            $CopySummary.Add([pscustomobject]@{
                InputRow = $candidate.InputIndex + 1
                ReportNaming = $candidate.ReportNaming
                ReportType = $candidate.ReportType
                serverUNC = $candidate.serverUNC
                shareName = $candidate.shareName
                SourcePath = $candidate.SourcePath
                DestinationPath = $candidate.DestinationPath
                RowsChecked = 0
                InvalidRows = 0
                Status = "DestinationExists"
                Details = "Existing destination was not overwritten; source content was not revalidated."
            })
            continue
        }

        if ($WhatIfPreference) {
            $CopySummary.Add([pscustomobject]@{
                InputRow = $candidate.InputIndex + 1
                ReportNaming = $candidate.ReportNaming
                ReportType = $candidate.ReportType
                serverUNC = $candidate.serverUNC
                shareName = $candidate.shareName
                SourcePath = $candidate.SourcePath
                DestinationPath = $candidate.DestinationPath
                RowsChecked = 0
                InvalidRows = 0
                Status = "WhatIfFirstLevelMatch"
                Details = "Filename and source folder matched. Content validation and copy were skipped in WhatIf mode."
            })
            Write-RunStatus "WHATIF: Would copy '$($candidate.SourcePath)' -> '$($candidate.DestinationPath)' (content validation skipped)." -Level Warning -Indent 2
            continue
        }

        $rowsChecked = 0
        $invalidRows = 0
        $sampleIssues = [System.Collections.Generic.List[string]]::new()
        $pathColumn = if ($candidate.ReportType -eq "Perm") { "Path" } else { "FullName" }
        $request = $null
        $response = $null
        $streamReader = $null
        $csvParser = $null

        try {
            # Refresh the Azure token before expiry so a long run can keep reading reports without reconnecting.
            if ($storageTokenExpiresOn -le [DateTimeOffset]::UtcNow.AddMinutes(5)) {
                $storageToken = Get-AzAccessToken -ResourceUrl "https://storage.azure.com/" -ErrorAction Stop
                $storageTokenValue = if ($storageToken.Token -is [System.Security.SecureString]) {
                    [System.Net.NetworkCredential]::new("", $storageToken.Token).Password
                } else {
                    [string]$storageToken.Token
                }
                $storageTokenExpiresOn = $storageToken.ExpiresOn
            }

            $encodedFileSystem = [uri]::EscapeDataString($ContainerName)
            $encodedPath = ($candidate.SourcePath -split '/' | ForEach-Object { [uri]::EscapeDataString($_) }) -join '/'
            $sourceUri = [uri]"https://$StorageAccountName.dfs.core.windows.net/$encodedFileSystem/$encodedPath"
            $request = [System.Net.Http.HttpRequestMessage]::new([System.Net.Http.HttpMethod]::Get, $sourceUri)
            $request.Headers.Authorization = [System.Net.Http.Headers.AuthenticationHeaderValue]::new("Bearer", $storageTokenValue)
            [void]$request.Headers.TryAddWithoutValidation("x-ms-version", "2023-08-03")
            # ResponseHeadersRead avoids buffering the report body; disposing this response after the sample stops further reads.
            $response = $httpClient.SendAsync(
                $request,
                [System.Net.Http.HttpCompletionOption]::ResponseHeadersRead
            ).GetAwaiter().GetResult()
            [void]$response.EnsureSuccessStatusCode()

            $contentStream = $response.Content.ReadAsStreamAsync().GetAwaiter().GetResult()
            $streamReader = [System.IO.StreamReader]::new($contentStream, [System.Text.Encoding]::UTF8, $true, 4096, $true)
            $csvParser = [Microsoft.VisualBasic.FileIO.TextFieldParser]::new($streamReader)
            $csvParser.TextFieldType = [Microsoft.VisualBasic.FileIO.FieldType]::Delimited
            $csvParser.SetDelimiters(',')
            $csvParser.HasFieldsEnclosedInQuotes = $true
            $csvParser.TrimWhiteSpace = $false

            # Resolve only the columns needed for the scope test instead of materializing every CSV row as an object.
            $headers = $csvParser.ReadFields()
            if (-not $headers) {
                throw "Report contains no CSV header."
            }

            $headerIndexes = @{}
            for ($columnIndex = 0; $columnIndex -lt $headers.Count; $columnIndex++) {
                $headerIndexes[[string]$headers[$columnIndex]] = $columnIndex
            }
            foreach ($column in @("ServerUNC", "ShareName", $pathColumn)) {
                if (-not $headerIndexes.ContainsKey($column)) {
                    throw "Report is missing required column '$column'."
                }
            }

            # TextFieldParser handles quoted commas and embedded line breaks; a simple Split(',') would corrupt such rows.
            while ($rowsChecked -lt $ValidationRowLimit -and -not $csvParser.EndOfData) {
                $reportFields = $csvParser.ReadFields()
                if ($reportFields.Count -eq 1 -and [string]::IsNullOrWhiteSpace($reportFields[0])) {
                    continue
                }

                $rowsChecked++
                $rowIssues = [System.Collections.Generic.List[string]]::new()

                $serverIndex = $headerIndexes["ServerUNC"]
                $shareIndex = $headerIndexes["ShareName"]
                $pathIndex = $headerIndexes[$pathColumn]
                $reportServer = if ($serverIndex -lt $reportFields.Count) { [string]$reportFields[$serverIndex] } else { "" }
                $reportShare = if ($shareIndex -lt $reportFields.Count) { [string]$reportFields[$shareIndex] } else { "" }
                $reportPath = if ($pathIndex -lt $reportFields.Count) { [string]$reportFields[$pathIndex] } else { "" }

                if ([string]::IsNullOrWhiteSpace($reportServer) -or
                    -not $reportServer.Trim('\','/').Equals(([string]$candidate.InputRow.serverUNC).Trim('\','/'), [System.StringComparison]::OrdinalIgnoreCase)) {
                    $rowIssues.Add("ServerUNC does not match input serverUNC")
                }
                if (-not (Test-PathWithinRoot -CandidatePath $reportShare -RootPath $candidate.InputRow.shareName)) {
                    $rowIssues.Add("ShareName is not the requested share or a descendant")
                }
                if (-not (Test-PathWithinRoot -CandidatePath $reportPath -RootPath $candidate.InputRow.'Full UNC Path')) {
                    $rowIssues.Add("$pathColumn is not the requested UNC path or a descendant")
                }

                if ($rowIssues.Count -gt 0) {
                    $invalidRows++
                    if ($sampleIssues.Count -lt 5) {
                        $sampleIssues.Add("row ${rowsChecked}: $($rowIssues -join '; ')")
                    }
                }
            }

            if ($rowsChecked -eq 0) {
                throw "Report contains no data rows to validate."
            }
            if ($invalidRows -gt 0) {
                throw "$invalidRows of $rowsChecked rows did not match the requested server/share/path."
            }

            # Encode each path segment independently so spaces and other valid ADLS characters survive the URL.
            $encodedSourcePath = ($candidate.SourcePath -split '/' | ForEach-Object { [uri]::EscapeDataString($_) }) -join '/'
            $encodedDestinationPath = ($candidate.DestinationPath -split '/' | ForEach-Object { [uri]::EscapeDataString($_) }) -join '/'
            $sourceUrl = "https://$StorageAccountName.dfs.core.windows.net/$ContainerName/$encodedSourcePath"
            $destinationUrl = "https://$StorageAccountName.dfs.core.windows.net/$ContainerName/$encodedDestinationPath"
            if ($PSCmdlet.ShouldProcess($candidate.DestinationPath, "Copy validated report from $($candidate.SourcePath)")) {
                # AzCopy runs server-to-server; keep overwrite disabled to guard against destination races.
                $azOutput = & azcopy copy `
                    $sourceUrl `
                    $destinationUrl `
                    --overwrite=false `
                    --log-level=ERROR `
                    --output-type=text `
                    2>&1

                $exitCode = $LASTEXITCODE
                if ($exitCode -ne 0) {
                    throw "AzCopy failed (exit code $exitCode). Output: $azOutput"
                }

                # The destination was checked immediately before AzCopy; a zero exit code means this transfer completed.
                # Do not infer skip status from free-form output, which can mention skipped transfers even when none were skipped.
                $copyStatus = "CopiedAfterSampleValidation"
            }
            else {
                $copyStatus = "WhatIfSampleValidated"
            }

            $CopySummary.Add([pscustomobject]@{
                InputRow = $candidate.InputIndex + 1
                ReportNaming = $candidate.ReportNaming
                ReportType = $candidate.ReportType
                serverUNC = $candidate.serverUNC
                shareName = $candidate.shareName
                SourcePath = $candidate.SourcePath
                DestinationPath = $candidate.DestinationPath
                RowsChecked = $rowsChecked
                InvalidRows = $invalidRows
                Status = $copyStatus
                Details = "Validated the first $rowsChecked data rows against serverUNC, shareName, and Full UNC Path. Remaining rows, if any, were assumed to match."
            })
        }
        catch {
            $CopySummary.Add([pscustomobject]@{
                InputRow = $candidate.InputIndex + 1
                ReportNaming = $candidate.ReportNaming
                ReportType = $candidate.ReportType
                serverUNC = $candidate.serverUNC
                shareName = $candidate.shareName
                SourcePath = $candidate.SourcePath
                DestinationPath = $candidate.DestinationPath
                RowsChecked = $rowsChecked
                InvalidRows = $invalidRows
                Status = "ReviewRequired"
                Details = $_.Exception.Message
            })
            $ReviewItems.Add([pscustomobject]@{
                InputRow = $candidate.InputIndex + 1
                ReportNaming = $candidate.ReportNaming
                FullUNCPath = $candidate.InputRow.'Full UNC Path'
                serverUNC = $candidate.serverUNC
                shareName = $candidate.shareName
                SourcePath = $candidate.SourcePath
                DestinationPath = $candidate.DestinationPath
                ReportType = $candidate.ReportType
                Reason = if ($invalidRows -gt 0) { "Report content did not match input scope" } else { "Report could not be validated or copied" }
                Details = (@($_.Exception.Message) + @($sampleIssues)) -join " | "
            })
        }
        finally {
            if ($csvParser) { $csvParser.Dispose() }
            if ($streamReader) { $streamReader.Dispose() }
            if ($response) { $response.Dispose() }
            if ($request) { $request.Dispose() }
        }
    }
    Write-Progress -Id 2 -Activity "Validating and consolidating reports" -Completed
    $runCompleted = $true

    #endregion 7.2 Content Validation and Copy
}
finally {
    #region 8. Completed Script

    try {
        # Resolve and create the report folder here as well so early setup failures still leave diagnostic CSVs.
        if (-not [System.IO.Path]::IsPathRooted($OutputReportPath)) {
            $OutputReportPath = Join-Path $PSScriptRoot $OutputReportPath
        }
        New-Item -ItemType Directory -Path $OutputReportPath -Force -WhatIf:$false | Out-Null

        if (-not $CopySummary) { $CopySummary = [System.Collections.Generic.List[object]]::new() }
        if (-not $ReviewItems) { $ReviewItems = [System.Collections.Generic.List[object]]::new() }

        $copySummaryPath = Join-Path $OutputReportPath "CopySummary-$reportTimestamp.csv"
        if ($CopySummary.Count -gt 0) {
            $CopySummary | Export-Csv -NoTypeInformation -LiteralPath $copySummaryPath -Force -WhatIf:$false
        }
        else {
            [pscustomobject][ordered]@{
                InputRow = $null; ReportNaming = $null; ReportType = $null; serverUNC = $null; shareName = $null; SourcePath = $null
                DestinationPath = $null; RowsChecked = $null; InvalidRows = $null; Status = $null; Details = $null
            } | ConvertTo-Csv -NoTypeInformation | Select-Object -First 1 | Set-Content -LiteralPath $copySummaryPath -Encoding utf8 -WhatIf:$false
        }

        $reviewItemsPath = Join-Path $OutputReportPath "ReviewItems-$reportTimestamp.csv"
        if ($ReviewItems.Count -gt 0) {
            $ReviewItems | Export-Csv -NoTypeInformation -LiteralPath $reviewItemsPath -Force -WhatIf:$false
        }
        else {
            [pscustomobject][ordered]@{
                InputRow = $null; ReportNaming = $null; FullUNCPath = $null; serverUNC = $null; shareName = $null; SourcePath = $null
                DestinationPath = $null; ReportType = $null; Reason = $null; Details = $null
            } | ConvertTo-Csv -NoTypeInformation | Select-Object -First 1 | Set-Content -LiteralPath $reviewItemsPath -Encoding utf8 -WhatIf:$false
        }
    }
    finally {
        if ($httpClient) { $httpClient.Dispose() }
        if ($transcriptStarted) {
            $runStatus = if ($runCompleted) { "completed" } else { "failed before completion" }
            if (Get-Command Write-RunStatus -CommandType Function -ErrorAction SilentlyContinue) {
                $reportTypeSummary = [System.Collections.Generic.List[object]]::new()
                foreach ($reportType in @(
                    [pscustomobject]@{ Label = "directories"; Token = "Dir" }
                    [pscustomobject]@{ Label = "files"; Token = "File" }
                    [pscustomobject]@{ Label = "permissions"; Token = "Perm" }
                )) {
                    $typeRows = @($CopySummary | Where-Object ReportType -eq $reportType.Token)
                    $typeReviewRows = @($ReviewItems | Where-Object ReportType -eq $reportType.Token)
                    $reportTypeSummary.Add([pscustomobject]@{
                        ReportType = $reportType.Label
                        Copied = @($typeRows | Where-Object Status -eq "CopiedAfterSampleValidation").Count
                        Skipped = @($typeRows | Where-Object Status -in @("DestinationExists", "SkippedDuringCopy", "DuplicateDestinationName")).Count
                        Review = $typeReviewRows.Count
                        WouldCopy = @($typeRows | Where-Object Status -eq "WhatIfFirstLevelMatch").Count
                    })
                }

                $unclassifiedReviewCount = @($ReviewItems | Where-Object ReportType -notin @("Dir", "File", "Perm")).Count
                if ($unclassifiedReviewCount -gt 0) {
                    $reportTypeSummary.Add([pscustomobject]@{
                        ReportType = "unclassified"
                        Copied = 0
                        Skipped = 0
                        Review = $unclassifiedReviewCount
                        WouldCopy = 0
                    })
                }

                Write-RunStatus "Script Completion / Summary" -Level Info
                Write-RunStatus "Report status by type" -Level Job -Indent 1
                $tableText = ($reportTypeSummary | Format-Table -AutoSize | Out-String -Width 120).TrimEnd()
                foreach ($tableLine in ($tableText -split "`r?`n")) {
                    if ([string]::IsNullOrWhiteSpace($tableLine)) { continue }
                    Write-RunStatus $tableLine -Level Activity -Indent 2
                }
                Write-RunStatus "Run outcome" -Level Job -Indent 1
                Write-RunStatus "Consolidation $runStatus." -Level Activity -Indent 2
                Write-RunStatus "Output reports" -Level Job -Indent 1
                Write-RunStatus "Copy summary: $copySummaryPath" -Level Activity -Indent 2
                Write-RunStatus "Review items: $reviewItemsPath" -Level Activity -Indent 2
                Write-RunStatus "Transcript: $(Join-Path $PSScriptRoot "..\_Log Files\Cat.Az-ConsolidateSelectedFileShareReports-$($scriptStart[1]).log")" -Level Activity -Indent 2
            }
            else {
                $summaryTimestamp = Get-Date -Format "yy-MM-dd HH:mm:ss"
                Write-Host "$summaryTimestamp - Consolidation $runStatus before status logging initialized. See transcript and reports." -ForegroundColor $(if ($runCompleted) { "Green" } else { "Red" })
            }
            try {
                if ($scriptStart) { Complete-CatScript $scriptStart[0] }
            }
            finally {
                Stop-Transcript -WhatIf:$false | Out-Null
            }
        }
    }

    #endregion 8. Completed Script
}