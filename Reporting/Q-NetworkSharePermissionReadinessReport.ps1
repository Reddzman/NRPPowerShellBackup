<#
.SYNOPSIS
    Pre-migration network folder permissions and OneDrive readiness assessment with a consolidated Excel report.

.DESCRIPTION
    Reads an input CSV of user network folder paths, scans each path in parallel, and produces:
      - A primary consolidated report with all input CSV columns plus scan/readiness columns
      - A consolidated Excel workbook when ImportExcel is available
      - A fallback Excel-compatible .xls report when ImportExcel is not available
      - Consolidated detail CSVs for troubleshooting
      - Run "Get-Help .\Q-NetworkShareMigrationReadinessReport.ps1 -Examples" for usage examples.

.INPUT CSV FORMAT
    Required column:
      Path

    Optional columns are preserved in the consolidated report:
      User, UserPrincipalName, Wave, Server, Department, Location, MigrationBatch, etc.

.EXAMPLE
    .\Q-NetworkShareMigrationReadinessReport.ps1 `
        -InputCsv "C:\Quisitive\PowerShell\_Input Files\Wave1_UserPaths.csv" `
        -OutputFolder "C:\Quisitive\PowerShell\_Output Files\Wave1_Readiness" `
        -ThrottleLimit 6 `
        -MigrationRootPrefixLength 12

    Performs a standard migration readiness assessment for all paths listed in
    Wave1_UserPaths.csv. The scan runs with six parallel workers and adds
    12 characters to the effective path length calculation to account for the
    target migration root folder structure. Generates a consolidated readiness
    report, dashboard, and detailed issue reports.

.EXAMPLE
    .\Q-NetworkShareMigrationReadinessReport.ps1 `
        -InputCsv "C:\Quisitive\PowerShell\_Input Files\Wave1_UserPaths.csv" `
        -OutputFolder "C:\Quisitive\PowerShell\_Output Files\Wave1_Readiness" `
        -ThrottleLimit 6 `
        -ExportAclOnFailures `
        -KeepTempFiles

    Performs a migration readiness assessment and exports detailed NTFS ACL
    information whenever access-denied or permission-related errors are
    encountered. Temporary scan files are preserved for troubleshooting,
    allowing administrators to investigate permission issues, inaccessible
    folders, or unexpected scan failures after the scan completes.


.EXAMPLE
    .\Q-NetworkShareMigrationReadinessReport.ps1 `
        -InputCsv "C:\Quisitive\PowerShell\_Input Files\Wave1_UserPaths.csv" `
        -OutputFolder "C:\Quisitive\PowerShell\_Output Files\Wave1_Readiness" `
        -ExportInventory

    Performs a readiness assessment and exports a complete inventory of all
    files and folders discovered during the scan. Useful for migration sizing,
    content analysis, and validation before migration planning.

.EXAMPLE
    .\Q-NetworkShareMigrationReadinessReport.ps1 `
        -InputCsv "C:\Quisitive\PowerShell\_Input Files\Wave1_UserPaths.csv" `
        -OutputFolder "C:\Quisitive\PowerShell\_Output Files\Wave1_Readiness" `
        -FlagHashAndPercent

    Scans for files and folders that contain hash (#) or percent (%) characters
    in their names in addition to standard OneDrive and SharePoint naming
    restrictions. Useful when validating content for modern Microsoft 365
    migration compatibility.

.EXAMPLE
    .\Q-NetworkShareMigrationReadinessReport.ps1 `
        -InputCsv "C:\Quisitive\PowerShell\_Input Files\Wave1_UserPaths.csv" `
        -OutputFolder "C:\Quisitive\PowerShell\_Output Files\Wave1_Readiness" `
        -ThrottleLimit 12 `
        -InstallImportExcel

    Performs a high-speed assessment using 12 parallel workers and
    automatically attempts to install the ImportExcel module if needed.
    Generates a multi-sheet .xlsx workbook containing executive dashboards,
    user readiness summaries, and detailed issue reports.

.NOTES
    For a true .xlsx workbook, install ImportExcel:
      Install-Module ImportExcel -Scope CurrentUser

    If ImportExcel is not available, this script still creates:
      - Consolidated_User_Readiness.csv
      - Migration_Readiness_Report.xls    # Excel-compatible HTML fallback
#>

#requires -Version 7.0

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$InputCsv,

    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$OutputFolder,

    [ValidateRange(1, 64)]
    [int]$ThrottleLimit = 6,

    [switch]$ExportInventory,
    [switch]$ExportAclOnFailures,
    [switch]$SkipFileReadCheck,

    [ValidateRange(1, 1000)]
    [int]$OneDrivePathLimit = 400,

    [ValidateRange(0, 500)]
    [int]$MigrationRootPrefixLength = 0,

    [ValidateRange(1, 1048576)]
    [int]$LargeFileMB = 250,

    [ValidateRange(1, 1048576)]
    [int]$VeryLargeFileMB = 15360,

    [ValidateRange(1, 1024)]
    [int]$MaxOneDriveFileSizeGB = 250,

    [ValidateRange(1, 10000)]
    [double]$EstimatedThroughputGBPerHour = 50,

    [ValidateRange(1, 120)]
    [int]$ProgressRefreshSeconds = 5,

    [switch]$FlagHashAndPercent,
    [switch]$InstallImportExcel,
    [switch]$KeepTempFiles
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
$ScriptStart = Get-Date

if (-not (Test-Path -LiteralPath $InputCsv)) {
    throw "Input CSV not found: $InputCsv"
}

if (-not (Test-Path -LiteralPath $OutputFolder)) {
    New-Item -Path $OutputFolder -ItemType Directory -Force | Out-Null
}

$OutputFolder = (Resolve-Path -LiteralPath $OutputFolder).Path
$TempFolder = Join-Path $OutputFolder "_TempScanParts"
$ProgressFile = Join-Path $OutputFolder "_ScanProgress.csv"
$LogPath = Join-Path $OutputFolder "Scan_Log.txt"

foreach ($p in @($TempFolder)) {
    if (Test-Path -LiteralPath $p) { Remove-Item -LiteralPath $p -Recurse -Force }
    New-Item -Path $p -ItemType Directory -Force | Out-Null
}

foreach ($p in @($ProgressFile,$LogPath)) {
    if (Test-Path -LiteralPath $p) { Remove-Item -LiteralPath $p -Force }
}

"Timestamp,Status,Index,RootPath,FilesScanned,FoldersScanned,AccessFailures,Message" | Out-File -LiteralPath $ProgressFile -Encoding UTF8

function Write-Log {
    param(
        [Parameter(Mandatory = $true)][string]$Message,
        [string]$Level = "INFO"
    )
    $line = "[{0}] [{1}] {2}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Level, $Message
    Add-Content -LiteralPath $LogPath -Value $line -Encoding UTF8
    Write-Host $line
}

function ConvertTo-HtmlTable {
    param(
        [Parameter(Mandatory = $true)][object[]]$Data,
        [Parameter(Mandatory = $true)][string]$Title
    )

    $html = New-Object System.Text.StringBuilder
    [void]$html.AppendLine("<h2>$Title</h2>")
    [void]$html.AppendLine("<table border='1' cellspacing='0' cellpadding='3'>")

    $rows = @($Data)
    if (@($rows).Count -eq 0) {
        [void]$html.AppendLine("<tr><td>No data</td></tr>")
        [void]$html.AppendLine("</table>")
        return $html.ToString()
    }

    $headers = @($rows[0].PSObject.Properties.Name)
    [void]$html.AppendLine("<tr>" + (($headers | ForEach-Object { "<th>$($_)</th>" }) -join "") + "</tr>")
    foreach ($row in $rows) {
        $cells = foreach ($h in $headers) {
            $value = [System.Net.WebUtility]::HtmlEncode([string]$row.$h)
            "<td>$value</td>"
        }
        [void]$html.AppendLine("<tr>$($cells -join '')</tr>")
    }
    [void]$html.AppendLine("</table>")
    return $html.ToString()
}

function Merge-CsvParts {
    param(
        [Parameter(Mandatory = $true)][string]$TempFolder,
        [Parameter(Mandatory = $true)][string]$FileName,
        [Parameter(Mandatory = $true)][string]$DestinationPath,
        [bool]$Required = $true
    )

    $parts = @(Get-ChildItem -Path $TempFolder -Recurse -Filter $FileName -File | Sort-Object FullName)
    if (@($parts).Count -eq 0) {
        if ($Required) { Write-Log "No output parts found for $FileName" "WARN" }
        return $false
    }

    $utf8Bom = [System.Text.UTF8Encoding]::new($true)
    $writer = [System.IO.StreamWriter]::new($DestinationPath, $false, $utf8Bom)
    try {
        $isFirstFile = $true
        foreach ($part in @($parts)) {
            $lineNumber = 0
            foreach ($line in [System.IO.File]::ReadLines($part.FullName)) {
                $lineNumber++
                if ($lineNumber -eq 1 -and -not $isFirstFile) { continue }
                $writer.WriteLine($line)
            }
            $isFirstFile = $false
        }
    }
    finally {
        $writer.Flush(); $writer.Close(); $writer.Dispose()
    }
    return $true
}

function New-ReadinessExcelReport {
    param(
        [Parameter(Mandatory = $true)][string]$OutputFolder,
        [Parameter(Mandatory = $true)][object[]]$DashboardRows,
        [Parameter(Mandatory = $true)][object[]]$UserSummaryRows,
        [string]$PermissionFailuresCsv,
        [string]$LongPathCsv,
        [string]$NamingIssuesCsv,
        [string]$FatalErrorsCsv,
        [switch]$InstallImportExcel
    )

    $xlsxPath = Join-Path $OutputFolder "Migration_Readiness_Report.xlsx"
    $xlsFallbackPath = Join-Path $OutputFolder "Migration_Readiness_Report.xls"

    $importExcelAvailable = $false
    if ($InstallImportExcel -and -not (Get-Module -ListAvailable -Name ImportExcel)) {
        try {
            Write-Log "Attempting to install ImportExcel module for .xlsx output."
            Install-Module ImportExcel -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
        }
        catch {
            Write-Log "ImportExcel installation failed or is unavailable: $($_.Exception.Message)" "WARN"
        }
    }

    try {
        Import-Module ImportExcel -ErrorAction Stop
        $importExcelAvailable = $true
    }
    catch {
        $importExcelAvailable = $false
    }

    if ($importExcelAvailable) {
        if (Test-Path -LiteralPath $xlsxPath) { Remove-Item -LiteralPath $xlsxPath -Force }

        $DashboardRows | Export-Excel -Path $xlsxPath -WorksheetName "Dashboard" -AutoSize -BoldTopRow -FreezeTopRow -TableName "Dashboard" -ClearSheet
        $UserSummaryRows | Export-Excel -Path $xlsxPath -WorksheetName "User Summary" -AutoSize -BoldTopRow -FreezeTopRow -TableName "UserSummary" -Append

        if ($PermissionFailuresCsv -and (Test-Path -LiteralPath $PermissionFailuresCsv)) {
            @(Import-Csv -LiteralPath $PermissionFailuresCsv) | Export-Excel -Path $xlsxPath -WorksheetName "Permission Failures" -AutoSize -BoldTopRow -FreezeTopRow -TableName "PermissionFailures" -Append
        }
        if ($LongPathCsv -and (Test-Path -LiteralPath $LongPathCsv)) {
            @(Import-Csv -LiteralPath $LongPathCsv) | Export-Excel -Path $xlsxPath -WorksheetName "Long Paths" -AutoSize -BoldTopRow -FreezeTopRow -TableName "LongPaths" -Append
        }
        if ($NamingIssuesCsv -and (Test-Path -LiteralPath $NamingIssuesCsv)) {
            @(Import-Csv -LiteralPath $NamingIssuesCsv) | Export-Excel -Path $xlsxPath -WorksheetName "Naming Issues" -AutoSize -BoldTopRow -FreezeTopRow -TableName "NamingIssues" -Append
        }
        if ($FatalErrorsCsv -and (Test-Path -LiteralPath $FatalErrorsCsv)) {
            @(Import-Csv -LiteralPath $FatalErrorsCsv) | Export-Excel -Path $xlsxPath -WorksheetName "Fatal Errors" -AutoSize -BoldTopRow -FreezeTopRow -TableName "FatalErrors" -Append
        }

        return [PSCustomObject]@{ Path = $xlsxPath; Format = "xlsx"; ImportExcelUsed = $true }
    }

    Write-Log "ImportExcel module not found. Creating Excel-compatible .xls fallback report." "WARN"

    $html = New-Object System.Text.StringBuilder
    [void]$html.AppendLine("<html><head><meta charset='utf-8'><style>body{font-family:Segoe UI,Arial;} table{border-collapse:collapse;margin-bottom:24px;} th{background:#1f4e78;color:white;} td,th{border:1px solid #999;padding:4px;} h1{color:#1f4e78;}</style></head><body>")
    [void]$html.AppendLine("<h1>Migration Readiness Report</h1>")
    [void]$html.AppendLine((ConvertTo-HtmlTable -Data $DashboardRows -Title "Dashboard"))
    [void]$html.AppendLine((ConvertTo-HtmlTable -Data $UserSummaryRows -Title "User Summary"))

    if ($PermissionFailuresCsv -and (Test-Path -LiteralPath $PermissionFailuresCsv)) { [void]$html.AppendLine((ConvertTo-HtmlTable -Data @(Import-Csv -LiteralPath $PermissionFailuresCsv) -Title "Permission Failures")) }
    if ($LongPathCsv -and (Test-Path -LiteralPath $LongPathCsv)) { [void]$html.AppendLine((ConvertTo-HtmlTable -Data @(Import-Csv -LiteralPath $LongPathCsv) -Title "Long Paths")) }
    if ($NamingIssuesCsv -and (Test-Path -LiteralPath $NamingIssuesCsv)) { [void]$html.AppendLine((ConvertTo-HtmlTable -Data @(Import-Csv -LiteralPath $NamingIssuesCsv) -Title "Naming Issues")) }
    if ($FatalErrorsCsv -and (Test-Path -LiteralPath $FatalErrorsCsv)) { [void]$html.AppendLine((ConvertTo-HtmlTable -Data @(Import-Csv -LiteralPath $FatalErrorsCsv) -Title "Fatal Errors")) }

    [void]$html.AppendLine("</body></html>")
    $html.ToString() | Out-File -LiteralPath $xlsFallbackPath -Encoding UTF8
    return [PSCustomObject]@{ Path = $xlsFallbackPath; Format = "xls-html-fallback"; ImportExcelUsed = $false }
}

Write-Host ""
Write-Host "========== Migration Readiness Assessment ==========" -ForegroundColor Cyan
Write-Host "Input CSV     : $InputCsv"
Write-Host "Output Folder : $OutputFolder"
Write-Host "ThrottleLimit : $ThrottleLimit"
Write-Host "====================================================" -ForegroundColor Cyan
Write-Host ""

Write-Log "Loading input CSV."
$InputRows = @(Import-Csv -LiteralPath $InputCsv)
if (@($InputRows).Count -eq 0) { throw "Input CSV contains no rows." }

$InputColumns = @($InputRows[0].PSObject.Properties.Name)
if ($InputColumns -notcontains "Path") { throw "Input CSV must contain a column named 'Path'." }

Write-Log "Preparing scan rows."
$PreparedRows = [System.Collections.Generic.List[object]]::new()
$rowIndex = 0
foreach ($row in @($InputRows)) {
    $rowIndex++
    $pathValue = [string]$row.Path
    if ([string]::IsNullOrWhiteSpace($pathValue)) {
        Write-Log "Skipping input row $rowIndex because Path is blank." "WARN"
        continue
    }

    # Store the original row as compressed JSON so every input column is preserved later.
    $inputJson = ($row | ConvertTo-Json -Compress -Depth 5)
    [void]$PreparedRows.Add([PSCustomObject]@{
        Index     = $rowIndex
        Path      = $pathValue.Trim()
        InputJson = $inputJson
    })
}

if (@($PreparedRows).Count -eq 0) { throw "No usable paths were found in the input CSV." }
$PreparedRowsArray = @($PreparedRows.ToArray())
$TotalRoots = @($PreparedRowsArray).Count

Write-Log "Prepared root path count: $TotalRoots"
Write-Host "Step 1/5: Prepared $TotalRoots root paths for scanning." -ForegroundColor Green

$MaxOneDriveFileSizeBytes = [int64]$MaxOneDriveFileSizeGB * 1GB
$LargeFileBytes = [int64]$LargeFileMB * 1MB
$VeryLargeFileBytes = [int64]$VeryLargeFileMB * 1MB

$ExportInventoryValue = [bool]$ExportInventory
$ExportAclOnFailuresValue = [bool]$ExportAclOnFailures
$SkipFileReadCheckValue = [bool]$SkipFileReadCheck
$FlagHashAndPercentValue = [bool]$FlagHashAndPercent

Write-Host "Step 2/5: Starting parallel scan. Progress will update every $ProgressRefreshSeconds seconds." -ForegroundColor Green
Write-Log "Starting parallel scan jobs."

$scanJob = $PreparedRowsArray | ForEach-Object -Parallel {
    $ScanRow = $_

    $TempFolder = $using:TempFolder
    $ProgressFile = $using:ProgressFile
    $ExportInventory = [bool]$using:ExportInventoryValue
    $ExportAclOnFailures = [bool]$using:ExportAclOnFailuresValue
    $SkipFileReadCheck = [bool]$using:SkipFileReadCheckValue
    $OneDrivePathLimit = [int]$using:OneDrivePathLimit
    $MigrationRootPrefixLength = [int]$using:MigrationRootPrefixLength
    $LargeFileBytes = [int64]$using:LargeFileBytes
    $VeryLargeFileBytes = [int64]$using:VeryLargeFileBytes
    $MaxOneDriveFileSizeBytes = [int64]$using:MaxOneDriveFileSizeBytes
    $FlagHashAndPercent = [bool]$using:FlagHashAndPercentValue

    Set-StrictMode -Version Latest
    $ErrorActionPreference = "Stop"

    function Append-ProgressLine {
        param([string]$Line)
        for ($i = 1; $i -le 10; $i++) {
            try {
                [System.IO.File]::AppendAllText($ProgressFile, $Line + [Environment]::NewLine, [System.Text.Encoding]::UTF8)
                break
            }
            catch {
                Start-Sleep -Milliseconds 100
            }
        }
    }

    function Escape-CsvValue {
        param([object]$Value)
        if ($null -eq $Value) { return '""' }
        $s = [string]$Value
        $s = $s -replace '"', '""'
        return '"' + $s + '"'
    }

    function Write-CsvRow {
        param([System.IO.StreamWriter]$Writer,[object[]]$Values)
        $escaped = foreach ($v in @($Values)) { Escape-CsvValue -Value $v }
        $Writer.WriteLine(($escaped -join ','))
    }

    function New-SafeFileName {
        param([string]$Value)
        $safe = $Value -replace '[\\/:*?"<>|]', '_'
        $safe = $safe -replace '\s+', '_'
        if ($safe.Length -gt 120) { $safe = $safe.Substring(0, 120) }
        return $safe
    }

    function Test-InvalidOneDriveName {
        param([string]$Name,[bool]$FlagHashAndPercent)
        $issues = [System.Collections.Generic.List[string]]::new()
        if ([string]::IsNullOrWhiteSpace($Name)) { [void]$issues.Add("BlankName"); return @($issues.ToArray()) }
        [char[]]$invalidChars = @('"','*',':','<','>','?','/','\','|')
        if ($Name.IndexOfAny($invalidChars) -ge 0) { [void]$issues.Add("InvalidCharacter") }
        if ($FlagHashAndPercent -and ($Name.Contains('#') -or $Name.Contains('%'))) { [void]$issues.Add("HashOrPercentReview") }
        if ($Name.StartsWith(' ')) { [void]$issues.Add("LeadingSpace") }
        if ($Name.EndsWith(' ')) { [void]$issues.Add("TrailingSpace") }
        if ($Name.EndsWith('.')) { [void]$issues.Add("TrailingPeriod") }
        if ($Name.StartsWith('~$')) { [void]$issues.Add("StartsWithTildeDollar") }
        if ($Name -match '_vti_') { [void]$issues.Add("Contains_vti") }
        $lowerName = $Name.ToLowerInvariant()
        if ($lowerName -eq ".lock" -or $lowerName -eq "desktop.ini") { [void]$issues.Add("ReservedName") }
        $baseName = $Name
        if ($Name.Contains('.')) { $baseName = ($Name -split '\.')[0] }
        $reservedBaseNames = @("CON","PRN","AUX","NUL","COM0","COM1","COM2","COM3","COM4","COM5","COM6","COM7","COM8","COM9","LPT0","LPT1","LPT2","LPT3","LPT4","LPT5","LPT6","LPT7","LPT8","LPT9")
        if ($reservedBaseNames -contains $baseName.ToUpperInvariant()) { [void]$issues.Add("ReservedDeviceName") }
        return @($issues.ToArray())
    }

    function Get-OwnerSafe {
        param([string]$Path)
        try { return [string](Get-Acl -LiteralPath $Path -ErrorAction Stop).Owner }
        catch { return "" }
    }

    function Write-AclRowsSafe {
        param([string]$RootPath,[string]$ItemPath,[System.IO.StreamWriter]$Writer)
        try {
            $acl = Get-Acl -LiteralPath $ItemPath -ErrorAction Stop
            foreach ($ace in @($acl.Access)) {
                Write-CsvRow -Writer $Writer -Values @((Get-Date -Format "yyyy-MM-dd HH:mm:ss"),$RootPath,$ItemPath,[string]$acl.Owner,[string]$ace.IdentityReference,[string]$ace.AccessControlType,[string]$ace.FileSystemRights,[string]$ace.IsInherited,[string]$ace.InheritanceFlags,[string]$ace.PropagationFlags,"")
            }
        }
        catch {
            Write-CsvRow -Writer $Writer -Values @((Get-Date -Format "yyyy-MM-dd HH:mm:ss"),$RootPath,$ItemPath,"","","","","","","","ACLReadFailed: $($_.Exception.Message)")
        }
    }

    $rootPath = [string]$ScanRow.Path
    $rootIndex = [int]$ScanRow.Index
    $scanStart = Get-Date
    Append-ProgressLine ('"{0}","Started","{1}","{2}","0","0","0","Scan started"' -f $scanStart.ToString("yyyy-MM-dd HH:mm:ss"),$rootIndex,($rootPath -replace '"','""'))

    $partName = "{0:D6}_{1}" -f $rootIndex,(New-SafeFileName -Value $rootPath)
    $partFolder = Join-Path $TempFolder $partName
    New-Item -Path $partFolder -ItemType Directory -Force | Out-Null

    $utf8Bom = [System.Text.UTF8Encoding]::new($true)
    $permissionFailuresPath = Join-Path $partFolder "Permission_Failures.csv"
    $rootErrorsPath = Join-Path $partFolder "Root_Path_Errors.csv"
    $inventoryPath = Join-Path $partFolder "Inventory.csv"
    $permissionsPath = Join-Path $partFolder "Permissions.csv"
    $longPathPath = Join-Path $partFolder "Long_Path_Report.csv"
    $namingIssuesPath = Join-Path $partFolder "Naming_Issues.csv"
    $fatalErrorsPath = Join-Path $partFolder "Fatal_Scan_Errors.csv"

    $pfWriter = [System.IO.StreamWriter]::new($permissionFailuresPath,$false,$utf8Bom)
    $reWriter = [System.IO.StreamWriter]::new($rootErrorsPath,$false,$utf8Bom)
    $lpWriter = [System.IO.StreamWriter]::new($longPathPath,$false,$utf8Bom)
    $niWriter = [System.IO.StreamWriter]::new($namingIssuesPath,$false,$utf8Bom)
    $feWriter = [System.IO.StreamWriter]::new($fatalErrorsPath,$false,$utf8Bom)
    $invWriter = $null
    $aclWriter = $null
    if ($ExportInventory) { $invWriter = [System.IO.StreamWriter]::new($inventoryPath,$false,$utf8Bom) }
    if ($ExportAclOnFailures) { $aclWriter = [System.IO.StreamWriter]::new($permissionsPath,$false,$utf8Bom) }

    try {
        Write-CsvRow -Writer $pfWriter -Values @("ScanTime","RootPath","ItemPath","ItemName","ItemType","ErrorType","ErrorMessage","Owner","SizeMB","LastModified","PathLength","EffectivePathLength")
        Write-CsvRow -Writer $reWriter -Values @("ScanTime","RootPath","RootAccessible","ErrorType","ErrorMessage")
        Write-CsvRow -Writer $lpWriter -Values @("ScanTime","RootPath","ItemPath","ItemType","PathLength","EffectivePathLength","Limit")
        Write-CsvRow -Writer $niWriter -Values @("ScanTime","RootPath","ItemPath","ItemName","ItemType","Issue")
        Write-CsvRow -Writer $feWriter -Values @("ScanTime","RootPath","ErrorType","ErrorMessage","ScriptStackTrace")
        if ($ExportInventory -and $null -ne $invWriter) { Write-CsvRow -Writer $invWriter -Values @("ScanTime","RootPath","ItemPath","ItemName","ItemType","SizeMB","LastModified","PathLength","EffectivePathLength") }
        if ($ExportAclOnFailures -and $null -ne $aclWriter) { Write-CsvRow -Writer $aclWriter -Values @("ScanTime","RootPath","ItemPath","Owner","Identity","AccessType","Rights","Inherited","InheritanceFlags","PropagationFlags","AclError") }

        $filesScanned = [int64]0; $foldersScanned = [int64]0; $totalBytes = [int64]0
        $accessFailures = [int64]0; $rootPathErrors = [int64]0; $longPathCount = [int64]0; $invalidNameCount = [int64]0
        $largeFileCount = [int64]0; $veryLargeFileCount = [int64]0; $overMaxFileSizeCount = [int64]0
        $largestFileBytes = [int64]0
        $rootAccessible = $true

        try {
            if (-not (Test-Path -LiteralPath $rootPath)) { throw [System.IO.DirectoryNotFoundException]::new("Root path not found or not visible to the executing account.") }
            [void][System.IO.Directory]::GetFileSystemEntries($rootPath)
        }
        catch {
            $rootAccessible = $false; $rootPathErrors++; $accessFailures++
            Write-CsvRow -Writer $reWriter -Values @((Get-Date -Format "yyyy-MM-dd HH:mm:ss"),$rootPath,"False",$_.Exception.GetType().FullName,$_.Exception.Message)
            Write-CsvRow -Writer $pfWriter -Values @((Get-Date -Format "yyyy-MM-dd HH:mm:ss"),$rootPath,$rootPath,(Split-Path -Path $rootPath -Leaf),"RootFolder",$_.Exception.GetType().FullName,$_.Exception.Message,"","","",$rootPath.Length,($rootPath.Length + $MigrationRootPrefixLength))
        }

        if ($rootAccessible) {
            $stack = [System.Collections.Generic.Stack[string]]::new()
            $stack.Push($rootPath)
            while ($stack.Count -gt 0) {
                $currentFolder = $stack.Pop()
                try { $children = @([System.IO.Directory]::GetFileSystemEntries($currentFolder)) }
                catch {
                    $accessFailures++
                    $owner = if ($ExportAclOnFailures) { Get-OwnerSafe -Path $currentFolder } else { "" }
                    Write-CsvRow -Writer $pfWriter -Values @((Get-Date -Format "yyyy-MM-dd HH:mm:ss"),$rootPath,$currentFolder,(Split-Path -Path $currentFolder -Leaf),"Folder",$_.Exception.GetType().FullName,$_.Exception.Message,$owner,"","",$currentFolder.Length,($currentFolder.Length + $MigrationRootPrefixLength))
                    if ($ExportAclOnFailures -and $null -ne $aclWriter) { Write-AclRowsSafe -RootPath $rootPath -ItemPath $currentFolder -Writer $aclWriter }
                    continue
                }

                foreach ($child in @($children)) {
                    $itemName = [System.IO.Path]::GetFileName($child)
                    $pathLength = $child.Length
                    $effectivePathLength = $pathLength + $MigrationRootPrefixLength
                    $lastModified = ""; $sizeMB = ""; $isDirectory = $false

                    try {
                        $attributes = [System.IO.File]::GetAttributes($child)
                        $isDirectory = (($attributes -band [System.IO.FileAttributes]::Directory) -eq [System.IO.FileAttributes]::Directory)
                    }
                    catch {
                        $accessFailures++
                        Write-CsvRow -Writer $pfWriter -Values @((Get-Date -Format "yyyy-MM-dd HH:mm:ss"),$rootPath,$child,$itemName,"Unknown",$_.Exception.GetType().FullName,$_.Exception.Message,"","","",$pathLength,$effectivePathLength)
                        if ($ExportAclOnFailures -and $null -ne $aclWriter) { Write-AclRowsSafe -RootPath $rootPath -ItemPath $child -Writer $aclWriter }
                        continue
                    }

                    if ($isDirectory) {
                        $foldersScanned++
                        try { $lastModified = ([System.IO.DirectoryInfo]::new($child)).LastWriteTime.ToString("yyyy-MM-dd HH:mm:ss") } catch { $lastModified = "" }

                        if ($effectivePathLength -gt $OneDrivePathLimit) {
                            $longPathCount++
                            Write-CsvRow -Writer $lpWriter -Values @((Get-Date -Format "yyyy-MM-dd HH:mm:ss"),$rootPath,$child,"Folder",$pathLength,$effectivePathLength,$OneDrivePathLimit)
                        }

                        $nameIssues = @(Test-InvalidOneDriveName -Name $itemName -FlagHashAndPercent $FlagHashAndPercent)
                        if (@($nameIssues).Count -gt 0) {
                            $invalidNameCount++
                            foreach ($issue in @($nameIssues)) { Write-CsvRow -Writer $niWriter -Values @((Get-Date -Format "yyyy-MM-dd HH:mm:ss"),$rootPath,$child,$itemName,"Folder",$issue) }
                        }

                        if ($ExportInventory -and $null -ne $invWriter) { Write-CsvRow -Writer $invWriter -Values @((Get-Date -Format "yyyy-MM-dd HH:mm:ss"),$rootPath,$child,$itemName,"Folder","",$lastModified,$pathLength,$effectivePathLength) }

                        try { [void][System.IO.Directory]::GetFileSystemEntries($child); $stack.Push($child) }
                        catch {
                            $accessFailures++
                            $owner = if ($ExportAclOnFailures) { Get-OwnerSafe -Path $child } else { "" }
                            Write-CsvRow -Writer $pfWriter -Values @((Get-Date -Format "yyyy-MM-dd HH:mm:ss"),$rootPath,$child,$itemName,"Folder",$_.Exception.GetType().FullName,$_.Exception.Message,$owner,"",$lastModified,$pathLength,$effectivePathLength)
                            if ($ExportAclOnFailures -and $null -ne $aclWriter) { Write-AclRowsSafe -RootPath $rootPath -ItemPath $child -Writer $aclWriter }
                        }
                    }
                    else {
                        $filesScanned++
                        try {
                            $fileInfo = [System.IO.FileInfo]::new($child)
                            $totalBytes += $fileInfo.Length
                            if ($fileInfo.Length -gt $largestFileBytes) { $largestFileBytes = $fileInfo.Length }
                            $sizeMB = [math]::Round(($fileInfo.Length / 1MB), 3)
                            $lastModified = $fileInfo.LastWriteTime.ToString("yyyy-MM-dd HH:mm:ss")
                            if ($fileInfo.Length -ge $LargeFileBytes) { $largeFileCount++ }
                            if ($fileInfo.Length -ge $VeryLargeFileBytes) { $veryLargeFileCount++ }
                            if ($fileInfo.Length -gt $MaxOneDriveFileSizeBytes) { $overMaxFileSizeCount++ }
                        }
                        catch {
                            $accessFailures++
                            Write-CsvRow -Writer $pfWriter -Values @((Get-Date -Format "yyyy-MM-dd HH:mm:ss"),$rootPath,$child,$itemName,"File",$_.Exception.GetType().FullName,$_.Exception.Message,"","","",$pathLength,$effectivePathLength)
                            if ($ExportAclOnFailures -and $null -ne $aclWriter) { Write-AclRowsSafe -RootPath $rootPath -ItemPath $child -Writer $aclWriter }
                            continue
                        }

                        if (-not $SkipFileReadCheck) {
                            try {
                                $stream = [System.IO.File]::Open($child,[System.IO.FileMode]::Open,[System.IO.FileAccess]::Read,([System.IO.FileShare]::ReadWrite -bor [System.IO.FileShare]::Delete))
                                $stream.Close(); $stream.Dispose()
                            }
                            catch {
                                $accessFailures++
                                $owner = if ($ExportAclOnFailures) { Get-OwnerSafe -Path $child } else { "" }
                                Write-CsvRow -Writer $pfWriter -Values @((Get-Date -Format "yyyy-MM-dd HH:mm:ss"),$rootPath,$child,$itemName,"File",$_.Exception.GetType().FullName,$_.Exception.Message,$owner,$sizeMB,$lastModified,$pathLength,$effectivePathLength)
                                if ($ExportAclOnFailures -and $null -ne $aclWriter) { Write-AclRowsSafe -RootPath $rootPath -ItemPath $child -Writer $aclWriter }
                            }
                        }

                        if ($effectivePathLength -gt $OneDrivePathLimit) {
                            $longPathCount++
                            Write-CsvRow -Writer $lpWriter -Values @((Get-Date -Format "yyyy-MM-dd HH:mm:ss"),$rootPath,$child,"File",$pathLength,$effectivePathLength,$OneDrivePathLimit)
                        }

                        $nameIssues = @(Test-InvalidOneDriveName -Name $itemName -FlagHashAndPercent $FlagHashAndPercent)
                        if (@($nameIssues).Count -gt 0) {
                            $invalidNameCount++
                            foreach ($issue in @($nameIssues)) { Write-CsvRow -Writer $niWriter -Values @((Get-Date -Format "yyyy-MM-dd HH:mm:ss"),$rootPath,$child,$itemName,"File",$issue) }
                        }

                        if ($ExportInventory -and $null -ne $invWriter) { Write-CsvRow -Writer $invWriter -Values @((Get-Date -Format "yyyy-MM-dd HH:mm:ss"),$rootPath,$child,$itemName,"File",$sizeMB,$lastModified,$pathLength,$effectivePathLength) }
                    }
                }
            }
        }

        $scanEnd = Get-Date
        $durationSeconds = [math]::Round(($scanEnd - $scanStart).TotalSeconds,2)
        $totalItems = $filesScanned + $foldersScanned
        $totalSizeGB = [math]::Round(($totalBytes / 1GB),3)
        $largestFileMB = [math]::Round(($largestFileBytes / 1MB),3)

        $impactNotes = [System.Collections.Generic.List[string]]::new()
        $impactLevel = "Low"
        if ($accessFailures -gt 0) { [void]$impactNotes.Add("Access denied or inaccessible items detected"); $impactLevel = "High" }
        if ($overMaxFileSizeCount -gt 0) { [void]$impactNotes.Add("Files exceed configured OneDrive maximum file size threshold"); $impactLevel = "High" }
        if ($longPathCount -gt 0) { [void]$impactNotes.Add("Long path items detected"); if ($impactLevel -ne "High") { $impactLevel = "Medium" } }
        if ($invalidNameCount -gt 0) { [void]$impactNotes.Add("Invalid OneDrive or SharePoint names detected"); if ($impactLevel -ne "High") { $impactLevel = "Medium" } }
        if ($veryLargeFileCount -gt 0) { [void]$impactNotes.Add("Very large files detected"); if ($impactLevel -ne "High") { $impactLevel = "Medium" } }
        if (@($impactNotes).Count -eq 0) { [void]$impactNotes.Add("No major OneDrive migration impact indicators detected") }

        $permissionStatus = if ($accessFailures -gt 0 -or $rootPathErrors -gt 0) { "FAIL" } else { "PASS" }
        $readiness = "Green"
        $readinessReason = "No major readiness issues detected"
        if ($permissionStatus -eq "FAIL" -or $overMaxFileSizeCount -gt 0) {
            $readiness = "Red"
            $readinessReason = "Access failures, root path errors, or max file size issues detected"
        }
        elseif ($longPathCount -gt 0 -or $invalidNameCount -gt 0 -or $veryLargeFileCount -gt 0) {
            $readiness = "Yellow"
            $readinessReason = "Migration risk indicators detected but no access-blocking issue found"
        }

        Append-ProgressLine ('"{0}","Completed","{1}","{2}","{3}","{4}","{5}","{6}"' -f $scanEnd.ToString("yyyy-MM-dd HH:mm:ss"),$rootIndex,($rootPath -replace '"','""'),$filesScanned,$foldersScanned,$accessFailures,"Scan completed")

        [PSCustomObject]@{
            Index                = $rootIndex
            Path                 = $rootPath
            ScanDate             = $scanEnd.ToString("yyyy-MM-dd")
            ScanStart            = $scanStart.ToString("yyyy-MM-dd HH:mm:ss")
            ScanEnd              = $scanEnd.ToString("yyyy-MM-dd HH:mm:ss")
            ScanDurationSeconds  = $durationSeconds
            RootAccessible       = [string]$rootAccessible
            PermissionStatus     = $permissionStatus
            MigrationReady       = if ($readiness -eq "Green") { "Yes" } else { "No" }
            Readiness            = $readiness
            ReadinessReason      = $readinessReason
            ImpactLevel          = $impactLevel
            ImpactNotes          = (@($impactNotes) -join "; ")
            FilesScanned         = $filesScanned
            FoldersScanned       = $foldersScanned
            TotalItemsScanned    = $totalItems
            TotalSizeGB          = $totalSizeGB
            LargestFileMB        = $largestFileMB
            AccessFailures       = $accessFailures
            RootPathErrors       = $rootPathErrors
            LongPathCount        = $longPathCount
            InvalidNameCount     = $invalidNameCount
            LargeFileCount       = $largeFileCount
            VeryLargeFileCount   = $veryLargeFileCount
            OverMaxFileSizeCount = $overMaxFileSizeCount
        }
    }
    catch {
        try {
            Write-CsvRow -Writer $feWriter -Values @((Get-Date -Format "yyyy-MM-dd HH:mm:ss"),$rootPath,$_.Exception.GetType().FullName,$_.Exception.Message,$_.ScriptStackTrace)
        } catch {}
        Append-ProgressLine ('"{0}","Failed","{1}","{2}","0","0","1","{3}"' -f (Get-Date).ToString("yyyy-MM-dd HH:mm:ss"),$rootIndex,($rootPath -replace '"','""'),($_.Exception.Message -replace '"','""'))
        [PSCustomObject]@{
            Index = $rootIndex; Path = $rootPath; ScanDate = (Get-Date -Format "yyyy-MM-dd"); ScanStart = $scanStart.ToString("yyyy-MM-dd HH:mm:ss"); ScanEnd = (Get-Date -Format "yyyy-MM-dd HH:mm:ss"); ScanDurationSeconds = 0
            RootAccessible = "False"; PermissionStatus = "FAIL"; MigrationReady = "No"; Readiness = "Red"; ReadinessReason = "Fatal scan error"; ImpactLevel = "High"; ImpactNotes = $_.Exception.Message
            FilesScanned = 0; FoldersScanned = 0; TotalItemsScanned = 0; TotalSizeGB = 0; LargestFileMB = 0; AccessFailures = 1; RootPathErrors = 1; LongPathCount = 0; InvalidNameCount = 0; LargeFileCount = 0; VeryLargeFileCount = 0; OverMaxFileSizeCount = 0
        }
    }
    finally {
        foreach ($writer in @($pfWriter,$reWriter,$lpWriter,$niWriter,$feWriter,$invWriter,$aclWriter)) {
            if ($null -ne $writer) { $writer.Flush(); $writer.Close(); $writer.Dispose() }
        }
    }
} -ThrottleLimit $ThrottleLimit -AsJob

$lastCompleted = -1
while ($scanJob.State -eq 'Running') {
    Start-Sleep -Seconds $ProgressRefreshSeconds
    $progressRows = @()
    try { $progressRows = @(Import-Csv -LiteralPath $ProgressFile) } catch { $progressRows = @() }
    $completedRows = @($progressRows | Where-Object { $_.Status -eq 'Completed' -or $_.Status -eq 'Failed' })
    $startedRows = @($progressRows | Where-Object { $_.Status -eq 'Started' })
    $completedCount = @($completedRows).Count
    $startedCount = @($startedRows).Count
    $activeApprox = [math]::Max(0, $startedCount - $completedCount)
    $percent = if ($TotalRoots -gt 0) { [math]::Min(100, [math]::Round(($completedCount / $TotalRoots) * 100, 0)) } else { 0 }

    Write-Progress -Activity "Migration readiness parallel scan" -Status "$completedCount of $TotalRoots complete; approx active workers: $activeApprox" -PercentComplete $percent
    if ($completedCount -ne $lastCompleted) {
        Write-Host ("[{0}] Scanning... Completed {1}/{2}; Active approx: {3}" -f (Get-Date -Format "HH:mm:ss"),$completedCount,$TotalRoots,$activeApprox) -ForegroundColor Cyan
        $lastCompleted = $completedCount
    }
}

Write-Progress -Activity "Migration readiness parallel scan" -Completed
$jobResults = @(Receive-Job -Job $scanJob -Wait -AutoRemoveJob)
Write-Log "Parallel scan completed. Result count: $(@($jobResults).Count)"
Write-Host "Step 3/5: Parallel scan complete." -ForegroundColor Green

Write-Host "Step 4/5: Merging detail files and building consolidated summary." -ForegroundColor Green
Write-Log "Merging detail outputs."

$PermissionFailuresCsv = Join-Path $OutputFolder "Permission_Failures.csv"
$RootErrorsCsv = Join-Path $OutputFolder "Root_Path_Errors.csv"
$InventoryCsv = Join-Path $OutputFolder "Inventory.csv"
$PermissionsCsv = Join-Path $OutputFolder "Permissions.csv"
$LongPathCsv = Join-Path $OutputFolder "Long_Path_Report.csv"
$NamingIssuesCsv = Join-Path $OutputFolder "Naming_Issues.csv"
$FatalErrorsCsv = Join-Path $OutputFolder "Fatal_Scan_Errors.csv"
$ConsolidatedCsv = Join-Path $OutputFolder "Consolidated_User_Readiness.csv"

[void](Merge-CsvParts -TempFolder $TempFolder -FileName "Permission_Failures.csv" -DestinationPath $PermissionFailuresCsv)
[void](Merge-CsvParts -TempFolder $TempFolder -FileName "Root_Path_Errors.csv" -DestinationPath $RootErrorsCsv)
[void](Merge-CsvParts -TempFolder $TempFolder -FileName "Long_Path_Report.csv" -DestinationPath $LongPathCsv)
[void](Merge-CsvParts -TempFolder $TempFolder -FileName "Naming_Issues.csv" -DestinationPath $NamingIssuesCsv)
[void](Merge-CsvParts -TempFolder $TempFolder -FileName "Fatal_Scan_Errors.csv" -DestinationPath $FatalErrorsCsv -Required $false)
if ($ExportInventory) { [void](Merge-CsvParts -TempFolder $TempFolder -FileName "Inventory.csv" -DestinationPath $InventoryCsv) }
if ($ExportAclOnFailures) { [void](Merge-CsvParts -TempFolder $TempFolder -FileName "Permissions.csv" -DestinationPath $PermissionsCsv) }

$resultMap = @{}
foreach ($r in @($jobResults)) { $resultMap[[int]$r.Index] = $r }

$ConsolidatedRows = [System.Collections.Generic.List[object]]::new()
foreach ($scanRow in @($PreparedRowsArray)) {
    $original = $scanRow.InputJson | ConvertFrom-Json
    $result = if ($resultMap.ContainsKey([int]$scanRow.Index)) { $resultMap[[int]$scanRow.Index] } else { $null }

    $ordered = [ordered]@{}
    foreach ($col in @($InputColumns)) { $ordered[$col] = $original.$col }

    if ($null -ne $result) {
        foreach ($prop in @($result.PSObject.Properties.Name)) {
            if ($prop -in @('Index','Path')) { continue }
            $ordered[$prop] = $result.$prop
        }
        $ordered["EstimatedMigrationHours"] = if ([double]$result.TotalSizeGB -gt 0) { [math]::Round(([double]$result.TotalSizeGB / $EstimatedThroughputGBPerHour), 2) } else { 0 }
    }
    else {
        $ordered["ScanDate"] = Get-Date -Format "yyyy-MM-dd"
        $ordered["PermissionStatus"] = "UNKNOWN"
        $ordered["MigrationReady"] = "No"
        $ordered["Readiness"] = "Red"
        $ordered["ReadinessReason"] = "No scan result returned for this row"
        $ordered["EstimatedMigrationHours"] = 0
    }

    [void]$ConsolidatedRows.Add([PSCustomObject]$ordered)
}

$ConsolidatedRowsArray = @($ConsolidatedRows.ToArray())
$ConsolidatedRowsArray | Export-Csv -LiteralPath $ConsolidatedCsv -NoTypeInformation -Encoding UTF8
Write-Log "Consolidated user readiness CSV created: $ConsolidatedCsv"

$totalUsers = @($ConsolidatedRowsArray).Count
$readyUsers = @($ConsolidatedRowsArray | Where-Object { $_.MigrationReady -eq "Yes" }).Count
$notReadyUsers = $totalUsers - $readyUsers
$permissionFailUsers = @($ConsolidatedRowsArray | Where-Object { $_.PermissionStatus -eq "FAIL" }).Count
$yellowUsers = @($ConsolidatedRowsArray | Where-Object { $_.Readiness -eq "Yellow" }).Count
$redUsers = @($ConsolidatedRowsArray | Where-Object { $_.Readiness -eq "Red" }).Count
$totalFiles = [int64](($ConsolidatedRowsArray | Measure-Object -Property FilesScanned -Sum).Sum)
$totalFolders = [int64](($ConsolidatedRowsArray | Measure-Object -Property FoldersScanned -Sum).Sum)
$totalItems = [int64](($ConsolidatedRowsArray | Measure-Object -Property TotalItemsScanned -Sum).Sum)
$totalSizeGB = [math]::Round([double](($ConsolidatedRowsArray | Measure-Object -Property TotalSizeGB -Sum).Sum),3)
$totalSizeTB = [math]::Round(($totalSizeGB / 1024),3)
$totalEstimatedHours = [math]::Round([double](($ConsolidatedRowsArray | Measure-Object -Property EstimatedMigrationHours -Sum).Sum),2)
$totalAccessFailures = [int64](($ConsolidatedRowsArray | Measure-Object -Property AccessFailures -Sum).Sum)
$totalLongPaths = [int64](($ConsolidatedRowsArray | Measure-Object -Property LongPathCount -Sum).Sum)
$totalInvalidNames = [int64](($ConsolidatedRowsArray | Measure-Object -Property InvalidNameCount -Sum).Sum)
$totalVeryLargeFiles = [int64](($ConsolidatedRowsArray | Measure-Object -Property VeryLargeFileCount -Sum).Sum)
$totalOverMaxFiles = [int64](($ConsolidatedRowsArray | Measure-Object -Property OverMaxFileSizeCount -Sum).Sum)

$DashboardRows = @(
    [PSCustomObject]@{ Metric = "Users Scanned"; Value = $totalUsers }
    [PSCustomObject]@{ Metric = "Ready To Migrate"; Value = $readyUsers }
    [PSCustomObject]@{ Metric = "Not Ready"; Value = $notReadyUsers }
    [PSCustomObject]@{ Metric = "Permission Failure Users"; Value = $permissionFailUsers }
    [PSCustomObject]@{ Metric = "Yellow Readiness Users"; Value = $yellowUsers }
    [PSCustomObject]@{ Metric = "Red Readiness Users"; Value = $redUsers }
    [PSCustomObject]@{ Metric = "Total Files"; Value = $totalFiles }
    [PSCustomObject]@{ Metric = "Total Folders"; Value = $totalFolders }
    [PSCustomObject]@{ Metric = "Total Items"; Value = $totalItems }
    [PSCustomObject]@{ Metric = "Total Size GB"; Value = $totalSizeGB }
    [PSCustomObject]@{ Metric = "Total Size TB"; Value = $totalSizeTB }
    [PSCustomObject]@{ Metric = "Estimated Migration Hours"; Value = $totalEstimatedHours }
    [PSCustomObject]@{ Metric = "Access Failures"; Value = $totalAccessFailures }
    [PSCustomObject]@{ Metric = "Long Path Items"; Value = $totalLongPaths }
    [PSCustomObject]@{ Metric = "Invalid Name Items"; Value = $totalInvalidNames }
    [PSCustomObject]@{ Metric = "Very Large Files"; Value = $totalVeryLargeFiles }
    [PSCustomObject]@{ Metric = "Over Max File Size"; Value = $totalOverMaxFiles }
    [PSCustomObject]@{ Metric = "Report Generated"; Value = (Get-Date -Format "yyyy-MM-dd HH:mm:ss") }
)

Write-Host "Step 5/5: Creating consolidated Excel report." -ForegroundColor Green
$excelResult = New-ReadinessExcelReport -OutputFolder $OutputFolder -DashboardRows $DashboardRows -UserSummaryRows $ConsolidatedRowsArray -PermissionFailuresCsv $PermissionFailuresCsv -LongPathCsv $LongPathCsv -NamingIssuesCsv $NamingIssuesCsv -FatalErrorsCsv $FatalErrorsCsv -InstallImportExcel:$InstallImportExcel
Write-Log "Excel report created: $($excelResult.Path)"

$ScriptEnd = Get-Date
$totalDuration = [math]::Round(($ScriptEnd - $ScriptStart).TotalSeconds,2)
Write-Log "Scan complete. Total duration seconds: $totalDuration"

if (-not $KeepTempFiles) {
    try { Remove-Item -LiteralPath $TempFolder -Recurse -Force; Write-Log "Temporary scan parts removed." }
    catch { Write-Log "Unable to remove temporary scan parts: $($_.Exception.Message)" "WARN" }
}
else { Write-Log "Temporary scan parts retained: $TempFolder" }

Write-Host ""
Write-Host "====================================================" -ForegroundColor Green
Write-Host "Migration Readiness Assessment Complete" -ForegroundColor Green
Write-Host "====================================================" -ForegroundColor Green
Write-Host "Output Folder              : $OutputFolder"
Write-Host "Consolidated CSV           : $ConsolidatedCsv"
Write-Host "Excel Report               : $($excelResult.Path)"
Write-Host "Excel Format               : $($excelResult.Format)"
Write-Host "Permission Failures        : $PermissionFailuresCsv"
Write-Host "Long Path Report           : $LongPathCsv"
Write-Host "Naming Issues              : $NamingIssuesCsv"
Write-Host "Fatal Scan Errors          : $FatalErrorsCsv"
Write-Host "Scan Log                   : $LogPath"
Write-Host ""
Write-Host "Totals:" -ForegroundColor Yellow
Write-Host "  Users Scanned            : $totalUsers"
Write-Host "  Ready To Migrate         : $readyUsers"
Write-Host "  Not Ready                : $notReadyUsers"
Write-Host "  Total Files              : $totalFiles"
Write-Host "  Total Size GB            : $totalSizeGB"
Write-Host "  Access Failures          : $totalAccessFailures"
Write-Host "  Long Paths               : $totalLongPaths"
Write-Host "  Invalid Names            : $totalInvalidNames"
Write-Host "====================================================" -ForegroundColor Green
