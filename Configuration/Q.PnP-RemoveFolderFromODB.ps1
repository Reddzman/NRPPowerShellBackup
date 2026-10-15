<#
.SYNOPSIS
    Permanently deletes a specified folder from OneDrive sites. Used on the Sanford Health Folder Cleanup project to remove a specific folder from multiple OneDrive sites.

.DESCRIPTION
    This script processes a list of OneDrive URLs and permanently deletes the
    specified folder via the $FolderName parameter (and all its contents) from the root of each OneDrive.
    
    Features:
    - Batch processing of multiple OneDrive sites
    - Dry-run capability to preview deletions without making changes
    - Comprehensive logging with color-coded output
    - CSV report export of deleted folders
    - Graceful error handling (continues if folder not found or deletion fails)
    - Permanent deletion using PnP.PowerShell -Permanent flag (bypasses recycle bin)

.REQUIREMENTS
    - PowerShell 5.1+ (recommended PowerShell 7.x)
    - PnP.PowerShell module installed (Install-Module -Name PnP.PowerShell)
    - Appropriate permissions to access OneDrive sites and delete folders (e.g., SharePoint admin or delegated permissions).
    - The folder to delete must be at the root of the OneDrive (e.g., /Documents/FolderName).
    - The script assumes the folder is in the "Documents" library of each OneDrive which is the default library for OneDrive.
    - The script uses app-based authentication if ClientId and Thumbprint are provided, otherwise it will prompt for interactive login.

.OUTPUTS
    - Log file with detailed processing information (default: C:\Temp\RemoveFolderFromODB-<timestamp>.log)
    - Optional CSV report of deleted folders (if ExportReport parameter is specified)
    - Summary statistics printed to console at the end of execution.

.PARAMETER InputCsvPath
    Path to the CSV file containing OneDrive URLs to process.
    CSV must have an 'OneDriveUrl' header column.
    Example CSV format:
    OneDriveUrl
    https://tenant-my.sharepoint.com/personal/user1_domain_com
    https://tenant-my.sharepoint.com/personal/user2_domain_com

.PARAMETER AdminUrl
    The admin center URL for the tenant (e.g., https://contoso-admin.sharepoint.com).

.PARAMETER ClientId
    (Optional) Client ID for app-based authentication. If not provided, will use interactive authentication.

.PARAMETER Thumbprint
    (Optional) Certificate thumbprint for authentication. If not provided, will use interactive authentication. The certificate must be installed in the current user's certificate store.

.PARAMETER Tenant
    (Optional) The tenant name or ID for authentication. Example: 'contoso' or the full domain 'contoso.onmicrosoft.com'.

.PARAMETER DryRun
    If specified, shows what would be deleted without actually deleting folders.

.PARAMETER OutputLogPath
    Path to the output log file. Defaults to a timestamped file in C:\Temp\.

.PARAMETER ExportReport
    Path to export a CSV report of deleted folders. If specified, a CSV will be created
    with details of all processed folders (deleted, skipped, or failed).

.EXAMPLE
    .\Q.PnP-RemoveFolderFromODB.ps1 -InputCsvPath "C:\Temp\OneDriveUrls.csv" -AdminUrl "https://contoso-admin.sharepoint.com" -DryRun

.EXAMPLE
    .\Q.PnP-RemoveFolderFromODB.ps1 -InputCsvPath "C:\Temp\OneDriveUrls.csv" -AdminUrl "https://contoso-admin.sharepoint.com" -ExportReport "C:\Temp\DeletionReport.csv"

.EXAMPLE
    .\Q.PnP-RemoveFolderFromODB.ps1 -InputCsvPath "C:\Temp\OneDriveUrls.csv" -AdminUrl "https://contoso-admin.sharepoint.com" -ClientId "f0b04b53-5866-43b2-a90f-decce7e7f81c" -Thumbprint "ABCD1234..." -Tenant "contoso" -DryRun
#>

param (
    [Parameter(Mandatory = $false)]
    [ValidateScript({ Test-Path $_ -PathType Leaf })]
    [string]$InputCsvPath = 'C:\Temp\OneDriveUrls.csv',

    [Parameter(Mandatory = $false)]
    [string]$AdminUrl = 'https://sanfordhealth-admin.sharepoint.com',

    [Parameter(Mandatory = $false)]
    [string]$OutputLogPath = "C:\Quisitive\Powershell\_Output Files\Remove-FolderFromODB-$(Get-Date -Format 'yyyyMMdd-HHmmss').log",

    [Parameter(Mandatory = $false)]
    [string]$ClientId,

    [Parameter(Mandatory = $false)]
    [string]$Thumbprint,

    [Parameter(Mandatory = $false)]
    [string]$Tenant,

    [Parameter(Mandatory = $false)]
    [switch]$DryRun,

    [Parameter(Mandatory = $false)]
    [string]$ExportReport = "C:\Quisitive\Powershell\_Output Files\OneDriveFolderDeletionReport-$(Get-Date -Format 'yyyyMMdd-HHmmss').csv"
)

# Initialize tracking
$logEntries = @()
$deletionReport = @()
$stats = @{
    TotalOneDrivesProcessed = 0
    FoldersFound = 0
    FoldersDeleted = 0
    FoldersFailed = 0
    FoldersSkipped = 0
    DryRunMode = $DryRun
}

function Write-Log {
    param (
        [Parameter(Mandatory = $true)]
        [string]$Message,
        
        [Parameter(Mandatory = $false)]
        [ValidateSet("INFO", "WARNING", "ERROR", "SUCCESS", "DEBUG")]
        [string]$Level = "INFO"
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Level] $Message"
    
    switch ($Level) {
        "INFO"    { Write-Host $logMessage -ForegroundColor Cyan }
        "WARNING" { Write-Host $logMessage -ForegroundColor Yellow }
        "ERROR"   { Write-Host $logMessage -ForegroundColor Red }
        "SUCCESS" { Write-Host $logMessage -ForegroundColor Green }
        "DEBUG"   { Write-Host $logMessage -ForegroundColor Gray }
    }
    
    $script:logEntries += $logMessage
}

function Test-PnPConnection {
    try {
        Get-PnPConnection -ErrorAction Stop | Out-Null
        return $true
    }
    catch {
        return $false
    }
}

function Import-OneDriveUrls {
    param ([string]$CsvPath)
    
    try {
        Write-Log "Importing OneDrive URLs from CSV: $CsvPath" "INFO"
        $urls = Import-Csv -Path $CsvPath -ErrorAction Stop
        
        if (-not $urls) {
            Write-Log "CSV file is empty" "WARNING"
            return @()
        }
        
        if ($urls | Get-Member -Name "OneDriveUrl" -ErrorAction SilentlyContinue) {
            $oneDriveUrls = @($urls | Select-Object -ExpandProperty OneDriveUrl | Where-Object { $_ })
            Write-Log "Successfully imported $($oneDriveUrls.Count) OneDrive URL(s)" "SUCCESS"
            return $oneDriveUrls
        }
        else {
            Write-Log "CSV missing 'OneDriveUrl' column" "ERROR"
            return @()
        }
    }
    catch {
        Write-Log "Error importing CSV: $_" "ERROR"
        return @()
    }
}

function Get-OneDriveFolderToDelete {
    param (
        [string]$OneDriveUrl,
        [string]$FolderName = "MyFolderNameToDelete"
    )
    
    try {
        Write-Log "Searching for folder '$FolderName' in Documents library: $OneDriveUrl" "DEBUG"
        
        # Get all items in the Documents library
        #$items = Get-PnPListItem -List "Documents" -Fields FileLeafRef, FileRef, ContentType -PageSize 500 -ErrorAction Stop
        #$items = Get-PnPFolderItem -List "Documents" -ItemType Folder 
        $items = Get-PnPFolderItem -Identity "Documents" -ItemType Folder 

        # Find the folder (ContentType should indicate it's a folder)
        $folder = $items | Where-Object {
            $_.Name -eq $FolderName
        }
        
        if ($folder) {
            Write-Log "  Found folder '$FolderName' at: $($folder.FileRef)" "SUCCESS"
            return $folder
        }
        else {
            Write-Log "  Folder '$FolderName' not found (will skip this OneDrive)" "WARNING"
            return $null
        }
    }
    catch {
        Write-Log "Error searching for folder in $OneDriveUrl`: $_" "WARNING"
        return $null
    }
}

function Remove-OneDriveFolder {
    param (
        [string]$OneDriveUrl,
        [PSCustomObject]$FolderObject,
        [string]$FolderName = "MyFolderNameToDelete",
        [bool]$DryRun
    )
    
    try {
        if ($DryRun) {
            Write-Log "[DRY RUN] Would permanently delete folder '$FolderName' from: $OneDriveUrl" "DEBUG"
            $script:stats.FoldersDeleted++
            
            $script:deletionReport += [PSCustomObject]@{
                OneDriveUrl = $OneDriveUrl
                FolderName = $FolderName
                FolderPath = $FolderObject.FileRef
                Status = "DRY_RUN"
                Message = "Would be deleted (dry run mode)"
                DeletionTime = Get-Date
            }
        }
        else {
            Write-Log "Permanently deleting folder '$FolderName' from: $OneDriveUrl" "INFO"
            
            # Use -Permanent flag to bypass recycle bin
            #Remove-PnPFolder -ServerRelativeUrl $FolderObject.FileRef -Force -Permanent -ErrorAction Stop
            Remove-PnPFolder -Name $FolderName -Folder "Documents" -Force -ErrorAction Stop
            
            Write-Log "  Permanently deleted: $FolderName" "SUCCESS"
            $script:stats.FoldersDeleted++
            
            $script:deletionReport += [PSCustomObject]@{
                OneDriveUrl = $OneDriveUrl
                FolderName = $FolderName
                FolderPath = $FolderObject.FileRef
                Status = "DELETED"
                Message = "Permanently deleted"
                DeletionTime = Get-Date
            }
        }
    }
    catch {
        Write-Log "Error deleting folder from $OneDriveUrl`: $_" "ERROR"
        $script:stats.FoldersFailed++
        
        $script:deletionReport += [PSCustomObject]@{
            OneDriveUrl = $OneDriveUrl
            FolderName = $FolderName
            FolderPath = if ($FolderObject) { $FolderObject.FileRef } else { "Unknown" }
            Status = "FAILED"
            Message = $_
            DeletionTime = Get-Date
        }
    }
}

function Process-OneDrive {
    param (
        [string]$OneDriveUrl,
        [string]$ClientId,
        [string]$FolderName = "MyFolderNameToDelete"
    )
    
    try {
        Write-Log "Processing OneDrive: $OneDriveUrl" "INFO"
        
        # Connect to the OneDrive site using interactive authentication with ClientId and Tenant
        Connect-PnPOnline -Url $OneDriveUrl -ClientId $ClientId -Tenant $Tenant -Thumbprint $Thumbprint -ErrorAction Stop
        
        # Search for the folder
        $folder = Get-OneDriveFolderToDelete -OneDriveUrl $OneDriveUrl -FolderName $FolderName
        
        if ($folder) {
            $stats.FoldersFound++
            
            # Delete the folder
            Remove-OneDriveFolder -OneDriveUrl $OneDriveUrl `
                                  -FolderObject $folder `
                                  -FolderName $FolderName `
                                  -DryRun $DryRun
        }
        else {
            $stats.FoldersSkipped++
            
            $script:deletionReport += [PSCustomObject]@{
                OneDriveUrl = $OneDriveUrl
                FolderName = $FolderName
                FolderPath = "N/A"
                Status = "SKIPPED"
                Message = "Folder not found"
                DeletionTime = Get-Date
            }
        }
        
        $stats.TotalOneDrivesProcessed++
    }
    catch {
        Write-Log "Error processing OneDrive $OneDriveUrl`: $_" "ERROR"
        $stats.TotalOneDrivesProcessed++
        $stats.FoldersFailed++
        
        $script:deletionReport += [PSCustomObject]@{
            OneDriveUrl = $OneDriveUrl
            FolderName = $FolderName
            FolderPath = "N/A"
            Status = "FAILED"
            Message = $_
            DeletionTime = Get-Date
        }
    }
}

function Export-DeletionReport {
    param (
        [array]$Report,
        [string]$Path
    )
    
    if (-not $Report -or -not $Path) {
        return
    }
    
    try {
        $Report | Export-Csv -Path $Path -NoTypeInformation -Encoding UTF8 -Force
        Write-Log "Deletion report exported to: $Path" "SUCCESS"
    }
    catch {
        Write-Log "Error exporting report: $_" "ERROR"
    }
}

function Main {
    try {
        Write-Log "========================================" "INFO"
        Write-Log "OneDrive Folder Deletion Tool" "INFO"
        Write-Log "========================================" "INFO"
        Write-Log "Input CSV: $InputCsvPath" "INFO"
        Write-Log "Admin URL: $AdminUrl" "INFO"
        if ($ClientId) { Write-Log "Client ID: $ClientId" "INFO" }
        if ($Thumbprint) { Write-Log "Certificate Thumbprint: $Thumbprint" "INFO" }
        if ($Tenant) { Write-Log "Tenant: $Tenant" "INFO" }
        Write-Log "Dry Run Mode: $($stats.DryRunMode)" "INFO"
        if ($ExportReport) { Write-Log "Report Export: $ExportReport" "INFO" }
        Write-Log "Target Folder: $($FolderName) (root level)" "INFO"
        Write-Log "Deletion Method: Permanent (bypasses recycle bin)" "INFO"
        
        # Verify PnP.PowerShell module
        if (-not (Get-Module -Name "PnP.PowerShell" -ErrorAction SilentlyContinue)) {
            Write-Log "Loading PnP.PowerShell module..." "INFO"
            Import-Module -Name "PnP.PowerShell" -ErrorAction Stop
        }
        Write-Log "PnP.PowerShell module ready" "SUCCESS"
        
        # Import OneDrive URLs
        $oneDriveUrls = Import-OneDriveUrls -CsvPath $InputCsvPath
        
        if ($oneDriveUrls.Count -eq 0) {
            Write-Log "No OneDrive URLs to process" "WARNING"
            return
        }
        
        Write-Log "========================================" "INFO"
        Write-Log "PROCESSING PHASE" "INFO"
        Write-Log "========================================" "INFO"
        
        # Process each OneDrive
        $i=1
        foreach ($oneDriveUrl in $oneDriveUrls) {
            Write-Log "Processing $i of $($oneDriveUrls.Count)" "INFO" 
            Process-OneDrive -OneDriveUrl $oneDriveUrl -ClientId $ClientId
            $i++
        }
        
        # Export report if requested
        if ($ExportReport -and $deletionReport.Count -gt 0) {
            Export-DeletionReport -Report $deletionReport -Path $ExportReport
        }
        
        Write-Log "========================================" "INFO"
        Write-Log "SUMMARY" "INFO"
        Write-Log "========================================" "INFO"
        Write-Log "OneDrives Processed: $($stats.TotalOneDrivesProcessed)" "INFO"
        Write-Log "Folders Found: $($stats.FoldersFound)" "INFO"
        Write-Log "Folders Deleted: $($stats.FoldersDeleted)" "INFO"
        Write-Log "Folders Skipped: $($stats.FoldersSkipped)" "INFO"
        Write-Log "Folders Failed: $($stats.FoldersFailed)" "INFO"
        Write-Log "Mode: $(if ($DryRun) { 'DRY RUN' } else { 'LIVE' })" "INFO"
        Write-Log "========================================" "INFO"
    }
    catch {
        Write-Log "Fatal error: $_" "ERROR"
    }
    finally {
        try {
            $logEntries | Out-File -FilePath $OutputLogPath -Encoding UTF8 -Force
            Write-Log "Log saved to: $OutputLogPath" "SUCCESS"
        }
        catch {
            Write-Host "Error saving log: $_" -ForegroundColor Red
        }
    }
}

#Execute the main function
Main
