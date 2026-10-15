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

#region Script Parameters
Param(	[string] $inputFile = "{Teams input file path relative to _Input Files}", # Location from _Input Files folder - IE. "sample\filename.csv" or "filename.csv"
		[string] $clientAppID = $null, # Leave Blank for interactive login
		[string] $clientCertPath = "{Path to client certificate file must be installed on local user}",# Leave Blank for interactive login
		[string] $clientTenant = "{Client tenant SharePoint prefix}", # IE. ClientTenant.sharepoint.com
		[string] $outputLocation = "{c:\path to output files location}", # A /clienttenant/datetime/ folder will be created for output files
		[boolean] $boolean = $false, # $true/$false
		[string] $clientInteractiveAppID = $null # IE. PnP Now requires an app registration be deployed to support interactive login. Provide the appropriate App ID.
)
#endregion Script Parameters

#region Start Date / Time
# Version Update Check
$env:PNPPOWERSHELL_UPDATECHECK="false"

# Remove Existing Catapult Modules to ensure PS session has most up to date functions
Get-Module Cat.* | Remove-Module

# Load General Functions
Import-Module "..\General\Cat.GeneralFunctions.psm1"

# Get Start Date / Time
$scriptStart = Start-CatScript

# Start Transcript
Start-Transcript -Path ("..\_Log Files\{0}.Tech-ScriptName-{1}.log" -f $clientTenant, $scriptStart[1])
#endregion Start Date / Time

#region 1.	Snapin(s), Modules & Environment Connection(s)
# Load MS / PnP PowerShell Modules
Get-Module -ListAvailable | Where-Object{($_.Name -like "*SharePoint*") -or ($_.Name -like "*PnP*") } | Remove-Module
Import-Module PnP.PowerShell
Import-Module Microsoft.Online.SharePoint.PowerShell -DisableNameChecking -erroraction SilentlyContinue
Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue
Import-Module AzureAD -erroraction SilentlyContinue
Import-Module MSOnline -erroraction SilentlyContinue
Import-module MicrosoftTeams -erroraction SilentlyContinue
Import-Module -Name Microsoft.PowerApps.Administration.PowerShell -DisableNameChecking -erroraction SilentlyContinue
Import-Module -Name Microsoft.PowerApps.PowerShell -erroraction SilentlyContinue
Import-Module Microsoft365DSC

# Load Custom Modules
Import-Module "..\FolderPath\ModuleName.psm1"
#endregion Snapin(s), Modules & Environment Connection(s)

#region 2.	Input & Output Files
# File Paths
$inPath = ("..\_Input Files\{0}" -f $inputFile)
$outputPath = ("{0}\{1}\{2}" -f $outputLocation, $clientTenant, $scriptStart[1])
$errorPath = ("{0}\errors.csv" -f $outputPath)
$outputFile = ("{0}\file.csv" -f $outputPath)

# Import Input File(s)
# Load CSV file
If (Test-Path ($inPath)){
	$inFile = Import-Csv $inPath -Encoding Unicode
}
else {
	Write-Host
	throw ("The input file does not exist: {0}" -f $inPath)	
}
[xml]$inFile = Get-Content $inPath # Importing XML File
$inFile = Get-Content $inPath -Raw # Importing JSON File
#endregion Input & Output Files

#region 3.	Variables
# User Inputs
If (($clientAppID) -or ($boolean)) {
	Write-Host
	Write-Host "User Input Variables" -ForegroundColor Cyan
}
If ($boolean) {
	Write-Host
	Write-Host "*** Bypassing the recycle bin will perminantly delete files from the Input File and there is no way to recover from this deletion. ***" -ForegroundColor Red
	$proceed = Read-Host "   Are you sure you want to proceed with PERMINANTLY deleting files in the input file (Yes or No)" -ForegroundColor Red

	# if user chose to proceed
	if ($proceed.ToLower() -eq "yes") {
		$boolean = $true
	}
	else {
		$boolean = $false
	}
}
if ($clientAppID) {
	$certPass = ConvertTo-SecureString -String (Read-Host -Prompt "   Enter certificate password for the app registration") -AsPlainText -Force
	$cert = Get-PnPAzureCertificate -Path $clientCertPath -Password $certPass
	Remove-Variable certPass
	Remove-Variable clientCertPath
}

# Script Variables
$tenantDomain = ("{0}.onmicrosoft.com" -f $clientTenant)
$spoAdminUrl = ("https://{0}-admin.sharepoint.com" -f $clientTenant)
$itemCount = 5000
$specialCharacters = "[~``!@#$%^&*()_-]" #(($pnpGroup.Description -replace $specialCharacters,'') -replace "`n",' ') -replace "`r",' '
$excludedColumns = @("A2ODMountCount",
						"AppAuthor",
						"AppEditor",
						"xd_Signature")
$query = `
"<View Scope='Recursive'>
	<Query>
		<Where>
			<Eq>
				<FieldRef Name='FSObjType' />
				<Value Type='Integer'>0</Value>
			</Eq>
		</Where>
		<QueryOptions>
			<ViewAttributes Scope='Recursive' />
		</QueryOptions>
	</Query>
</View>"

#region Output Parameters & Variables
Write-Host
Write-Host "Script Parameters" -ForegroundColor Cyan
Write-Host ("   Input File: {0}" -f $inputFile)
if ($clientAppID) {
	Write-Host ("   App ID: {0}" -f $clientAppID)
}
Write-Host ("   Client Tenant Prefix: {0}" -f $clientTenant)
Write-Host
Write-Host "Output Variables" -ForegroundColor Cyan
Write-Host ("   Output Path: {0}" -f $outputPath)
Write-Host ("   Error Path: {0}" -f $errorPath)
Write-Host ("   Existing Files Path: {0}" -f $existingFilesReport)
If (($clientAppID) -or ($boolean)) {
	Write-Host
	Write-Host "User Inputs" -ForegroundColor Cyan
}
If ($boolean) {
	Write-Host ("   Proceed with deleting files: {0}" -f $boolean)
}
if ($clientAppID) {
	Write-Host ("   Certificate Thumbprint: {0}" -f "Only output for troubleshooting", $cert.Thumbprint)
}
Write-Host
Write-Host "Script Variables" -ForegroundColor Cyan
Write-Host ("   Microsoft Tenant: {0}" -f $tenantDomain)
Write-Host ("   SPO admin URL: {0}" -f $spoAdminUrl)
Write-Host ("   itemCount: {0}" -f $itemCount)
Write-Host ("   Special Characters to Remove: {0}" -f $specialCharacters)
Write-Host
Write-Host "Array of Info" -ForegroundColor Cyan
Write-Host $arrayName
#endregion Output Parameters & Variables
#endregion
    
#region 4.	Main Script
#region Created directories for reports
Write-Host
Write-Host ("{0} - Creating Report Directories" -f (Get-Date -Format "yy-MM-dd HH:mm:ss")) -ForegroundColor Cyan
New-Item -Path $outputLocation -ItemType Directory
#endregion Created directories for reports

#region Connecting to SharePoint
Write-Host ("   {0} - Connecting to SharePoint" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"))
try {
	if ($clientAppID) {
		Connect-PnPOnline $spoAdminUrl -tenant $tenantDomain -ClientId $clientAppID -Thumbprint $cert.Thumbprint -ValidateConnection -ErrorAction Stop # PnP Modules w/ App Reg
	} else {
		Connect-PnPOnline -Url $spoAdminUrl -Interactive -ClientId $clientInteractiveAppID -ValidateConnection -ErrorAction Stop # Interactive Login requires PNP Interactive App Reg be deployed by a GA. 
	}
}
catch {
	Write-CatErrorCSV -funcCSVItem $list -funcAction "Connecting to SharePoint" -funcException $_.Exception -funcDetails $_.ErrorDetails -funcStackTrace $_.ScriptStackTrace -funcFilePath $errorPath
	Continue listLoop
}
#endregion Connecting to SharePoint

#region CSV Loop Example
Write-Host
Write-Host "Main CSV Loop" -ForegroundColor Magenta
:loopName foreach ($item in $inFile){
	Write-Host
	Write-Host ("   {0} - Doing Something per item - {1}" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"), $item.DisplayName)

	#region Sub Task / Error Handling Example
	Write-Host "      Doing Sub action" -ForegroundColor DarkGray
	try {
		$group = Get-PnPMicrosoft365Group -Identity $item.GroupId -ErrorAction Stop
	}
	catch {
		Write-CatErrorCSV -funcCSVItem $item -funcAction "Doing Sub action" -funcException $_.Exception -funcDetails $_.ErrorDetails -funcStackTrace $_.ScriptStackTrace -funcFilePath $errorPath
		Continue loopName
	}
	#endregion Sub Task / Error Handling Example
} 
#endregion CSV Loop Example

#endregion Main Script

#region Completed Script
# Disconnect from Services
Write-Host
Write-Host "Disconnecting from Services:" -ForegroundColor Cyan
Write-Host ("   Disconnecting from SharePoint Online: {0}" -f $spoAdminURL)
Disconnect-PnPOnline

# Output Script Info
Complete-CatScript $scriptStart[0]

# Stop logging
Stop-Transcript
#endregion Completed Script