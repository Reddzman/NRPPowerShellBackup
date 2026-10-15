<#  Cat.SG-LargeListMigrations.ps1

This script leverage PnP PowerShell and the ShareGate modules to break up a large list for migration purposes. The script gets the latest
SharePoint item ID and then iterates from 0 to latest and runs Sharegate migration jobs in batches to create smaller more managable jobs.

Use Case
	-	This script was origionally created for SP13 on premises to SharePoint Online.
	-	Assumes source and destination lists already exist.

Dependencies
	-	ShareGate PowerShell Snapins
	-	..\General\Cat.PnP-GeneralFunctionsFunctions.psm1
	-	..\Reporting\Cat.ReportingFunctions.psm1
    -	Input CSV File in Input Files folder - Script Parameter
    -   User mapping file export from ShareGate (UserMappingFile.sgum) - Script Parameter

Call script with parameters
	.\Cat.SG-LargeListMigrations.ps1 -inputFile Cat.SG-LargeLists-Test.csv -clientName "Client" -batchRange 30

Sample CSV
	-	..\Samples\ShareGate\Cat.SG-LargeLists-Test.csv

Table of Contents
	1.	Input & Output Files
	2.	Variables
	3.	Snapin(s), Modules & Environment Connection(s)
	4.	Main Script
	5.	Migrating List in Batches
	6.	Creating Migration Job
#>

#region Script Parameters
Param(	
		[string] $inputFile = "Samples\ShareGate\Cat.SG-LargeLists-Test.csv",
		[string] $userFile = "Samples\ShareGate\Cat.SG-UserMapping-Test.sgum",
		[string] $clientName = "Testing",
		[int] $batchRange = 30
)
#endregion

#region Start Date / Time
# Remove Existing Catapult Modules to ensure PS session has most up to date functions
Get-Module Cat.* | Remove-Module

# Load General Functions
Import-Module "..\General\Cat.GeneralFunctions.psm1"

# Get Start Date / Time
$scriptStart = Start-CatScript

# Start Transcript
Start-Transcript -Path ("..\_Log FIles\{0}-LargeListMigraitons-{1}.log" -f $clientName, $scriptStart[1])
#endregion

#region 1.	Input & Output Files
# File Paths
$inPath = ("..\_Input Files\{0}" -f $inputFile)
$userPath = ("..\_Input Files\{0}" -f $userFile)
$exportPath = "..\_Output FIles\"

# Import Input File(s)
# Load CSV file
If (Test-Path ($inPath)){
	$inFile = Import-Csv $inPath -Encoding Unicode
}
else {
	Write-Host
	throw ("The input file does not exist: {0}" -f $inPath)	
}

# Load user mapping file
If (Test-Path ($userPath)){
	$userPath = Resolve-Path $userPath
}
else {
	Write-Host
	throw ("The user file does not exist: {0}" -f $userPath)	
}
#endregion

#region 2.	Variables
# User Inputs
Write-Host
Write-Host "User Input Variables" -ForegroundColor Cyan
Write-Host "   Enter credentials for source SharePoint environment:"
$sourceCred = Get-Credential -Message "Enter credentials for source SharePoint environment:"
Write-Host "   Enter credentials for destination SharePoint environment:"
$destinationCred = Get-Credential -Message "Enter credentials for destination SharePoint environment:" 

# ShareGate Copy Settings
$sgOnWarning = "Continue" # Options: Continue | Cancel    Recommendation: Continue
$sgOnError = "Skip" # Options: Skip | SkipAllVersions | Cancel    Recommendation: Skip
$sgOnSiteObjectExists = "Merge" # Options: Merge | Skip    Recommendation: Merge
$sgOnContentItemExists = "IncrementalUpdate" # Options: Overwrite | Skip | Rename | IncrementalUpdate    Recommendation: IncrementalUpdate
$sgVersionOrModerationComment = $null # Recommendation: $null
$sgNormalMode = $true # Recommendation: $false
$sgWaitForImportCompletion = $true # Recommendation: $true

# ShareGate Property Settings
$sgCheckInAs = "SameAsCurrent" # Options: SameAsCurrent | Publish | Draft    Recommendation: SameAsCurrent
$sgContentApproval = "SameAsCurrent" # Options: Ignore | SameAsCurrent | Approved | Rejected | Pending | Draft | Scheduled    Recommendation: SameAsCurrent
$sgAuthorsAndTimestamps = $true # Recommendation: $true
$sgVersionHistory = $true # Recommendation: $true
[int] $sgVersionLimit = 5 # Recommendation: 5
$sgPermissions = $false # Recommendation: $true
$sgWebParts = $false # Recommendation: $true
$sgNoLinkCorrection = $true # Recommendation: $false
$sgFlattenFolders = $false # Recommendation: $false
$sgUserAlerts = $false # Recommendation: $false

# Output Parameters & Variables
Write-Host
Write-Host "Script Parameters" -ForegroundColor Cyan
Write-Host ("   Input File: {0}" -f $inputFile)
Write-Host ("   User File: {0}" -f $userFile)
Write-Host ("   Client Name: {0}" -f $clientName)
Write-Host ("   Batch Size (Days): {0}" -f $batchRange)
Write-Host
Write-Host "User Inputs" -ForegroundColor Cyan
Write-Host ("   Source Credentials: {0}" -f $sourceCred.UserName)
Write-Host ("   Destination Credentials: {0}" -f $destinationCred.UserName)
Write-Host
Write-Host "ShareGate Copy Settings" -ForegroundColor Cyan
Write-Host ("   On Warning: {0}" -f $sgOnWarning)
Write-Host ("   On Error: {0}" -f $sgOnError)
Write-Host ("   On Site Object Exists: {0}" -f $sgOnSiteObjectExists)
Write-Host ("   On Content Item Exists: {0}" -f $sgOnContentItemExists)
Write-Host ("   Version Comment: {0}" -f $sgVersionOrModerationComment)
Write-Host ("   Normal Mode: {0}" -f $sgNormalMode)
Write-Host ("   Wait for Import Completion: {0}" -f $sgWaitForImportCompletion)
Write-Host
Write-Host "ShareGate Property Settings" -ForegroundColor Cyan
Write-Host ("   Check In As: {0}" -f $sgCheckInAs)
Write-Host ("   Content Approval: {0}" -f $sgContentApproval)
Write-Host ("   Authors and Timestamps: {0}" -f $sgAuthorsAndTimestamps)
Write-Host ("   VersionHistory: {0}" -f $sgVersionHistory)
Write-Host ("   Version Limit: {0}" -f $sgVersionLimit)
Write-Host ("   Permissions: {0}" -f $sgPermissions)
Write-Host ("   Web Parts: {0}" -f $sgWebParts)
Write-Host ("   No Link Correction: {0}" -f $sgNoLinkCorrection)
Write-Host ("   Flatten Folders: {0}" -f $sgFlattenFolders)
Write-Host ("   User Alerts: {0}" -f $sgUserAlerts)
#endregion

#region 3.	Snapin(s), Modules & Environment Connection(s)
# Load Custom Modules
Import-Module "..\Reporting\Cat.ReportingFunctions.psm1"

# Load Sharegate modules
Import-Module Sharegate	
#endregion
    
#region 4.	Main Script
# Screen Output
Write-Host
Write-Host ("Migrating lists from CSV file: {0}" -f $inputFile) -ForegroundColor Cyan

# Loop through Lists
ForEach ($list in $inFile){
	Write-Host
	Write-Host ("   Migrating list: {0}/{1}" -f $list.sourceURL, $list.listPath) -ForegroundColor Magenta

	#region 5.	Migrating List in Batches
	Write-Host "      Configuring ShareGate job settings"
	
	# Loading Source List
	Write-Host ("         Loading source site: {0}" -f $list.sourceURL) -ForegroundColor DarkGray
	$sgSourceSite = Connect-Site -Url $list.sourceURL -Username $sourceCred.UserName -Password $sourceCred.Password
	Write-Host ("         Loading source list: {0}/{1}" -f $list.sourceURL,$list.listPath) -ForegroundColor DarkGray
	$sgSourceList = Get-List -Site $sgSourceSite -Name $list.listPath

	# Loading Destination List
	Write-Host ("         Loading destination site: {0}" -f $list.destinationURL) -ForegroundColor DarkGray
	$sgDestinationSite = Connect-Site -Url $list.destinationURL  -Username $destinationCred.UserName -Password $destinationCred.Password
	Write-Host ("         Loading destination list: {0}/{1}" -f $list.destinationURL, $list.listPath) -ForegroundColor DarkGray
	$sgDestinationList = Get-List -Site $sgDestinationSite -Name $list.listPath

	# Creating Copy Settings
	Write-Host "         Creating Copy Settings" -ForegroundColor DarkGray
	$sgBaseSettings = New-CopySettings `
		-OnWarning $sgOnWarning `
		-OnError $sgOnError `
		-OnSiteObjectExists $sgOnSiteObjectExists `
		-OnContentItemExists $sgOnContentItemExists `
		-VersionOrModerationComment $sgVersionOrModerationComment
	$sgBaseTaskName = ("{0} - {1}" -f $list.siteTitle, $list.listName)

	# Creating Mapping Settings
	Write-Host ("         Loading Mappings file: {0}" -f $userFile) -ForegroundColor DarkGray
	$sgMappings = New-MappingSettings
	$sgMappings = Import-UserAndGroupMapping -MappingSettings $sgMappings -Path $userPath.Path
	# Could also have Property mappings, site/library template mappings, ect.

	# Iterate through batchs of items
	$sgStartDate = (Get-Date $list.createdDate)
	$sgEndDate = $sgStartDate.AddDays($batchRange)
	$listLastModifiedDate = Get-Date $list.lastModifiedDate
	Write-Host ("      Running migration jobs")
	while ($sgStartDate -lt $listLastModifiedDate) {
		#region 6.	Creating Migration Job
		# Configure Unique Job Settings
		$sgPropertyTemplate = New-PropertyTemplate `
			-CheckInAs $sgCheckInAs `
			-ContentApproval $sgContentApproval `
			-AuthorsAndTimestamps:$sgAuthorsAndTimestamps `
			-VersionHistory:$sgVersionHistory `
			-VersionLimit $sgVersionLimit `
			-Permissions:$sgPermissions `
			-WebParts:$sgWebParts `
			-NoLinkCorrection:$sgNoLinkCorrection `
			-FlattenFolders:$sgFlattenFolders `
			-From $sgStartDate `
			-To $sgEndDate `
			-UserAlerts:$sgUserAlerts
		$sgTaskName = ("{0} - {1} to {2}" -f $sgBaseTaskName, $sgStartDate.ToString("MM/dd/yyyy"), $sgEndDate.ToString("MM/dd/yyyy"))

		# Running Migration Job
		Write-Host ("         Running migration job: {0}" -f $sgTaskName) -ForegroundColor DarkGray
		$sgResult = Copy-Content `
			-SourceList $sgSourceList `
			-DestinationList $sgDestinationList `
			-MappingSettings $sgMappings `
			-CopySettings $sgBaseSettings `
			-TaskName $sgTaskName `
			-NormalMode:$sgNormalMode `
			-WaitForImportCompletion:$sgWaitForImportCompletion
		#endregion

		# Exporting Job Results
		Write-Host ("            Exporting migration job results: {0}" -f $sgTaskName) -ForegroundColor Blue
		$sgExportPath = ("{0}\{1}-LargeListMigraitons-Results-{2}-{3}" -f $exportPath, $clientName, $sgTaskName.Replace(" ",""), $scriptStart[1])
		Export-Report -CopyResult $sgResult -Path $sgExportPath
		
		# Incrment Start & End ID number
		$sgStartDate = $sgEndDate
		$sgEndDate = $sgEndDate.AddDays($batchRange)
	}

	# Screen Output
	Write-Host ("      List migration complete") -ForegroundColor Green
	#endregion
}

# Screen Output
Write-Host
Write-Host ("All List Migrations Complete Complete") -ForegroundColor Green
#endregion

#region Completed Script
# Output Script Info
Complete-CatScript $scriptStart[0]

# Stop logging
Stop-Transcript
#endregion