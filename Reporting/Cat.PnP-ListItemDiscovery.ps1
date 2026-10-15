<#  Cat.PnP-ListItemDiscovery.ps1

This script iterates through a csv file of libraries or lists and outputs the folders and item metadata.

*** Needs to be tested/improved to handle a list with multiple content types.

Dependencies
	-	PnP.PowerShell Modules
	-	Azure App Registrations
	-	Input CSV File - Example: ..\_Input FIles\Samples\Reporting\Cat.PnP-ListItemDiscovery-Test.csv

Call script with parameters
	cd "C:\GitHub\CatPowerShell\Reporting\"
	.\Cat.PnP-ListItemDiscovery.ps1  -inputFile "Samples\Reporting\Cat.PnP-ListItemDiscovery-Test.csv" -clientAppID "caf6b11a-cf03-416c-b0f0-67a09a8e30ee" -clientCertPath "C:\GitHub\NRPDev\AppRegistration\PowerShell App Registration.pfx" -clientTenant "nrpdev2" -outputLocation "C:\GitHub\CatPowerShell\_Output Files" -itemBatchSize 100

Table of Contents
	1.	Snapin(s), Modules & Environment Connection(s)
	2.	Input & Output Files
	3.	Variables
	4.	Main Script
	5.	Connecting To SP Online
	6.	Retrieve List Information
	7.	Retrieve & Output All Item Information
#>

#region Script Parameters
Param(	[string] $inputFile = "{Teams input file path relative to _Input Files}", # IE. "sample\filename.csv"
		[string] $clientAppID = "{Client app registration ID from Azure}", # Leave Blank for interactive login
		[string] $clientCertPath = "{Path to client certificate file must be installed on local user}",# Leave Blank for interactive login
		[string] $clientTenant = "{Client tenant SharePoint prefix}", # IE. ClientTenant.sharepoint.com
		[string] $outputLocation = "{c:\path to output files location}", # A /clienttenant/datetime/ folder will be created for output files
		[int] $itemBatchSize = 4000
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
Start-Transcript -Path ("..\_Log Files\{0}.PnP-ListItemDiscovery-{1}.log" -f $clientTenant, $scriptStart[1])
#endregion

#region 1.	Snapin(s), Modules & Environment Connection(s)
# Load MS / PnP PowerShell Modules
Get-Module -ListAvailable | Where-Object{($_.Name -like "*SharePoint*") -or ($_.Name -like "*PnP*") } | Remove-Module
Import-Module PnP.PowerShell
#endregion

#region 2.	Input & Output Files
# File Paths
$inPath = ("..\_Input Files\{0}" -f $inputFile)
$outputPath = ("{0}\{1}\{2}" -f $outputLocation, $clientTenant, $scriptStart[1])
$errorPath = ("{0}\errors.csv" -f $outputPath)
$outputFoldersPath= ("{0}\Folders" -f $outputPath)
$outputItemsPath = ("{0}\Items" -f $outputPath)

# Import Input File(s)
# Load CSV file
If (Test-Path ($inPath)){
	$inFile = Import-Csv $inPath -Encoding Unicode
}
else {
	Write-Host
	throw ("The input file does not exist: {0}" -f $inPath)	
}
#endregion

#region 3.	Variables
# User Inputs
if ($clientAppID) {
	Write-Host
	Write-Host "User Input Variables" -ForegroundColor Cyan
	$certPass = ConvertTo-SecureString -String (Read-Host -Prompt "   Enter certificate password for the app registration") -AsPlainText -Force
	$cert = Get-PnPAzureCertificate -Path $clientCertPath -Password $certPass
	Remove-Variable certPass
	Remove-Variable clientCertPath
}

# Script Variables
$tenantDomain = ("{0}.onmicrosoft.com" -f $clientTenant)
$specialCharacters = "[~``!@#$%^*]" #(($pnpGroup.Description -replace $specialCharacters,'') -replace "`n",' ') -replace "`r",' '
$itemExcludedColumns = @("A2ODMountCount",
						"AppAuthor",
						"AppEditor",
						"BSN",
						"ContentVersion",
						"Created_x0020_By",
						"Created_x0020_Date",
						"DocConcurrencyNumber",
						"File_x0020_Type",
						"FileRef",
						"FolderChildCount",
						"FSObjType",
						"HTML_x0020_File_x0020_Type",
						"InstanceID",
						"ItemChildCount",
						"Last_x0020_Modified",
						"MediaLengthInSeconds",
						"MediaServiceAutoKeyPoints",
						"MediaServiceAutoTags",
						"MediaServiceDateTaken",
						"MediaServiceEventHashCode",
						"MediaServiceFastMetadata",
						"MediaServiceGenerationTime",
						"MediaServiceImageTags",
						"MediaServiceKeyPoints",
						"MediaLengthInSeconds",
						"MediaServiceLocation",
						"MediaServiceMetadata",
						"MediaServiceOCR",
						"MediaServiceSearchProperties",
						"MetaInfo",
						"Modified_x0020_By",
						"NoExecute",
						"Order",
						"OriginatorId",
						"owshiddenversion",
						"ParentVersionString",
						"ParentLeafName",
						"ProgId",
						"Restricted",
						"ScopeId",
						"SharedWithDetails",
						"SharedWithUsers",
						"SMLastModifiedDate",
						"SMTotalFileStreamSize",
						"SMTotalFileCount",
						"SMTotalSize",
						"SortBehavior",
						"StreamHash",
						"SyncClientId",
						"TriggerFlowInfo",
						"VirusStatus",
						"xd_ProgID",
						"xd_Signature")
$folderExcludedColumns = @("A2ODMountCount",
						"AppAuthor",
						"AppEditor",
						"BSN",
						"ContentVersion",
						"Created_x0020_By",
						"Created_x0020_Date",
						"DocConcurrencyNumber",
						"File_x0020_Type",
						"File_x0020_Size",
						"FileRef",
						"FSObjType",
						"HTML_x0020_File_x0020_Type",
						"IsCheckedoutToLocal",
						"Last_x0020_Modified",
						"MediaLengthInSeconds",
						"MediaServiceAutoKeyPoints",
						"MediaServiceAutoTags",
						"MediaServiceDateTaken",
						"MediaServiceEventHashCode",
						"MediaServiceFastMetadata",
						"MediaServiceGenerationTime",
						"MediaServiceImageTags",
						"MediaServiceKeyPoints",
						"MediaLengthInSeconds",
						"MediaServiceLocation",
						"MediaServiceMetadata",
						"MediaServiceOCR",
						"MediaServiceSearchProperties",
						"MetaInfo",
						"Modified_x0020_By",
						"NoExecute",
						"Order",
						"OriginatorId",
						"owshiddenversion",
						"ParentVersionString",
						"ParentLeafName",
						"ProgId",
						"Restricted",
						"ScopeId",
						"SharedWithDetails",
						"SharedWithUsers",
						"SMLastModifiedDate",
						"SMTotalFileStreamSize",
						"SMTotalFileCount",
						"SMTotalSize",
						"SortBehavior",
						"StreamHash",
						"SyncClientId",
						"TriggerFlowInfo",
						"VirusStatus",
						"xd_ProgID",
						"xd_Signature")

# Output Parameters & Variables
Write-Host
Write-Host "Script Parameters" -ForegroundColor Cyan
Write-Host ("   Input File: {0}" -f $inputFile)
if ($clientAppID) {
	Write-Host ("   App ID: {0}" -f $clientAppID)
}
Write-Host ("   Client Tenant Prefix: {0}" -f $clientTenant)
Write-Host ("   Output Location: {0}" -f $outputLocation)
Write-Host ("   Item Batch Size: {0}" -f $itemBatchSize)
if ($clientAppID) {
	Write-Host
	Write-Host "User Inputs" -ForegroundColor Cyan
	Write-Host ("   Certificate Thumbprint: {0}" -f "Only output for troubleshooting", $cert.Thumbprint)
}
Write-Host
Write-Host "Script Variables" -ForegroundColor Cyan
Write-Host ("   Microsoft Tenant: {0}" -f $tenantDomain)
Write-Host ("   Special Characters to Remove: {0}" -f $specialCharacters)
Write-Host
Write-Host "Item Excluded Columns" -ForegroundColor Cyan
Write-Host $itemExcludedColumns
Write-Host
Write-Host "Folder Excluded Columns" -ForegroundColor Cyan
Write-Host $itemExcludedColumns
#endregion
    
#region 4.	Main Script
#region Created directories for additional reports
Write-Host
Write-Host ("{0} - Creating Report Directories" -f (Get-Date -Format "yy-MM-dd HH:mm:ss")) -ForegroundColor Cyan
New-Item -Path $outputItemsPath -ItemType Directory
New-Item -Path $outputFoldersPath -ItemType Directory
#endregion

# Looping through List CSV File
Write-Host
Write-Host ("Scanning Lists: {0}" -f $inputFile) -ForegroundColor Cyan
:listLoop foreach ($List in $inFile){
	Write-Host
	Write-Host ("{0} - Scanning List - {1}" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"), $list.listName) -ForegroundColor Magenta

	#region 5.	Connecting To SP Online
	Write-Host ("   {0} - Connecting to SharePoint Online site: {1}" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"), $list.siteTitle)
	try {
		if ($clientAppID) {
			Connect-PnPOnline $list.siteURL -tenant $tenantDomain -ClientId $clientAppID -Thumbprint $cert.Thumbprint -ValidateConnection -ErrorAction Stop # PnP Modules w/ App Reg
		} else {
			Connect-PnPOnline -Url $list.siteURL -Interactive -ValidateConnection -ErrorAction Stop # PnP Modules w/o App Reg
		} 
	}
	catch {
		Write-CatErrorCSV -funcCSVItem $list -funcAction "Connecting to site" -funcException $_.Exception -funcDetails $_.ErrorDetails -funcStackTrace $_.ScriptStackTrace -funcFilePath $errorPath
		Continue listLoop
	}
	#endregion #>

	#region 6.	Retrieve List Information
	Write-Host ("   {0} - Retrieving List: {1}" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"), $list.listName)
	try {
		$pnpList = Get-PnPList $list.listName -ThrowExceptionIfListNotFound
	}
	catch {
		Write-CatErrorCSV -funcCSVItem $list -funcAction "Retrieving List" -funcException $_.Exception -funcDetails $_.ErrorDetails -funcStackTrace $_.ScriptStackTrace -funcFilePath $errorPath
		Continue listLoop
	}
	#endregion #>

	#region 7.	Retrieve & Output All Item Information
	Write-Host ("   {0} - Retrieving All Items" -f (Get-Date -Format "yy-MM-dd HH:mm:ss")) -NoNewline
	$outputItems = ("{0}\{1}-{2}-Items.csv" -f $outputItemsPath, (($list.siteTitle -replace $specialCharacters,'') -replace " ",''), (($list.listName -replace $specialCharacters,'') -replace " ",''))
	$outputFolders = ("{0}\{1}-{2}-Folders.csv" -f $outputFoldersPath, (($list.siteTitle -replace $specialCharacters,'') -replace " ",''), (($list.listName -replace $specialCharacters,'') -replace " ",''))

	$itemsQueue = New-Object System.Collections.Queue;
	try {
		$temp = Get-PnPListItem -List $pnpList  -PageSize $itemBatchSize -ErrorAction Stop | ForEach-Object{$itemsQueue.Enqueue($_)}
	}
	catch {
		Write-CatErrorCSV -funcCSVItem $list -funcAction "Retrieving Items" -funcException $_.Exception -funcDetails $_.ErrorDetails -funcStackTrace $_.ScriptStackTrace -funcFilePath $errorPath
		Continue listLoop
	}

	# Writing items to CSV
	If ($itemsQueue.Count -gt 0) {
		Write-Host (" (Item Count = {0})" -f $itemsQueue.Count)
		Write-Host ("   {0} - Writing All Items to CSV" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"))
		while($itemsQueue.Count -ne 0) {
			$itemsQueue.Dequeue() | ForEach-Object{
				$item = New-Object system.object
				$item | Add-Member -type NoteProperty -name "siteTitle" -value $list.siteTitle
				$item | Add-Member -type NoteProperty -name "siteURL" -value $list.siteURL
				$item | Add-Member -type NoteProperty -name "listName" -value $list.listName
				$item | Add-Member -type NoteProperty -name "listPath" -value $list.listPath
				$itemType = $_.FieldValues.FSObjType
				$fields = $_.FieldValues.GetEnumerator()

				# Looping through each field dynamically
				Foreach ($field in $fields) {
					$fieldKey = $field.Key
					$fieldValue = $field.Value
				
					# Different Output for folders or Items
					switch ($itemType) {
						0 { # List item / Document
							if(!($fieldKey.StartsWith('_')) -and !($fieldKey -in $itemExcludedColumns)) {					 
								# Pull in CreatedBy, LastModifiedBy, and LookupField Details or just create field.
								switch ($fieldKey) {
									Title {$item | add-member -type Noteproperty -name "Title" -value ((($fieldValue -replace $specialCharacters,'') -replace "`n",' ') -replace "`r",' ')}
									Description {$item | add-member -type Noteproperty -name "Description" -value ((($fieldValue -replace $specialCharacters,'') -replace "`n",' ') -replace "`r",' ')}
									Author {$item | add-member -type Noteproperty -name "itemCreatedBy" -value $fieldValue.LookupValue}
									Editor {$item | add-member -type Noteproperty -name "itemModifiedBy" -value $fieldValue.LookupValue}
									FileDirRef {$item | add-member -type Noteproperty -name "folderPath" -value $fieldValue}
									FileLeafRef {$item | add-member -type Noteproperty -name "fileName" -value $fieldValue}
									File_x0020_Size {$item | add-member -type Noteproperty -name "fileSizeBites" -value $fieldValue}
									{$fieldValue -like "*FieldLookupValue*"} {$item | add-member -type Noteproperty -name $fieldKey -value $fieldKey.LookupValue}
									Default {$item | add-member -type Noteproperty -name $fieldKey -value $fieldValue}
								}	
							}
						}
						1 { # Folder
							if(!($fieldKey.StartsWith('_')) -and !($fieldKey -in $folderExcludedColumns)) {					 
								# Pull in CreatedBy, LastModifiedBy, and LookupField Details or just create field.
								switch ($fieldKey) {
									Title {$item | add-member -type Noteproperty -name "Title" -value ((($fieldValue -replace $specialCharacters,'') -replace "`n",' ') -replace "`r",' ')}
									Description {$item | add-member -type Noteproperty -name "Description" -value ((($fieldValue -replace $specialCharacters,'') -replace "`n",' ') -replace "`r",' ')}
									Author {$item | add-member -type Noteproperty -name "folderCreatedBy" -value $fieldValue.LookupValue}
									Editor {$item | add-member -type Noteproperty -name "folderModifiedBy" -value $fieldValue.LookupValue}
									FileDirRef {$item | add-member -type Noteproperty -name "folderPath" -value $fieldValue}
									FileLeafRef {$item | add-member -type Noteproperty -name "folderName" -value $fieldValue}
									{$fieldValue -like "*FieldLookupValue*"} {$item | add-member -type Noteproperty -name $fieldKey -value $fieldKey.LookupValue}
									Default {$item | add-member -type Noteproperty -name $fieldKey -value $fieldValue}
								}	
							}
						}
						Default {Write-Host ("      Unknown FSObjType - {0}" -f $itemType) -ForegroundColor Red}
					}
				}

				# Exporting to CSV file
				switch ($itemType) {
					0 {$item | Export-Csv -Path $outputItems -NoTypeInformation -Append}
					1 {$item | Export-Csv -Path $outputFolders -NoTypeInformation -Append}
				}
				Remove-Variable item
			}
		}
		Remove-Variable itemsQueue
	}
	Else {
		Write-Host " (No Items Found)" -ForegroundColor Yellow
	}
	#endregion #>

	# Disconnecting from SharePoing
	Write-Host ("   {0} - Disconnecting From SharePoint Online site: {1}" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"), $list.siteTitle)
	Disconnect-PnPOnline
} 
#endregion

#region Completed Script
# Output Script Info
Complete-CatScript $scriptStart[0]

# Stop logging
Stop-Transcript
#endregion