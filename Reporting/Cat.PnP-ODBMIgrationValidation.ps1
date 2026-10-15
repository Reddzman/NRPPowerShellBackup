<#  Cat.PnP-ODBMigrationValidation.ps1

This script collects item counts from source & destination OneDrive Documents libraries. It iterates through the same CSV file format at the Cat.SG-OD4BMigrations.ps1 script. 
The intent is to be able to do high level validation across OneDrive migration jobs.

Dependencies
	-	PnP.PowerShell Modules
	-	General Functions - ..\General\Cat.GeneralFunctions.psm1
	-	Input CSV File - Example: ..\__Input Files\Samples\Reporting\Cat.PnP-ODBMigrationValidation-Client.csv
	-	Source & Destination Azure App Registrations are recommended but it the srcAppID & dstAppID are left blank it will use interactive login.

Call script with parameters
	cd "C:\GitHub\CatPowerShell\Reporting\"
	.\Cat.PnP-ODBMigrationValidation  -inputFile "Samples\Reporting\Cat.PnP-ODBMigrationValidation-Client.csv" -srcAppID "caf6b11a-cf03-416c-b0f0-67a09a8e30ee" -srcCertPath "C:\GitHub\NRPDev\AppRegistration\PowerShell App Registration.pfx" -srcTenant "nrpdev2" -dstAppID "caf6b11a-cf03-416c-b0f0-67a09a8e30ee" -dstCertPath "C:\GitHub\NRPDev\AppRegistration\PowerShell App Registration.pfx" -dstTenant "nrpdev2" -outputLocation "C:\GitHub\CatPowerShell\_Output Files" -itemLevelValidation $false -itemVariance 100 -itemBatchSize 10

Table of Contents
	1.	Snapin(s), Modules & Environment Connection(s)
	2.	Input & Output Files
	3.	Variables
	4.	Functions
		4a.	Connect to SharePoint Site
		4b.	Get User OneDrive Information
	5.	Main Script
		5b.	CSV Loop of OneDrive Accounts
#>

#region Script Parameters
Param(	[string] $inputFile = "{Teams input file path relative to _Input Files}", # IE. "sample\filename.csv"
		[string] $srcAppID = "{Client app registration ID from Azure}", # Leave Blank for interactive login
		[string] $srcCertPath = "{Path to client certificate file must be installed on local user}",# Leave Blank for interactive login
		[string] $srcTenant = "{Client tenant SharePoint prefix}", # IE. ClientTenant.sharepoint.com
		[string] $dstAppID = "{Client app registration ID from Azure}", # Leave Blank for interactive login
		[string] $dstCertPath = "{Path to client certificate file must be installed on local user}",# Leave Blank for interactive login
		[string] $dstTenant = "{Client tenant SharePoint prefix}", # IE. ClientTenant.sharepoint.com
		[string] $outputLocation = "{c:\path to output files location}", # A /clienttenant/datetime/ folder will be created for output files
		[boolean] $itemLevelValidation = $false, # If $true, the script will output all item level metadata for all accounts where library validation doesn't pass
		[int] $itemVariance = 50, # Acceptable # of files added since migration, Not % of files. IF Source Item Count < Destination Item Count & Variance is greater than this number, all item metadata will be output to CSV for that ODB.
		[int] $itemBatchSize = 4000 # If $itemLevelValidation = $true, this represents the batch size for the all items query
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
Start-Transcript -Path ("..\_Log Files\{0}.PnP-ODBMigrationValidation-{1}.log" -f $dstTenant, $scriptStart[1])
#endregion

#region 1.	Snapin(s), Modules & Environment Connection(s)
# Load MS / PnP PowerShell Modules
Get-Module -ListAvailable | Where-Object{($_.Name -like "*SharePoint*") -or ($_.Name -like "*PnP*") } | Remove-Module
Import-Module PnP.PowerShell
#endregion

#region 2.	Input & Output Files
# File Paths
$inPath = ("..\_Input Files\{0}" -f $inputFile)
$outputPath = ("{0}\{1}\{2}" -f $outputLocation, $dstTenant, $scriptStart[1])
$errorPath = ("{0}\errors.csv" -f $outputPath)
$itemCountPath = ("{0}\ItemCounts.csv" -f $outputPath)
$srcItemDetailPath = ("{0}\SrcItemDetails" -f $outputPath)
$srcFolderDetailPath = ("{0}\SrcFolderDetails" -f $outputPath)
$dstItemDetailPath = ("{0}\DstItemDetails" -f $outputPath)
$dstFolderDetailPath = ("{0}\DstFolderDetails" -f $outputPath)

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
if (($srcAppID) -or ($dstAppID)) {
	Write-Host
	Write-Host "User Input Variables" -ForegroundColor Cyan
}
if ($srcAppID) {
	$certPass = ConvertTo-SecureString -String (Read-Host -Prompt "   Enter certificate password for the SOURCE app registration") -AsPlainText -Force
	$srcCert = Get-PnPAzureCertificate -Path $srcCertPath -Password $certPass
	Remove-Variable certPass
	Remove-Variable srcCertPath
}
if ($dstAppID) {
	$certPass = ConvertTo-SecureString -String (Read-Host -Prompt "   Enter certificate password for the DESTINATION app registration") -AsPlainText -Force
	$dstCert = Get-PnPAzureCertificate -Path $dstCertPath -Password $certPass
	Remove-Variable certPass
	Remove-Variable dstCertPath
}

# Script Variables
$specialCharacters = "[~``!@#$%^&*()_-]" #(($pnpGroup.Description -replace $specialCharacters,'') -replace "`n",' ') -replace "`r",' '
$srcTenantDomain = ("{0}.onmicrosoft.com" -f $srcTenant)
$srcAdminUrl = ("https://{0}-admin.sharepoint.com" -f $srcTenant)
$dstTenantDomain = ("{0}.onmicrosoft.com" -f $dstTenant)
$dstAdminUrl = ("https://{0}-admin.sharepoint.com" -f $dstTenant)
$itemExcludedColumns = @("A2ODMountCount",
						"A2ODIsMountPoint",
						"A2ODRemoteItemSiteId",
						"A2ODRemoteItemWebId",
						"A2ODRemoteItemListId",
						"A2ODRemoteItemUniqueId",
						"A2ODExtendedMetadata",
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
						"A2ODIsMountPoint",
						"A2ODRemoteItemSiteId",
						"A2ODRemoteItemWebId",
						"A2ODRemoteItemListId",
						"A2ODRemoteItemUniqueId",
						"A2ODExtendedMetadata",
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

#region Output Parameters & Variables
Write-Host
Write-Host "Script Parameters" -ForegroundColor Cyan
Write-Host ("   Input File: {0}" -f $inputFile)
if ($srcAppID) {
	Write-Host ("   Source App ID: {0}" -f $srcAppID)
}
Write-Host ("   Source Tenant Prefix: {0}" -f $srcTenant)
if ($dstAppID) {
	Write-Host ("   Destination App ID: {0}" -f $dstAppID)
}
Write-Host ("   Destination Tenant Prefix: {0}" -f $dstTenant)
Write-Host ("   Output Location: {0}" -f $outputLocation)
Write-Host ("   Item level Validation: {0}" -f $itemLevelValidation)
Write-Host ("   Item Acceptable Variance: {0}" -f $itemVariance)
Write-Host ("   Item Batch Size: {0}" -f $itemBatchSize)
Write-Host
Write-Host "Output Variables" -ForegroundColor Cyan
Write-Host ("   Output Path: {0}" -f $outputPath)
Write-Host ("   Error Path: {0}" -f $errorPath)
Write-Host ("   Item Count Path: {0}" -f $itemCountPath)
Write-Host ("   Source Item Details Path: {0}" -f $srcItemDetailPath)
Write-Host ("   Source Folder Details Path: {0}" -f $srcFolderDetailPath)
Write-Host ("   Destination Item Details Path: {0}" -f $dstItemDetailPath)
Write-Host ("   Destination Folder Details Path: {0}" -f $dstFolderDetailPath)
if (($srcAppID) -or ($dstAppID)) {
	Write-Host
	Write-Host "User Inputs" -ForegroundColor Cyan
}
if ($srcAppID) {
	Write-Host ("   Source Certificate Thumbprint: {0}" -f "Only output for troubleshooting", $srcCert.Thumbprint)
}
if ($dstAppID) {
	Write-Host ("   Destination Certificate Thumbprint: {0}" -f "Only output for troubleshooting", $dstCert.Thumbprint)
}
Write-Host
Write-Host "Script Variables" -ForegroundColor Cyan
Write-Host ("   Special Characters to Remove: {0}" -f $specialCharacters)
Write-Host ("   Source Tenant Domain: {0}" -f $srcTenantDomain)
Write-Host ("   Source Admin URL: {0}" -f $srcAdminUrl)
Write-Host ("   Destination Tenant Domain: {0}" -f $dstTenantDomain)
Write-Host ("   Destination Admin URL: {0}" -f $dstAdminUrl)
<#Write-Host
Write-Host "Item Excluded Columns" -ForegroundColor Cyan
Write-Host $itemExcludedColumns
Write-Host
Write-Host "Folder Excluded Columns" -ForegroundColor Cyan
Write-Host $itemExcludedColumns #>
#endregion
#endregion

#region 4.	Functions
	#region 4a.	Connect to SharePoint Site
	# Example Cmd: Connect-CatSharePointSite -funItem $odb -funURL $srcAdminUrl -funDomain $srcTenantDomain -funAppID $srcAppID -funThumb $srcCert.Thumbprint
	function Connect-CatSharePointSite { 
		param(
			$conItem,
			[string]$conURL,
			[string]$conDomain,
			[string]$conAppID,
			[string]$conThumb
		)

		try {
			if ($conAppID) {
				Connect-PnPOnline $conURL -tenant $conDomain -ClientId $conAppID -Thumbprint $conThumb -ValidateConnection -ErrorAction Stop # PnP Modules w/ App Reg
			} else {
				Connect-PnPOnline -Url $conURL -Interactive -ValidateConnection -ErrorAction Stop # PnP Modules w/o App Reg
			} 
		}
		catch {
			Write-CatErrorCSV -funcCSVItem $conItem -funcAction ("Connecting to SharePoint - {0}" -f $conURL) -funcException $_.Exception -funcDetails $_.ErrorDetails -funcStackTrace $_.ScriptStackTrace -funcFilePath $errorPath
			Continue odbLoop
		}

		# Clean Up variables
		Remove-Variable conItem
		Remove-Variable conURL
		Remove-Variable conDomain
		Remove-Variable conAppID
		Remove-Variable conThumb
	}
	#endregion

	#region 4b.	Get User OneDrive Information
	#Example Cmd: $srcLibrary = Get-CatOneDriveData -odbItem $odb -odbURL $srcAdminUrl -odbDomain $srcTenantDomain -odbAppID $srcAppID -odbThumb $srcCert.Thumbprint -odbUPN $odb.sourceUserEmail
	function Get-CatOneDriveData { 
		param(
			$odbItem,
			[string]$odbURL,
			[string]$odbDomain,
			[string]$odbAppID,
			[string]$odbThumb,
			[string]$odbUPN
		)

		# Connecting To SharePoint Administration
		Connect-CatSharePointSite -conItem $odbItem -conURL $odbURL -conDomain $odbDomain -conAppID $odbAppID -conThumb $odbThumb

		# Get OneDrive Site URL
		Write-Host ("      {0} - Retrieving Source site URL" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"))
		try {
			$odbProfile = Get-PnPUserProfileProperty $odbUPN -ErrorAction Stop
		}
		catch {
			Write-CatErrorCSV -funcCSVItem $odbItem -funcAction "Retrieving Source site URL" -funcException $_.Exception -funcDetails $_.ErrorDetails -funcStackTrace $_.ScriptStackTrace -funcFilePath $errorPath
			Continue odbLoop
		}
		
		# Disconnect from SPO Admin
		Disconnect-PnPOnline

		# region Connecting To User OneDrive
		Connect-CatSharePointSite -conItem $odbItem -conURL $odbProfile.PersonalUrl -conDomain $odbDomain -conAppID $odbAppID -conThumb $odbThumb

		# Get Document Library Information
		Write-Host ("      {0} - Retrieving OneDrive Library" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"))
		try {
			$odbLibrary = Get-PnPList -Identity "Documents" -Includes ParentWeb -ErrorAction Stop
		}
		catch {
			Write-CatErrorCSV -funcCSVItem $odbItem -funcAction "Retrieving OneDrive Library" -funcException $_.Exception -funcDetails $_.ErrorDetails -funcStackTrace $_.ScriptStackTrace -funcFilePath $errorPath
			Continue odbLoop
		}

		# Disconnect from SPO Admin
		Disconnect-PnPOnline

		Return $odbLibrary

		# Clean up variables
		Remove-Variable odbItem
		Remove-Variable odbURL
		Remove-Variable odbDomain
		Remove-Variable odbAppID
		Remove-Variable odbThumb
		Remove-Variable odbUPN
		Remove-Variable odbProfile
		Remove-Variable odbLibrary
	}
	#endregion

	#region 4c. Get Item Details
	function Get-CatItemDetails { 
		param(
			$detItem,
			[string]$detTarget,
			[string]$detUPN,
			[string]$detItemPath,
			[string]$detFolderPath,
			[string]$detURL,
			[string]$detDomain,
			[string]$detAppID,
			[string]$detThumb,
			[string]$detList			
		)

		Write-Host ("         {0} - Retrieving {1} All Items" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"),$detTarget) -NoNewline
		$outputItems = ("{0}\{1}-ItemDetails.csv" -f $detItemPath, ($detUPN -replace $specialCharacters,''))
		$outputFolders = ("{0}\{1}-FolderDetails.csv" -f $detFolderPath, ($detUPN -replace $specialCharacters,''))

		# Connecting To User OneDrive
		Connect-CatSharePointSite -conItem $detItem -conURL $detURL -conDomain $detDomain -conAppID $detAppID -conThumb $detThumb
		
		# Retrieving all items
		$itemsQueue = New-Object System.Collections.Queue;
		try {
			$temp = Get-PnPListItem -List $detList -PageSize $itemBatchSize -ErrorAction Stop | ForEach-Object{$itemsQueue.Enqueue($_)}
		}
		catch {
			Write-CatErrorCSV -funcCSVItem $detItem -funcAction ("Retrieving {0} Items" -f $detTarget) -funcException $_.Exception -funcDetails $_.ErrorDetails -funcStackTrace $_.ScriptStackTrace -funcFilePath $errorPath
			Continue odbLoop
		}

		# Writing items to CSV
		If ($itemsQueue.Count -gt 0) {
			Write-Host (" (Item Count = {0})" -f $itemsQueue.Count)
			Write-Host ("         {0} - Writing All {1} Items to CSV" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"),$detTarget)
			while($itemsQueue.Count -ne 0) {
				$itemsQueue.Dequeue() | ForEach-Object{
					$item = New-Object system.object
					$item | Add-Member -MemberType NoteProperty -Name "sourceUserEmail" -Value $odb.sourceUserEmail
					$item | Add-Member -MemberType NoteProperty -Name "destUserEmail" -Value $odb.destUserEmail
					$item | Add-Member -type NoteProperty -name "siteURL" -value $detURL
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

						# Cleaning up variables
						Remove-Variable fieldKey
						Remove-Variable fieldValue
					}

					# Exporting to CSV file
					switch ($itemType) {
						0 {$item | Export-Csv -Path $outputItems -NoTypeInformation -Append}
						1 {$item | Export-Csv -Path $outputFolders -NoTypeInformation -Append}
					}

					# Cleaning up variables
					#Remove-Variable item
					Remove-Variable itemType
					Remove-Variable fields
				}
			}
		}
		Else {
			Write-Host " (No Items Found)" -ForegroundColor Yellow
		}

		# Cleaning up variables
		Remove-Variable detItem
		Remove-Variable detUPN
		Remove-Variable detItemPath
		Remove-Variable detFolderPath
		Remove-Variable detURL
		Remove-Variable detDomain
		Remove-Variable detAppID
		Remove-Variable detThumb
		Remove-Variable detList
		Remove-Variable itemsQueue
		Remove-Variable temp
	}
	#endregion
#endregion
    
#region 5.	Main Script
#region Created directories for additional reports
Write-Host
Write-Host ("{0} - Creating Report Directories" -f (Get-Date -Format "yy-MM-dd HH:mm:ss")) -ForegroundColor Cyan
New-Item -Path $srcItemDetailPath -ItemType Directory
New-Item -Path $srcFolderDetailPath -ItemType Directory
New-Item -Path $dstItemDetailPath -ItemType Directory
New-Item -Path $dstFolderDetailPath -ItemType Directory
#endregion

#region 5b.	CSV Loop of OneDrive Accounts
Write-Host
Write-Host "Scanning OneDrive Accounts (Documents library only)" -ForegroundColor Cyan
:odbLoop foreach ($odb in $inFile) {
	Write-Host
	Write-Host ("   {0} - Scanning OneDrive: {1}" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"), $odb.destUserEmail) -ForegroundColor Magenta

	# Retreiving Source Library
	$srcLibrary = Get-CatOneDriveData -odbItem $odb -odbURL $srcAdminUrl -odbDomain $srcTenantDomain -odbAppID $srcAppID -odbThumb $srcCert.Thumbprint -odbUPN $odb.sourceUserEmail

	# Retreiving Destination Library
	$dstLibrary = Get-CatOneDriveData -odbItem $odb -odbURL $dstAdminUrl -odbDomain $dstTenantDomain -odbAppID $dstAppID -odbThumb $dstCert.Thumbprint -odbUPN $odb.destUserEmail
	If (($srcLibrary.ItemCount -lt $dstLibrary.ItemCount) -and ($srcLibrary.ItemCount -ne 0)) {$curVariance = $dstLibrary.ItemCount % $srcLibrary.ItemCount} Else {$curVariance = $null}

	# Output OneDrive Item Counts
	Write-Host ("      {0} - Writing Item Counts to CSV" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"))
	$record = New-Object system.object
	$record | Add-Member -MemberType NoteProperty -Name "sourceUserEmail" -Value $odb.sourceUserEmail
	$record | Add-Member -MemberType NoteProperty -Name "destUserEmail" -Value $odb.destUserEmail
    $record | Add-Member -MemberType NoteProperty -Name "sourceItemCount" -Value $srcLibrary.ItemCount
    $record | Add-Member -MemberType NoteProperty -Name "destItemCount" -Value $dstLibrary.ItemCount
	$record | Add-Member -MemberType NoteProperty -Name "currentVariance" -Value $curVariance
    $record | Export-Csv -Path $itemCountPath -NoTypeInformation -Append
	Remove-Variable record

	# Determine if additional Validation is necessary
	If (($itemLevelValidation) -and ($srcLibrary.ItemCount -lt $dstLibrary.ItemCount) -and ($curVariance -gt $itemVariance)) {
		Write-Host ("      {0} - Extracting Item Level Detail - Source Count: {1}  Destination Count: {2}" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"),$srcLibrary.ItemCount,$dstLibrary.ItemCount)
		
		# Ouput Source OneDrive Metadata
		Get-CatItemDetails -detItem $odb -detTarget "Source" -detUPN $odb.sourceUserEmail -detItemPath $srcItemDetailPath -detFolderPath $srcFolderDetailPath -detURL $srcLibrary.ParentWeb.URL -detDomain $srcTenantDomain -detAppID $srcAppID -detThumb $srcCert.Thumbprint -detList $srcLibrary.Title
		
		# Ouput Destination OneDrive Metadata
		Get-CatItemDetails -detItem $odb -detTarget "Destination" -detUPN $odb.destUserEmail -detItemPath $dstItemDetailPath -detFolderPath $dstFolderDetailPath -detURL $dstLibrary.ParentWeb.URL -detDomain $dstTenantDomain -detAppID $dstAppID -detThumb $dstCert.Thumbprint -detList $dstLibrary.Title
	}

	# Cleaning up variables
	Remove-Variable srcLibrary
	Remove-Variable dstLibrary
} 
#endregion
#endregion

#region Completed Script
# Cleaning up variables
Remove-Variable inputFile
Remove-Variable srcAppID
Remove-Variable srcTenant
Remove-Variable dstAppID
Remove-Variable dstTenant
Remove-Variable outputLocation
Remove-Variable inPath
Remove-Variable outputPath
Remove-Variable errorPath
Remove-Variable itemCountPath
Remove-Variable srcCert
Remove-Variable dstCert
Remove-Variable specialCharacters
Remove-Variable srcTenantDomain
Remove-Variable srcAdminUrl
Remove-Variable dstTenantDomain
Remove-Variable dstAdminUrl

# Output Script Info
Complete-CatScript $scriptStart[0]

# Stop logging
Stop-Transcript
#endregion