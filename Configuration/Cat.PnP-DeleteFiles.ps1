<#  Cat.PnP-DeleteFiles.ps1

This script iterates through a list of file and folder URLs and checks to see if the file exists.  If it does exist it will report it by outputing to CSV.
The option to delete files exists as a parameter.

Dependencies
	-	PnP.PowerShell Modules
	-	Input CSV File - Example: _Input Files\Samples\Configuration\Cat.PnP-DeleteFiles-Test.csv

*** Retrieving the file works with Full Path, Tenant/Server Relative Path, and Site Relative Path.  Deleting the file only works with Site Relative Path. ***

Call script with parameters
	cd "C:\GitHub\CatPowerShell\Configuration\"
	.\Cat.PnP-DeleteFiles.ps1 -inputFile "Samples\Configuration\Cat.PnP-DeleteFiles-Test.csv" -clientAppID "caf6b11a-cf03-416c-b0f0-67a09a8e30ee" -clientCertPath "C:\GitHub\NRPDev\AppRegistration\PowerShell App Registration.pfx" -clientTenant "nrpdev2" -outputLocation "C:\GitHub\CatPowerShell\_Output Files" -deleteFiles $false -recycleFiles $true

Table of Contents
	1.	Snapin(s), Modules & Environment Connection(s)
	2.	Input & Output Files
	3.	Variables
	4.	Main Script
#>

#region Script Parameters
Param(	[string] $inputFile = "{Teams input file path relative to _Input Files}", # Location from _Input Files folder - IE. "sample\filename.csv" or "filename.csv"
		[string] $clientAppID = "{Client app registration ID from Azure}", # Leave Blank for interactive login
		[string] $clientCertPath = "{Path to client certificate file must be installed on local user}",# Leave Blank for interactive login
		[string] $clientTenant = "{Client tenant SharePoint prefix}", # IE. ClientTenant.sharepoint.com
		[string] $outputLocation = "{c:\path to output files location}", # A /clienttenant/datetime/ folder will be created for output files
		[boolean] $deleteFiles = $false, # If True script will delete the files it finds
		[boolean] $recycleFiles = $true # If False script will PERMINANTLY delete the files it finds
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
Start-Transcript -Path ("..\_Log Files\{0}.PnP-DeleteFiles-{1}.log" -f $clientTenant, $scriptStart[1])
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
$existingFilesReport = ("{0}\ExistingFiles.csv" -f $outputPath)

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
If (($clientAppID) -or ($deleteFiles)) {
	Write-Host
	Write-Host "User Input Variables" -ForegroundColor Cyan
}
If ($deleteFiles) {
	$deleteProceed = Read-Host "   Are you sure you want to proceed with deleting files in the input file (Yes or No)"

	# if user chose to proceed
	if ($deleteProceed.ToLower() -eq "yes") {
		$deleteFiles = $true
	}
	else {
		$deleteFiles = $false
	}
}
If (($deleteFiles) -and (-not $recycleFiles)) {
	Write-Host "   *** Bypassing the recycle bin will perminantly delete files. There is no way to recover from this deletion. ***" -ForegroundColor Red
	$recycleProceed = Read-Host "   Are you sure you want to proceed with PERMINANTLY deleting files in the input file (Yes or No)"

	# if user chose to proceed
	if ($recycleProceed.ToLower() -eq "yes") {
		$recycleFiles = $false
	}
	else {
		$recycleFiles = $true
	}
}
if ($clientAppID) {
	$certPass = ConvertTo-SecureString -String (Read-Host -Prompt "   Enter certificate password for the app registration") -AsPlainText -Force
	$cert = Get-PnPAzureCertificate -Path $clientCertPath -Password $certPass
	Remove-Variable certPass
	Remove-Variable clientCertPath
}

# Script Variables
$specialCharacters = "[~``!@#$%^&*()_-]" #(($pnpGroup.Description -replace $specialCharacters,'') -replace "`n",' ') -replace "`r",' '
$tenantDomain = ("{0}.onmicrosoft.com" -f $clientTenant)
$spoAdminUrl = ("https://{0}-admin.sharepoint.com" -f $clientTenant)
$fileQuery = `
"<View Scope='Recursive'>
	<Query>
		<Where>
			<Contains>
				<FieldRef Name='FileRef' />
				<Value Type='Text'>{filePath}</Value>
			</Contains>
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
Write-Host ("   Output Location: {0}" -f $outputLocation)
Write-Host ("   Delete Files: {0}" -f $deleteFiles)
Write-Host ("   Recycle Files: {0}" -f $recycleFiles)
Write-Host
Write-Host "Output Variables" -ForegroundColor Cyan
Write-Host ("   Output Path: {0}" -f $outputPath)
Write-Host ("   Error Path: {0}" -f $errorPath)
Write-Host ("   Existing Files Path: {0}" -f $existingFilesReport)
If (($clientAppID) -or ($deleteFiles)) {
	Write-Host
	Write-Host "User Inputs" -ForegroundColor Cyan
}
If ($deleteFiles) {
	Write-Host ("   Proceed with deleting files: {0}" -f $deleteProceed)
}
If (($deleteFiles) -and (-not $recycleFiles)) {
	Write-Host ("   Proceed with bypassing the recycle bin: {0}" -f $recycleProceed)
}
if ($clientAppID) {
	Write-Host ("   Certificate Thumbprint: {0}" -f "Only output for troubleshooting", $cert.Thumbprint)
}
Write-Host
Write-Host "Script Variables" -ForegroundColor Cyan
Write-Host ("   Special Characters to Remove: {0}" -f $specialCharacters)
Write-Host ("   Microsoft Tenant: {0}" -f $tenantDomain)
Write-Host ("   SPO admin URL: {0}" -f $spoAdminUrl)
Write-Host
Write-Host "File Query" -ForegroundColor Cyan
Write-Host $fileQuery
#endregion

# Cleaning up variables
If ($deleteFiles) {Remove-Variable deleteProceed}
If (($deleteFiles) -and (-not $recycleFiles)) {Remove-Variable recycleProceed}
#endregion
    
#region 4.	Main Script
#region Created directories for additional reports
Write-Host
Write-Host ("{0} - Creating Report Directories" -f (Get-Date -Format "yy-MM-dd HH:mm:ss")) -ForegroundColor Cyan
New-Item -Path $outputPath -ItemType Directory
#endregion

#region Looping through files in CSV
Write-Host
Write-Host ("Checking for existing files/folders: {0}" -f $inputFile) -ForegroundColor Cyan
:fileLoop foreach ($file in $inFile){
	Write-Host
	Write-Host ("   {0} - Checking if file/folder exists: {1}" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"), $file.name) -ForegroundColor Magenta
	if ($file.isFolder -eq "TRUE") {$isFolder = $true} Else {$isFolder = $false}

	#region Connecting To SP Online
	Write-Host ("      {0} - Connecting to SharePoint Online Admin site: {1}" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"), $file.siteTitle)
	try {
		if ($clientAppID) {
			Connect-PnPOnline $file.siteURL -tenant $tenantDomain -ClientId $clientAppID -Thumbprint $cert.Thumbprint -ValidateConnection -ErrorAction Stop # PnP Modules w/ App Reg
		} else {
			Connect-PnPOnline -Url $file.siteURL -Interactive -ValidateConnection -ErrorAction Stop # PnP Modules w/o App Reg
		} 
	}
	catch {
		Write-CatErrorCSV -funcCSVItem $file -funcAction "Connecting to SharePoint" -funcException $_.Exception -funcDetails $_.ErrorDetails -funcStackTrace $_.ScriptStackTrace -funcpath $errorPath
		Continue fileLoop
	}
	#endregion

	#region Retrieving File
	Write-Host ("      {0} - Retrieving File" -f (Get-Date -Format "yy-MM-dd HH:mm:ss")) -NoNewline
	try {
		# Get file based on URL
		$tempPath = $file.path -replace "%20"," "
		if ($isFolder) {
			$tempFile = Get-PnPFolder -Url $tempPath -Includes TimeCreated, TimeLastModified, ListItemAllFields -ErrorAction Stop
		} else {
			$tempFile = Get-PnPFile -Url $tempPath -ErrorAction Stop
		}
		Write-Host " - File/folder Found" -ForegroundColor Yellow
	}
	catch {
		Write-host " - File/folder Not Found" -ForegroundColor Green
		Continue fileLoop
	}
	#endregion

	# If file was found, output information
	If ($tempFile) {
		#region Write information to CSV
		Write-Host ("      {0} - Writing File/folder Information to CSV" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"))
		$fileOutput = New-Object system.object
		$fileOutput | Add-Member -MemberType NoteProperty -Name "siteTitle" -Value $file.siteTitle
		$fileOutput | Add-Member -MemberType NoteProperty -Name "siteURL" -Value $file.siteURL
		$fileOutput | Add-Member -MemberType NoteProperty -Name "path" -Value $file.path
		$fileOutput | Add-Member -MemberType NoteProperty -Name "name" -Value $file.name
		$fileOutput | Add-Member -MemberType NoteProperty -Name "isFolder" -Value $file.isFolder
		$fileOutput | Add-Member -MemberType NoteProperty -Name "deletefile" -Value $deleteFiles
		$fileOutput | Add-Member -MemberType NoteProperty -Name "recycleFiles" -Value $recycleFiles
		$fileOutput | Add-Member -MemberType NoteProperty -Name "fileCreatedDate" -Value $tempFile.TimeCreated
		$fileOutput | Add-Member -MemberType NoteProperty -Name "fileModifiedDate" -Value $tempFile.TimeLastModified
		if ($isFolder) {
			$fileOutput | Add-Member -MemberType NoteProperty -Name "fileAuthorName" -Value $tempFile.ListItemAllFields.FieldValues.Author.LookupValue
			$fileOutput | Add-Member -MemberType NoteProperty -Name "fileAuthorEmail" -Value $tempFile.ListItemAllFields.FieldValues.Author.email
			$fileOutput | Add-Member -MemberType NoteProperty -Name "fileModifiedByName" -Value $tempFile.ListItemAllFields.FieldValues.Editor.LookupValue
			$fileOutput | Add-Member -MemberType NoteProperty -Name "fileModifiedByEmail" -Value $tempFile.ListItemAllFields.FieldValues.Editor.email
		} else {
			$fileOutput | Add-Member -MemberType NoteProperty -Name "fileAuthorName" -Value $tempFile.Author.Title
			$fileOutput | Add-Member -MemberType NoteProperty -Name "fileAuthorEmail" -Value $tempFile.Author.Email
			$fileOutput | Add-Member -MemberType NoteProperty -Name "fileModifiedByName" -Value $tempFile.ModifiedBy.Title
			$fileOutput | Add-Member -MemberType NoteProperty -Name "fileModifiedByEmail" -Value $tempFile.ModifiedBy.Email
		}
		$fileOutput | Export-Csv -Path $existingFilesReport -NoTypeInformation -Append
		Remove-Variable fileOutput
		#endregion

		#region Delete File
		# If $deleteFiles is true delete file
		if ($deleteFiles) {
			try {
				$tempQuery = $fileQuery.Replace("{filePath}",$file.path.Replace("%20"," "))
				$tempList = get-PnPList $file.listName
				$tempItem = Get-PnPListItem -List $tempList -Query $tempQuery
				if ($isFolder) {$tempType = "Folder"} Else {$tempType = "File"}
				if ($isFolder) {$tempID = $tempFile.ListItemAllFields.Id} else {$tempID = $tempItem.Id}
				Write-Host ("      {0} - Deleting {1}" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"), $tempType) -NoNewline
				$tempDelete = Remove-PnPListItem -List $tempList -Identity $tempID -Recycle:$recycleFiles -Force
				Write-Host " - Success" -ForegroundColor Green
			}
			catch {
				Write-Host " - Failed" -ForegroundColor Red
				Write-CatErrorCSV -funcCSVItem $file -funcAction "Deleting File/Folder" -funcException $_.Exception -funcDetails $_.ErrorDetails -funcStackTrace $_.ScriptStackTrace -funcFilePath $errorPath
			}

			# Cleaning up Variables
			Remove-Variable tempQuery
			If ($tempList) {Remove-Variable tempList}
			If ($tempItem) {Remove-Variable tempItem}
			If ($tempDelete) {Remove-Variable tempDelete}
			If ($tempType) {Remove-Variable tempType}
			If ($tempID) {Remove-Variable tempID}
		}
		#endregion

		# Cleaning up Variables
		If ($tempFile) {Remove-Variable tempFile}
		If ($tempPath) {Remove-Variable tempPath}
	}
} 
#endregion 

#endregion

#region Completed Script
# Clean up variables
Remove-Variable inputFile
Remove-Variable clientAppID
Remove-Variable clientTenant
Remove-Variable outputLocation
Remove-Variable deleteFiles
Remove-Variable recycleFiles
Remove-Variable inPath
Remove-Variable outputPath
Remove-Variable errorPath
Remove-Variable existingFilesReport
Remove-Variable specialCharacters
Remove-Variable tenantDomain
Remove-Variable spoAdminUrl
if ($clientAppID) {Remove-Variable cert}
Remove-Variable inFile

# Output Script Info
Complete-CatScript $scriptStart[0]

# Stop logging
Stop-Transcript
#endregion