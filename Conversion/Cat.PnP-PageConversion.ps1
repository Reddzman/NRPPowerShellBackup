<#  Cat.PnP-PageConversion.ps1 Description

This script script goes through a CSV file with source and destination pages libraries (Pages or Site Pages) and converts
them to a modern sharepoint page in SharePoint. The destination must be a SharePoint Communication Site. The log file will
be created as a modern page in the destination site pages library.

Tested Scenarios:
	-	SP19 to SPO
	-	SP16 to SPO

Dependencies
	-	PnP.PowerShell Modules
	-	SharePoint PnP Modernization Framework
	-	General Functions - ..\General\Cat.GeneralFunctions.psm1
	-	Input CSV File - Example: ..\_Input FIles\Testing\Cat.Tech-Action-Test.csv

Call script with parameters
	cd "C:\GitHub\PowerShell\Conversion"
	.\Cat.PnP-PageConversion.ps1 -inputFile "Samples\Conversion\Cat.PageConversion-Test.csv" -clientAppID "" -clientCertPath "" -clientTenant "nrpdev2" -overwrite $true

Table of Contents
	1.	Snapin(s), Modules & Environment Connection(s)
	2.	Input & Output Files
	3.	Variables
	4.	Main Script
	5.	CSV Loop Example
	6.	...
#>

#region Script Parameters
Param(	[string] $inputFile = "{Teams input file path relative to _Input Files}", # Location from _Input Files folder - IE. "sample\filename.csv" or "filename.csv"
		[string] $clientAppID = $null, # Leave Blank for interactive login
		[string] $clientCertPath = "{Path to client certificate file must be installed on local user}",# Leave Blank for interactive login
		[string] $clientTenant = "{Client tenant SharePoint prefix}",
		[bool] $overwrite = $true
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
Start-Transcript -Path ("..\_Log Files\{0}.PnP-PageConversion-{1}.log" -f $clientTenant, $scriptStart[1])
#endregion Start Date / Time

#region 1.	Snapin(s), Modules & Environment Connection(s)
# Load MS / PnP PowerShell Modules
Get-Module -ListAvailable | Where-Object{($_.Name -like "*SharePoint*") -or ($_.Name -like "*PnP*") } | Remove-Module
Import-Module PnP.PowerShell
#endregion Snapin(s), Modules & Environment Connection(s)

#region 2.	Input & Output Files
# File Paths
$inPath = ("..\_Input Files\{0}" -f $inputFile)
$errorPath = "..\_Output Files\errors.csv"


# Import Input File(s)
# Load CSV file
If (Test-Path ($inPath)){
	$inFile = Import-Csv $inPath -Encoding Unicode
}
else {
	Write-Host
	throw ("The input file does not exist: {0}" -f $inPath)	
}
#endregion Input & Output Files

#region 3.	Variables
# User Inputs
Write-Host
Write-Host "User Input Variables" -ForegroundColor Cyan
$srcCreds = Get-Credential -Message "Enter credentials for the SOURCE SharePoint on premises environment:" 
If (($clientAppID)) {
	$certPass = ConvertTo-SecureString -String (Read-Host -Prompt "   Enter certificate password for the app registration") -AsPlainText -Force
	$cert = Get-PnPAzureCertificate -Path $clientCertPath -Password $certPass
	Remove-Variable certPass
	Remove-Variable clientCertPath
}

# Script Variables
$tenantDomain = ("{0}.onmicrosoft.com" -f $clientTenant)
$spoAdminUrl = ("https://{0}-admin.sharepoint.com" -f $clientTenant)
$libraryCount = 5000
$pageSkip = @("PageNotFoundError.aspx","Default.aspx","home.aspx","How To Use This Library.aspx") # Review for each client, May have a use case for copying default.aspx or home.aspx

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
Write-Host ("   Error Path: {0}" -f $errorPath)
Write-Host
Write-Host "User Inputs" -ForegroundColor Cyan
Write-Host ("   Source Credentials: {0}" -f $srcCreds.UserName)
If (($clientAppID)) {
	Write-Host ("   Certificate Thumbprint: {0}" -f "Only output for troubleshooting", $cert.Thumbprint)
}
Write-Host
Write-Host "Script Variables" -ForegroundColor Cyan
Write-Host ("   Microsoft Tenant: {0}" -f $tenantDomain)
Write-Host ("   SPO admin URL: {0}" -f $spoAdminUrl)
Write-Host ("   itemCount: {0}" -f $libraryCount)
Write-Host
Write-Host "Pages to be skipped" -ForegroundColor Cyan
$pageSkip
#endregion Output Parameters & Variables
#endregion
    
#region 4.	Main Script
#region Iterating through pages libraries
Write-Host
Write-Host ("Running Page Conversion for Pages libraries: {0}" -f $inputFile) -ForegroundColor Cyan
:libraryLoop foreach ($library in $inFile){
	Write-Host
	Write-Host ("   {0} - Converting - {1}/{2}" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"), $library.srcSiteURL, $library.srcLibrary) -ForegroundColor Magenta

	#region 5.	Connecting to Destination (SPO)
	Write-Host ("      {0} - Connecting to Destination - {1}" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"), $library.destSiteURL) -NoNewline
	try {
		$destConnection = Connect-PnPOnline $library.destSiteURL -tenant $tenantDomain -ClientId $clientAppID -Thumbprint $cert.Thumbprint -ReturnConnection -ErrorAction Stop # PnP Modules w/ App Reg
		Write-Host " - Success" -ForegroundColor Green
	}
	catch {
		Write-Host " - Failed" -ForegroundColor Red
		Write-CatErrorCSV -funcCSVItem $library -funcAction "Connecting to Destination" -funcException $_.Exception -funcDetails $_.ErrorDetails -funcStackTrace $_.ScriptStackTrace -funcFilePath $errorPath
		Continue libraryLoop
	}	
	#endregion

	#region 6.	Connecting to Source (SP On Premises)
	Write-Host ("      {0} - Connecting to Source - {1}" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"), $library.srcSiteURL) -NoNewline
	try {
		Connect-PnPOnline -Url $library.srcSiteURL -TransformationOnPrem -Credentials $srcCreds -ErrorAction Stop
		Write-Host " - Success" -ForegroundColor Green
	}
	catch {
		Write-Host " - Failed" -ForegroundColor Red
		Write-CatErrorCSV -funcCSVItem $library -funcAction "Connecting to Source" -funcException $_.Exception -funcDetails $_.ErrorDetails -funcStackTrace $_.ScriptStackTrace -funcFilePath $errorPath
		Continue libraryLoop
	}
	#endregion

	#region Iterating through all pages in library
	try {
		$srcPages = Get-PnPListItem -List $library.srcLibrary -PageSize $libraryCount -Fields ContentTypeId,ClientSideApplicationId | Where-Object {$_.FileSystemObjectType -ne "Folder"}
	}
	catch {
		Write-CatErrorCSV -funcCSVItem $library -funcAction "Retrieving Source Pages" -funcException $_.Exception -funcDetails $_.ErrorDetails -funcStackTrace $_.ScriptStackTrace -funcFilePath $errorPath
		Continue libraryLoop
	}
	:pageLoop foreach($page in $srcPages) {
		# Page variables
		$pageName = $page.FieldValues["FileLeafRef"]
		$pageCTId = $page.FieldValues["ContentTypeId"]
		$pageAppId = $page.FieldValues["ClientSideApplicationId"]

		<# Testing Variable Output
		Write-Host "         Name: $pageName" -ForegroundColor DarkGray
		Write-Host "         Content Type ID: $pageCTId" -ForegroundColor DarkGray
		Write-Host "         Client Side App ID: $pageAppId" -ForegroundColor DarkGray 
		#>

		# Determine if Page is Modern or should be skipped
		if(($pageAppId -ne "b6917cb1-93a0-4b97-a84d-7cf49975D4EC") -and (-not($pageName -in $pageSkip))) {
			Write-Host ("      {0} - Converting page: {1}" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"), $pageName)

			#region 7.	Converting page based on content type
			# Determining if page is a publishing page; May want to use a mapping file if there is a specific format to the pages.		
			if ($pageCTId -like "0x010100C5*" -or $pageCTId -like "0x01010007*")  {
				try {
					$pageTemp = ConvertTo-PnPPage -Identity $pageName `
						-Library $library.srcLibrary `
						-TakeSourcePageName:$true `
						-CopyPageMetadata `
						-KeepPageCreationModificationInformation `
						-TargetConnection $destConnection `
						-PublishingPage `
						-Overwrite:$overwrite `
						-LogType SharePoint `
						-LogSkipFlush
				}
				catch {
					Write-CatErrorCSV -funcCSVItem $library -funcAction "Converting Publishing Page" -funcException $_.Exception -funcDetails $_.ErrorDetails -funcStackTrace $_.ScriptStackTrace -funcFilePath $errorPath
					Continue pageLoop
				}
			} else {
				try {
					$pageTemp = ConvertTo-PnPPage -Identity $pageName `
						-Library $library.srcLibrary `
						-TakeSourcePageName:$true `
						-CopyPageMetadata `
						-KeepPageCreationModificationInformation `
						-TargetConnection $destConnection `
						-Overwrite:$overwrite `
						-LogType SharePoint `
						-LogSkipFlush
				}
				catch {
					Write-CatErrorCSV -funcCSVItem $library -funcAction "Converting Non-publishing pages" -funcException $_.Exception -funcDetails $_.ErrorDetails -funcStackTrace $_.ScriptStackTrace -funcFilePath $errorPath
					Continue pageLoop
				}
			}
			#endregion
		} else {
			Write-Host ("      {0} - Skipping because it is a Modern Page or in the skip list: {1}" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"), $pageName) -ForegroundColor Yellow
		}

		#region cleaning up variables
		Remove-Variable pageName
		Remove-Variable pageCTId
		Remove-Variable pageAppId
		if($null -ne $pageTemp) {Remove-Variable pageTemp}
		#endregion
	}

	# Writing log to SharePoint
	Write-Host "      " -NoNewline
	Save-PnPPageConversionLog #-PageName ("{0}-{1}-{2}" -f $library.srcSiteTitle,$library.srcLibrary,(Get-Date -Format "yyMMddHHmmss"))

	#region cleaning up variables
	Disconnect-PnPOnline
	Remove-Variable destConnection
	Remove-Variable srcPages
	#endregion
} 
#endregion CSV Loop Example

#endregion Main Script

#region Completed Script
# Output Script Info
Complete-CatScript $scriptStart[0]

# Stop logging
Stop-Transcript
#endregion Completed Script
