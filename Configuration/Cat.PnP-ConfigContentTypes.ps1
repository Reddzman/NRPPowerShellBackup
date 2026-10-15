#region Documentation
<#  Cat.PnP-ConfigMulitpleLists.ps1

This script iterates through a CSV file to create lists or libraries and configure content types.

Dependencies
	-	PnP PowerShell https://docs.microsoft.com/en-us/powershell/sharepoint/sharepoint-pnp/sharepoint-pnp-cmdlets?view=sharepoint-ps
	-	General Functions - ..\General\Cat.PnP-GeneralFunctionsFunctions.psm1
	-	Input CSV File - Example: ..\_Input Files\Sample\Configuration\Cat.Tech-Action-Test.csv

Call script with parameters
	.\Cat.PnP-CreateMultipleLists.ps1  -inputFile Cat.PnP-ConfigMultipleLists-Client.xml -appID "57577043-3de0-40c0-9c68-7960e2cebe8b" -CertificatePath "C:\GitHub\NRPDev\AppRegistration\PnP PowerShell Authentication.pfx" -clientTenant "tenantname" 

Table of Contents
	1.	Input & Output Files
	2.	Variables
	3.	Snapin(s), Modules & Environment Connection(s)
	4.	Main Script
	5.	Determine Settings
	6.	Creating List
	7.	Configuring List Settings
	8.	Adding Content Type
#>
#endregion

#region Script Parameters
Param(	[string] $inputFile = "Samples\Configuration\Cat.PnP-ConfigMultipleLists-Test.csv",
		[string] $appID = "57577043-3de0-40c0-9c68-7960e2cebe8b",
		[string] $certPath = "C:\GitHub\NRPDev\AppRegistration\PnP PowerShell Authentication.pfx",
		[string] $clientTenant = "nrpdev2"
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
Start-Transcript -Path ("..\_Log FIles\{0}.PnP-ConfigMultipleLists-{1}.log" -f $client, $scriptStart[1])
#endregion

#region 1.	Input & Output Files
# File Paths
$inPath = ("..\_Input Files\{0}" -f $inputFile)

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

#region 2.	Variables
# User Inputs
Write-Host
Write-Host "User Input Variables" -ForegroundColor Cyan
$certPass = ConvertTo-SecureString -String (Read-Host -Prompt "   Enter certificate password for the app registration") -AsPlainText -Force

# Script Variables
$tenant = ("{0}.onmicrosoft.com" -f $clientTenant)

# Output Parameters & Variables
Write-Host
Write-Host "Script Parameters" -ForegroundColor Cyan
Write-Host ("   Input File: {0}" -f $inputFile)
Write-Host ("   Client Tenant Domain: {0}" -f $clientTenant)
Write-Host ("   App ID: {0}" -f $appID)
Write-Host ("   Certificate Path: {0}" -f $certPath)
Write-Host
Write-Host "User Input" -ForegroundColor Cyan
Write-Host ("   Certificate Password: {0}" -f $certPass)
Write-Host
Write-Host "Script Variables" -ForegroundColor Cyan
Write-Host ("   Microsoft Tenant: {0}" -f $tenant)
#endregion

#region 3.	Snapin(s), Modules & Environment Connection(s)
# Load MS / PnP PowerShell Modules
Get-Module -ListAvailable | Where{$_.Name -like "SharePoint*"} | Remove-Module
Import-Module SharePointPnPPowershellOnline -DisableNameChecking # Add appropriate PnP Module for your use case
#endregion
    
#region 4.	Main Script
Write-Host
Write-Host "Configuring Lists" -ForegroundColor Cyan

# Looping through CSV File to create Document Libraries
foreach ($list in $inFile) {
	# Connecting To SP Online site
	Write-Host
	Write-Host ("   Connecting to SharePoint Online site: {0}" -f $list.siteURL) -ForegroundColor Magenta
	Connect-PnPOnline -ClientId $appID -Url $list.siteURL -Tenant $tenant -CertificatePath $certPath -CertificatePassword $certPass -WarningAction SilentlyContinue

	#region 5.	Determine Settings
	$defaultContentType = switch ($list.defaultContentType) { true {$true} false {$false}	Default {$null} }
	$listHidden = switch ($list.hidden) { true {$true} false {$false}	Default {$null} }
	$enableContentTypes  = switch ($list.enableContentTypes) { true {$true} false {$false}	Default {$null} }
	$onQuicklaunch = switch ($list.onQuicklaunch) { true {$true} false {$false}	Default {$null} }
	$forceCheckout = switch ($list.forceCheckout) { true {$true} false {$false}	Default {$null} }
	$enableAttachments = switch ($list.enableAttachments) { true {$true} false {$false}	Default {$null} }
	$enableFolderCreation = switch ($list.enableFolderCreation) { true {$true} false {$false}	Default {$null} }
	$enableVersions = switch ($list.enableversions) { true {$true} false {$false}	Default {$null} }
	$enableMinorVersions = switch ($list.enableMinorVersions) { true {$true} false {$false}	Default {$null} }
	$enableModeration = switch ($list.enableModeration) { true {$true} false {$false}	Default {$null} }
	#endregion

	#region 6.	Creating List
	# Determine if List Exists
	$tempList = Get-PnPList $list.listURL
	if ($tempList) {
		Write-Host ("      List already exists: {0}" -f $list.listURL)
	}
	else {
		Write-Host ("      Creating list: {0}" -f $list.listURL) -NoNewline
		$tempList = New-PnPList -Title $list.listName `
			-Template $list.listTemplate `
			-Url $list.listURL `
			-OnQuickLaunch:$onQuicklaunch 
		Write-Host " - Complete" -ForegroundColor Green	
	}
	#endregion

	#region 7.	Configuring List Settings
	Write-Host ("         Configuring list settings: {0}" -f $list.listURL) -ForegroundColor Blue -NoNewline
	Set-PnPList -Identity $tempList `
		-EnableContentTypes:$enableContentTypes `
        -Hidden:$listHidden `
        -ForceCheckout:$forceCheckout `
        -ListExperience:$list.listExperience `
        -EnableAttachments:$enableAttachments `
        -EnableFolderCreation:$enableFolderCreation `
        -EnableVersioning:$enableVersions `
        -EnableMinorVersions:$enableMinorVersions `
        -MajorVersions $list.majorVersions `
        -MinorVersions $list.minorVersions `
        -EnableModeration:$enableModeration
	Write-Host " - Complete" -ForegroundColor Green
	#endregion

	#region 8.	Adding Content Type
	Write-Host ("         Adding content type to list: {0}" -f $list.contentType) -ForegroundColor Blue -NoNewline
	$contentType = Get-PnPContentType -Identity $list.contentType
	Add-PnPContentTypeToList -List $tempList `
		-ContentType $contentType `
		-DefaultContentType:$defaultContentType
	Write-Host " - Complete" -ForegroundColor Green
	#endregion

	# Disconnect from SP Online Site
	Write-Host ("      Disconnecting from SharePoint Online site: {0}" -f $list.siteURL)
	Disconnect-PnPOnline
} # foreach ($list in $inFile)
#endregion

#region Completed Script
# Output Script Info
Complete-CatScript $scriptStart[0]

# Stop logging
Stop-Transcript
#endregion