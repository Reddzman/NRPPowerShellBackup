<#  Cat.PnP-ConfigVersioning.ps1

This script iterates through a CSV file to configure the versioning settings. Initial design is to leverage ShareGate list migration CSV files. This script only support app registration authentication

Dependencies
	-	PnP.PowerShell
	-	General Functions - ..\General\Cat.PnP-GeneralFunctionsFunctions.psm1
	-	Input CSV File

CSV Notes
	-	Force Check out is not supported on Lists, only libraries. Set all lists to false in CSV
	-	Minor Versions is not supported on Lists, only libraries. Set all lists to false in CSV

Call script with parameters
	cd C:\GitHub\CatPowerShell\Configuration\
	.\Cat.PnP-ConfigVersioning.ps1  -inputFile "Samples\Configuration\Cat.PnP-ConfigVersioning-Client.csv" -clientAppID "cat4f453-6253-4f7a-be0c-717ecd464321" -clientCertPath "C:\Quisitive\PowerShell\App Reg\ClientSharePointMigrationCert.pfx" -clientTenant "contoso"

Table of Contents
	1.	Snapin(s), Modules & Environment Connection(s)
	2.	Input & Output Files
	3.	Variables
	4.	Main Script
	5.	Connecting To SP Online
	6.	Determine Settings
	7.	Configuring List Settings
#>

#region Script Parameters
Param(	[string] $inputFile = "{Teams input file path relative to _Input Files}", # IE. "sample\filename.csv"
		[string] $clientAppID = "{Client app registration ID from Azure}", # Only Supports App Reg
		[string] $clientCertPath = "{Path to client certificate file must be installed on local user}",# Only Supports App Reg
		[string] $clientTenant = "{Client tenant SharePoint prefix}" # IE. ClientTenant.sharepoint.com
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
Start-Transcript -Path ("..\_Log FIles\{0}.PnP-ConfigVersioning-{1}.log" -f $clientTenant, $scriptStart[1])
#endregion

#region 1.	Snapin(s), Modules & Environment Connection(s)
# Load MS / PnP PowerShell Modules
Get-Module -ListAvailable | Where-Object{($_.Name -like "*SharePoint*") -or ($_.Name -like "*PnP*") } | Remove-Module
Import-Module PnP.PowerShell
#endregion

#region 2.	Input & Output Files
# File Paths
$inPath = ("..\_Input Files\{0}" -f $inputFile)
$errorPath = ("..\_Output Files\{0}.PnP-ConfigVersioningErrors-{1}.csv" -f $clientTenant, $scriptStart[1])

# Import Input File(s)
# Load CSV file
If (Test-Path ($inPath)){
	$inFile = Import-Csv $inPath -Encoding Unicode
}
else {
	stop-transcript
	throw ("The input file does not exist: {0}" -f $inPath)	
}
#endregion

#region 3.	Variables
# User Inputs
Write-Host
Write-Host "User Input Variables" -ForegroundColor Cyan
$certPass = ConvertTo-SecureString -String (Read-Host -Prompt "   Enter certificate password for the app registration") -AsPlainText -Force
$cert = Get-PnPAzureCertificate -Path $clientCertPath -Password $certPass
Remove-Variable certPass
Remove-Variable clientCertPath

# Script Variables
$tenantDomain = ("{0}.onmicrosoft.com" -f $clientTenant)
$optionsDraftVersionvisibility = @("Approver", "Author", "Reader")

# Output Parameters & Variables
Write-Host
Write-Host "Script Parameters" -ForegroundColor Cyan
Write-Host ("   Input File: {0}" -f $inputFile)
Write-Host ("   App ID: {0}" -f $clientAppID)
Write-Host ("   Client Tenant Prefix: {0}" -f $clientTenant)
Write-Host
Write-Host "User Inputs" -ForegroundColor Cyan
Write-Host ("   Certificate Thumbprint: {0}" -f $cert.Thumbprint)
Write-Host
Write-Host "Script Variables" -ForegroundColor Cyan
Write-Host ("   Microsoft Tenant: {0}" -f $tenantDomain)
Write-Host 
Write-Host "Draft Version Visibility Options" -ForegroundColor Cyan
$optionsDraftVersionvisibility
#endregion
    
#region 4.	Main Script
Write-Host
Write-Host "Configuring Lists" -ForegroundColor Cyan

# Looping through CSV File to create Document Libraries
:listLoop foreach ($list in $inFile) {
	#region 5.	Connecting To SP Online
	Write-Host
	Write-Host ("   {0} - Configuring {1} on {2}" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"), $list.listTitle, $list.siteTitle) -ForegroundColor Magenta
	Write-Host "      Connecting to Site:"
	Write-Host ("	  {0}" -f $list.siteURL) -ForegroundColor DarkGray
	try {
		Connect-PnPOnline $list.siteURL -tenant $tenantDomain -ClientId $clientAppID -Thumbprint $cert.Thumbprint -ErrorAction Stop
		$temp = Get-PnPSite -ErrorAction Stop
	}
	catch {
		Write-CatErrorCSV -funcCSVItem $list -funcAction "Connecting to site" -funcException $_.Exception -funcDetails $_.ErrorDetails -funcStackTrace $_.ScriptStackTrace -funcFilePath $errorPath
		Continue listLoop	
	}
	Remove-Variable temp
	#endregion

	#region 6.	Determine Settings
	Write-Host "      Settings to be Configured:"
	$settingEnableModeration = switch ($list.enableModeration.toLower()) { true {$true} false {$false} Default {$null} }
	Write-Host ("         Enable Moderation (Content Approval): {0}" -f $settingEnableModeration) -ForegroundColor DarkGray
	$settingEnableVersioning = switch ($list.enableVersioning.toLower()) { true {$true} false {$false} Default {$null} }
	Write-Host ("         Versioning Enabled: {0}" -f $settingEnableVersioning) -ForegroundColor DarkGray
	$settingEnableMinorVersions = switch ($list.enableMinorVersions.toLower()) { true {$true} false {$false} Default {$null} }
	Write-Host ("         Minor Versions Enabled: {0}" -f $settingEnableMinorVersions) -ForegroundColor DarkGray
	$settingForceCheckOut = switch ($list.forceCheckOut.toLower()) { true {$true} false {$false} Default {$null} }
	Write-Host ("         Force Check Out: {0}" -f $settingForceCheckOut) -ForegroundColor DarkGray
	Write-Host ("         Major Versions: {0}" -f $list.majorVersions) -ForegroundColor DarkGray
	Write-Host ("         Minor Versions: {0}" -f $list.minorVersions) -ForegroundColor DarkGray
	Write-Host ("         Draft Version Visibility (Draft Security): {0}" -f $list.draftVersionVisibility) -ForegroundColor DarkGray
	#endregion

	#region 7.	Configuring List Settings
	Write-Host "      Configuring list settings" -NoNewline
	try {
		if ($list.draftVersionVisibility -in $optionsDraftVersionvisibility) {
			$temp = Set-PnPList -Identity $list.listTitle `
				-EnableModeration:$settingEnableModeration `
				-EnableVersioning:$settingEnableVersioning `
				-EnableMinorVersions:$settingEnableMinorVersions `
				-ForceCheckout:$settingForceCheckOut `
				-MajorVersions $list.majorVersions `
				-MinorVersions $list.minorVersions `
				-DraftVersionVisibility $list.draftVersionVisibility `
				-ErrorAction Stop -WarningAction Stop
		} else {
			$temp = Set-PnPList -Identity $list.listTitle `
				-EnableModeration:$settingEnableModeration `
				-EnableVersioning:$settingEnableVersioning `
				-EnableMinorVersions:$settingEnableMinorVersions `
				-ForceCheckout:$settingForceCheckOut `
				-MajorVersions $list.majorVersions `
				-MinorVersions $list.minorVersions `
				-ErrorAction Stop -WarningAction Stop
		}
		Remove-Variable temp
		Write-Host " - Successfully Configured" -ForegroundColor Green
	}
	catch {
			Write-Host " - Error during configuration" -ForegroundColor Red
			Write-CatErrorCSV -funcCSVItem $list -funcAction "Configuring Versioning Settings" -funcException $_.Exception -funcDetails $_.ErrorDetails -funcStackTrace $_.ScriptStackTrace -funcFilePath $errorPath
	}
	#endregion

	# Disconnect from SP Online Site
	Write-Host ("      Disconnecting from SharePoint Online site: {0}" -f $list.siteTitle)
	Disconnect-PnPOnline
} # foreach ($list in $inFile)
#endregion

#region Completed Script
# Output Script Info
Complete-CatScript $scriptStart[0]

# Stop logging
Stop-Transcript
#endregion