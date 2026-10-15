<#  Cat.Tech-ScriptName.ps1 Description

This script script iterates through a CSV file and reset permissions inheritance (or deletes unique permission). The
CSV file is the same format as the locations report from the Cat.PnP.BrokenInheritanceBySite.ps1 script. (Reporting\Cat.PnP-BrokenInheritanceBySite.ps1) 

Dependencies
	-	PnP.PowerShell Modules
	-	Input CSV File - Example: ..\_Input Files\Samples\Configuration\Cat.PnP-ResetInheritanceByCSV-Test.csv

Call script with parameters
	cd c:\GitHub\CatPowerShell\Configuration
	.\Cat.PnP-ResetInheritanceByCSV.ps1  -inputFile Samples\Configuration\Cat.PnP-ResetInheritanceByCSV-Test.csv -clientAppID "caf6b11a-cf03-416c-b0f0-67a09a8e30ee" -clientCertPath "C:\GitHub\NRPDev\AppRegistration\PowerShell App Registration.pfx" -clientTenant "nrpdev2"

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
		[string] $clientAppID = "{Client app registration ID from Azure}", # Leave Blank for interactive login
		[string] $clientCertPath = "{Path to client certificate file must be installed on local user}",# Leave Blank for interactive login
		[string] $clientTenant = "{Client tenant SharePoint prefix}" # IE. ClientTenant.sharepoint.com
)
#endregion Script Parameters

#region Start Date / Time
# Remove Existing Catapult Modules to ensure PS session has most up to date functions
Get-Module Cat.* | Remove-Module

# Load General Functions
Import-Module "..\General\Cat.GeneralFunctions.psm1"

# Get Start Date / Time
$scriptStart = Start-CatScript

# Start Transcript
Start-Transcript -Path ("..\_Log Files\{0}.PnP-ResetInheritanceByCSV-{1}.log" -f $clientTenant, $scriptStart[1])
#endregion Start Date / Time

#region 1.	Snapin(s), Modules & Environment Connection(s)
# Load MS / PnP PowerShell Modules
Get-Module -ListAvailable | Where-Object{($_.Name -like "*SharePoint*") -or ($_.Name -like "*PnP*") } | Remove-Module
Import-Module PnP.PowerShell
#endregion Snapin(s), Modules & Environment Connection(s)

#region 2.	Input & Output Files
# File Paths
$inPath = ("..\_Input Files\{0}" -f $inputFile)
$errorPath = ("..\_Log Files\{0}.PnP-ResetInheritanceByCSV-Errors-{1}.csv" -f $clientTenant, $scriptStart[1])

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
If ($clientAppID) {
	Write-Host
	Write-Host "User Input Variables" -ForegroundColor Cyan
	$certPass = ConvertTo-SecureString -String (Read-Host -Prompt "   Enter certificate password for the app registration") -AsPlainText -Force
	$cert = Get-PnPAzureCertificate -Path $clientCertPath -Password $certPass
	Remove-Variable certPass
	Remove-Variable clientCertPath
}

# Script Variables
$tenantDomain = ("{0}.onmicrosoft.com" -f $clientTenant)

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
If ($clientAppID) {
	Write-Host
	Write-Host "User Inputs" -ForegroundColor Cyan
	Write-Host ("   Certificate Thumbprint: {0}" -f "Only output for troubleshooting", $cert.Thumbprint)
}
Write-Host
Write-Host "Script Variables" -ForegroundColor Cyan
Write-Host ("   Microsoft Tenant: {0}" -f $tenantDomain)
#endregion Output Parameters & Variables
#endregion Variables
    
#region 4.	Main Script
#region Iterating through CSV File
Write-Host
Write-Host ("{0} - Iterating through the object in CSV file ({1})" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"), $inFile.Count) -ForegroundColor Cyan
:loopObjects foreach ($object in $inFile){
	#region Outputing object information to screen
	Write-Host
	Write-Host ("   {0} - Removing Custom Permissions" -f (Get-Date -Format "yy-MM-dd HH:mm:ss")) -ForegroundColor Magenta -NoNewline
	switch ($object.objectType) {
		{ ($_ -eq "DocumentLibrary") -or ($_ -eq "GenericList") } {Write-Host (" - {0}" -f $object.listTitle) -ForegroundColor Magenta}
		{ ($_ -eq "File") -or ($_ -eq "Item") -or ($_ -eq "Folder") } {Write-Host (" - {0}" -f $object.itemTitle) -ForegroundColor Magenta}
		Default {
			Write-Host " - Unknown object Type" -ForegroundColor DarkRed
			Write-CatErrorCSV -funcCSVItem $list -funcAction "Unknown object Type" -funcException $_.Exception -funcDetails $_.ErrorDetails -funcStackTrace $_.ScriptStackTrace -funcFilePath $errorPath
			Continue loopObjects
		}
	}
	#endregion Outputing object information to screen
	
	#region Connecting To SP Online
	Write-Host ("      {0} - Connecting to SharePoint Online site" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"), $object.siteTitle)
	try {
		if ($clientAppID) {
			Connect-PnPOnline $object.siteURL -tenant $tenantDomain -ClientId $clientAppID -Thumbprint $cert.Thumbprint -ValidateConnection -ErrorAction Stop # PnP Modules w/ App Reg
		} else {
			Connect-PnPOnline -Url $object.siteURL -Interactive -ValidateConnection -ErrorAction Stop # PnP Modules w/o App Reg
		} 
	}
	catch {
		Write-CatErrorCSV -funcCSVItem $object -funcAction "Connecting to SharePoint" -funcException $_.Exception -funcDetails $_.ErrorDetails -funcStackTrace $_.ScriptStackTrace -funcFilePath $errorPath
		Continue loopObjects
	}
	#endregion Connecting To SP Online

	#region Removing Custom Permission
	Write-Host ("      {0} - Removing Permissions" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"))
	try {
		$list = Get-PnPList $object.listTitle
		switch ($object.objectType) {
			{ ($_ -eq "DocumentLibrary") -or ($_ -eq "GenericList") } {
				$temp = Set-PnPList -Identity $list -ResetRoleInheritance
				Remove-Variable temp

			}
			{ ($_ -eq "File") -or ($_ -eq "Item") -or ($_ -eq "Folder") } {
				$item = Get-PnPListItem -List $list -Id $object.itemId
				Set-PnPListItemPermission -List $list -Identity $item -InheritPermissions
				Remove-Variable item
			}
		}
		Remove-Variable list
	}
	catch {
		Write-CatErrorCSV -funcCSVItem $object-funcAction "Removing Permissions" -funcException $_.Exception -funcDetails $_.ErrorDetails -funcStackTrace $_.ScriptStackTrace -funcFilePath $errorPath
		Continue loopObjects
	}
	#endregion Sub Task / Error Handling Example

	#region Disconnecting from SP Online
	Write-Host ("      {0} - Disconnecting from SharePoint Online" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"))
	Disconnect-PnPOnline
	#endregion Disconnecting from SP Online
} 
#endregion Iterating through CSV File

#endregion Main Script

#region Completed Script
# Output Script Info
Complete-CatScript $scriptStart[0]

# Stop logging
Stop-Transcript
#endregion Completed Script