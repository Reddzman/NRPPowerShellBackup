#region Documentation
<#  Cat.SPO-AddSiteTemplate.ps1

This script...

Dependencies
	-	SPO PowerShell Modules
	-	General Functions - ..\General\Cat.GeneralFunctions.psm1
	-	Input JSON File - Example: ..\_Input FIles\Testing\Cat.Tech-Action-Test.csv

Call script with parameters
	.\Cat.SPO-AddSiteTemplate.ps1 -inputFile "Samples\Provisioning\Cat.SPO-ITPMOSCRUMTemplate.json" -clientTenant "nrpdev2" -scriptName "IT SCRUM Template" -designName "IT SCRUM Template" -webTemplateNumber 64

Table of Contents
	1.	Snapin(s), Modules & Environment Connection(s)
	2.	Input & Output Files
	3.	Variables
	4.	Main Script
	5.	Creating/Updating Site Script
	6.	Creating/Updating Site Design
#>
#endregion

#region Script Parameters
Param(	[string] $inputFile = "Samples\Provisioning\Cat.SPO-ITPMOSCRUMTemplate.json",
		[string] $clientTenant = "nrpdev2",
		[string] $scriptName = "IT PMO SCRUM Template",
		[string] $designName = "IT PMO SCRUM Template",
		[int] $webTemplateNumber = 64
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
Start-Transcript -Path ("..\_Log Files\{0}.SPO-AddSiteTemplate-{1}.log" -f $clientTenant, $scriptStart[1])
#endregion

#region 1.	Snapin(s), Modules & Environment Connection(s)
# Load MS / PnP PowerShell Modules
Get-Module -ListAvailable | Where-Object{($_.Name -like "*SharePoint*") -or ($_.Name -like "*PnP*") } | Remove-Module
Import-Module Microsoft.Online.SharePoint.PowerShell -DisableNameChecking -erroraction SilentlyContinue
#endregion

#region 2.	Input & Output Files
# File Paths
$inPath = ("..\_Input Files\{0}" -f $inputFile)

# Import Input File(s)
# Load CSV file
If (Test-Path ($inPath)){
	$inFile = Get-Content $inPath -Raw
}
else {
	Write-Host
	throw ("The input file does not exist: {0}" -f $inPath)	
}
#endregion

#region 3.	Variables
# Script Variables
$spoAdminUrl = ("https://{0}-admin.sharepoint.com" -f $clientTenant)

# Output Parameters & Variables
Write-Host
Write-Host "Script Parameters" -ForegroundColor Cyan
Write-Host ("   Input File: {0}" -f $inputFile)
Write-Host ("   Client Tenant Domain: {0}" -f $clientTenant)
Write-Host ("   Site Script Name: {0}" -f $scriptName)
Write-Host ("   Site Design Name: {0}" -f $designName)
Write-Host
Write-Host "Script Variables" -ForegroundColor Cyan
Write-Host ("   SPO admin URL: {0}" -f $spoAdminUrl)
#endregion
    
#region 4.	Main Script
# Connecting To SP Online
Write-Host
Write-Host "Creating Site Template:" -ForegroundColor Cyan
Write-Host ("   Connecting to SharePoint Online site: {0}" -f $spoAdminUrl)
Connect-SPOService -Url $spoAdminUrl

#region 5.	Creating/Updating Site Script
# Checking for existing Site Script
Write-Host
Write-Host ("   Checking for existing Site Script: {0}" -f $scriptName) -NoNewline
$existingScript = Get-SPOSiteScript | Where-Object {$_.Title -eq $scriptName}
$update = $null
switch ($existingScript.Count) {
	0 {Write-Host " - Script Not Found"}
	1 {
		Write-Host " - Script Found" -ForegroundColor Yellow
		$update = Read-Host -Prompt "      There is a site script with that name, do you want to update it(Y, N)?"
	}
	{$_ -gt 1}{
		Write-Host " - Multiple Scripts Found" -ForegroundColor Red
		$update = "n"
	}
}

# Add site schema as Site Script
if (($existingScript) -and ($update.ToLower() -eq "y")) {
	Write-Host ("   Updating Site Script: {0}" -f $scriptName)
	$tempScript = Set-SPOSiteScript -Identity $existingScript.Id -Title $scriptName -Content $inFile -Version ($existingScript.version +1)
} elseif (-not($existingScript)) {
	Write-Host ("   Adding Site Script: {0}" -f $scriptName)
	$tempScript = Add-SPOSiteScript -Title $scriptName -Content $inFile
}
#endregion #>

#region 6.	Creating/Updating Site Design
# Checking for existing Site Design
Write-Host ("   Checking for existing Site Design: {0}" -f $scriptName) -NoNewline
$existingDesign = Get-SPOSiteDesign | Where-Object {$_.Title -eq $designName}
$update = $null
switch ($existingDesign.Count) {
	0 {Write-Host " - Script Not Found"}
	1 {
		Write-Host " - Script Found" -ForegroundColor Yellow
		$update = Read-Host -Prompt "      There is a site script with that name, do you want to update it(Y, N)?"
	}
	{$_ -gt 1}{
		Write-Host " - Multiple Scripts Found" -ForegroundColor Red
		$update = "n"
	}
}

# Add site schema as Site Script
if ($existingDesign -and ($update.ToLower() -eq "y")) {
	Write-Host ("   Updating Site Design: {0}" -f $designName)
	$tempDesign = Set-SPOSiteDesign -Identity $existingDesign.Id -Title $designName -WebTemplate $webTemplateNumber -SiteScripts $tempScript.Id
} elseif (-not($existingDesign)) {
	Write-Host ("   Adding Site Design: {0}" -f $designName)
	$tempDesign = Add-SPOSiteDesign -Title $designName -WebTemplate $webTemplateNumber -SiteScripts $tempScript.Id
}
#endregion #>

# Disconnect from Services
Write-Host ("   Disconnecting from SharePoint Online: {0}" -f $spoAdminUrl)
Disconnect-SPOService
#endregion

#region Completed Script
# Output Script Info
Complete-CatScript $scriptStart[0]

# Stop logging
Stop-Transcript
#endregion