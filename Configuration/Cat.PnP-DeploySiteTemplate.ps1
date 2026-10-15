#region Documentation
<#  Cat.Tech-ScriptName.ps1

This script...

Dependencies
	-	PnP PowerShell https://pnp.github.io/powershell/
	-	Microsoft Teams Module
	-	Metalogix PowerShell Snapins
	-	General Functions - ..\General\Cat.GeneralFunctions.psm1
	-	SQL Server Command Utility - https://www.microsoft.com/en-us/download/details.aspx?id=53591
	-	Input CSV File - Example: ..\_Input FIles\Testing\Cat.Tech-Action-Test.csv
	-	Input XML File - Example: ..\_Input Files\Testing\Cat.Tech-Action-Test.xml

Call script with parameters
	.\Cat.Tech-ScriptName.ps1  -inputFile Clients\Cat.MT-DeployCrisisTeam-Client.xml -client "ClientNameNoSpacesOrAbriviation"

Table of Contents
	1.	Snapin(s), Modules & Environment Connection(s)
	2.	Input & Output Files
	3.	Variables
	4.	Main Script
	5.	...
#>
#endregion

#region Script Parameters
Param(	[string] $inputFile = "Samples\Configuration\Cat.PnP-ConfigMultipleLists-Test.csv",
		[string] $clientAppID = "57577043-3de0-40c0-9c68-7960e2cebe8b",
		[string] $clientCertPath = "C:\GitHub\NRPDev\AppRegistration\PowerShell App Registration.pfx",
		[string] $clientTenant = "nrpdev2",
		[int] $maxItems = 5000,
		[boolean] $setting = $true 
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
Start-Transcript -Path ("..\_Log Files\{0}.Tech-ScriptName-{1}.log" -f $clientTenant, $scriptStart[1])
#endregion

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
#endregion

#region 2.	Input & Output Files
# File Paths
$inPath = ("..\_Input Files\Clients\{0}" -f $inputFile)
$outPath = ("..\_Output Files\Clients\{0}{1}.csv" -f $inputFile.Replace(".csv","-"), $scriptStart[1])

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
$inFile = Get-Content $inPath -Raw
#endregion

#region 3.	Variables
# User Inputs
Write-Host
Write-Host "User Input Variables" -ForegroundColor Cyan
$certPass = ConvertTo-SecureString -String (Read-Host -Prompt "   Enter certificate password for the app registration") -AsPlainText -Force

# Script Variables
$tenantDomain = ("{0}.onmicrosoft.com" -f $clientTenant)
$spoAdminUrl = ("https://{0}.sharepoint.com" -f $clientTenant)
$itemCount = 5000

# Output Parameters & Variables
Write-Host
Write-Host "Script Parameters" -ForegroundColor Cyan
Write-Host ("   Input File: {0}" -f $inputFile)
Write-Host ("   App ID: {0}" -f $clientAppID)
Write-Host ("   Certificate Path: {0}" -f $clientCertPath)
Write-Host ("   Client Tenant Domain: {0}" -f $clientTenant)
Write-Host
Write-Host "User Inputs" -ForegroundColor Cyan
Write-Host ("   Certificate Password: {0}" -f $certPass)
Write-Host
Write-Host "Script Variables" -ForegroundColor Cyan
Write-Host ("   Microsoft Tenant: {0}" -f $tenantDomain)
Write-Host ("   SPO admin URL: {0}" -f $spoAdminUrl)
Write-Host ("   itemCount: {0}" -f $itemCount)
Write-Host
Write-Host "Array of Info" -ForegroundColor Cyan
Write-Host $arrayName
#endregion
    
#region 4.	Main Script
# Connecting To SP Online
Write-Host
Write-Host "Connecting to Services:" -ForegroundColor Cyan
Write-Host ("   Connecting to SharePoint Online site: {0}" -f $siteURL) -ForegroundColor Magenta
Connect-PnPOnline -ClientId $clientAppID -Url $siteURL -Tenant $tenantDomain -CertificatePath $clientCertPath -CertificatePassword $certPass -WarningAction SilentlyContinue

Add-PnPSiteScript -Title "IT PMO SCRUM Configration" -Description "This script contains the site configuration including lists and libraries for the IT PMO Scrum Template" -Content $scriptJSON
Add-PnPSiteDesign -Title "IT PMO SCRUM Template" -Description "This template is for use by the IT PMO group for use on SCRUM projects" -SiteScriptIds "f50449d1-d434-4fbd-88e7-04f16adaab5c"
$scriptJSON = Get-Content -Path "..\_Input Files\OnSemi.Template-ITPMOScrum.json" -Raw

#endregion

#region Completed Script
# Disconnect from Services
Write-Host
Write-Host "Disconnecting from Services:" -ForegroundColor Cyan
Write-Host ("   Disconnecting from SharePoint Online: {0}" -f $spAdminURL)
Disconnect-PnPService

# Output Script Info
Complete-CatScript $scriptStart[0]

# Stop logging
Stop-Transcript
#endregion