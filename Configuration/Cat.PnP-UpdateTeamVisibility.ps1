<#  Cat.PnP-UpdateTeamVisibility.ps1

This script iterates through a CSV file of Teams and updates the visibility settings in Outlook and the GAL.

Dependencies
	-	PnP.PowerShell Modules
	-	General Functions - ..\General\Cat.GeneralFunctions.psm1
	-	Input CSV File - Example: ..\_Input FIles\Testing\Cat.Tech-Action-Test.csv

Get all teams from a tenant
	-	Connect to Tenant via PnP with desired method
	-	$teams = Get-PnPTeamsTeam
	-	$teams | Export-Csv -Path "C:\Users\npatterson\Downloads\TeamsList.csv"
	-	Update CSV to UTF-8 csv format

Call script with parameters
	cd C:\GitHub\CatPowerShell\Configuration\
	.\Cat.PnP-UpdateTeamVisibility.ps1 -inputFile "Samples\Configuration\Cat.PnP-UpdateTeamsVisibility-Test.csv" -clientAppID "caf6b11a-cf03-416c-b0f0-67a09a8e30ee" -clientCertPath "C:\GitHub\NRPDev\AppRegistration\PowerShell App Registration.pfx" -clientTenant "nrpdev2" -hideFromAddressLists $false -HideFromOutlookClients $false

Table of Contents
	1.	Snapin(s), Modules & Environment Connection(s)
	2.	Input & Output Files
	3.	Variables
	4.	Main Script
	5.	Configuring Group Visiblity
#>

#region Script Parameters
Param(	[string] $inputFile = "{Teams input file path relative to _Input Files}", # IE. "sample\filename.csv"
		[string] $clientAppID = "{Client app registration ID from Azure}", # Leave Blank for interactive login
		[string] $clientCertPath = "{Path to client certificate file must be installed on local user}",# Leave Blank for interactive login
		[string] $clientTenant = "{Client tenant SharePoint prefix}", # IE. ClientTenant.sharepoint.com
		[boolean] $hideFromAddressLists = $false, # $true/$false
		[boolean] $HideFromOutlookClients = $false # $true/$false
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
Start-Transcript -Path ("..\_Log Files\{0}.PnP-UpdateTeamVisibility-{1}.log" -f $clientTenant, $scriptStart[1])
#endregion

#region 1.	Snapin(s), Modules & Environment Connection(s)
# Load MS / PnP PowerShell Modules
Get-Module -ListAvailable | Where-Object{($_.Name -like "*SharePoint*") -or ($_.Name -like "*PnP*") } | Remove-Module
Import-Module PnP.PowerShell
#endregion

#region 2.	Input & Output Files
# File Paths
$inPath = ("..\_Input Files\{0}" -f $inputFile)
$errorPath = ("..\_Output Files\{0}.PnP-UpdateTeamVisibility-Errors-{1}.csv" -f $clientTenant, $scriptStart[1])

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
}

# Script Variables
$tenantDomain = ("{0}.onmicrosoft.com" -f $clientTenant)
$spoAdminUrl = ("https://{0}-admin.sharepoint.com" -f $clientTenant)

# Output Parameters & Variables
Write-Host
Write-Host "Script Parameters" -ForegroundColor Cyan
Write-Host ("   Input File: {0}" -f $inputFile)
if ($clientAppID) {
	Write-Host ("   App ID: {0}" -f $clientAppID)
	Write-Host ("   Certificate Path: {0}" -f $clientCertPath)
}
Write-Host ("   Client Tenant Domain: {0}" -f $clientTenant)
Write-Host ("   Hide from Address Lists: {0}" -f $hideFromAddressLists)
Write-Host ("   Hide from Outlook Clients: {0}" -f $HideFromOutlookClients)
if ($clientAppID) {
	Write-Host
	Write-Host "User Inputs" -ForegroundColor Cyan
	Write-Host ("   Certificate Password: {0}" -f $certPass)
}
Write-Host
Write-Host "Script Variables" -ForegroundColor Cyan
Write-Host ("   Microsoft Tenant: {0}" -f $tenantDomain)
Write-Host ("   SPO admin URL: {0}" -f $spoAdminUrl)
#endregion
    
#region 4.	Main Script
#region Connecting To SP Online
Write-Host
Write-Host "Connecting to Services" -ForegroundColor Cyan
Write-Host ("   Connecting to SharePoint Online Administraton: {0}" -f $spoAdminURL)
try {
	if ($clientAppID) {
		Connect-PnPOnline $spoAdminURL -tenant $tenantDomain -ClientId $clientAppID -Thumbprint $cert.Thumbprint -WarningAction SilentlyContinue # PnP Modules w/ App Reg
	} else {
		Connect-PnPOnline -Url $spoAdminURL -Interactive -WarningAction SilentlyContinue # PnP Modules w/ App Reg
	} 
}
catch {
	Write-Host
	throw ("Unable to connect: {0}" -f $_.Exception)	
}
#endregion

#region 5.	Configuring Group Visiblity
Write-Host
Write-Host "Configuring Team/Group visiablity" -ForegroundColor Magenta
:teamLoop foreach ($team in $inFile){
	Write-Host
	Write-Host ("   {0} - Configuring {1}" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"), $team.DisplayName)

	# Retrieving Group Object
	Write-Host "      Retrieving M365 Group" -ForegroundColor DarkGray
	try {
		$group = Get-PnPMicrosoft365Group -Identity $team.GroupId -ErrorAction Stop
	}
	catch {
		Write-CatErrorCSV -funcCSVItem $team -funcAction "Retrieving M365 Group" -funcException $_.Exception -funcDetails $_.ErrorDetails -funcStackTrace $_.ScriptStackTrace -funcFilePath $errorPath
		Continue teamLoop
	}

	# Retrieving Group Object
	If ($group) {
		Write-Host "      Configuring M365 Group" -ForegroundColor DarkGray
		try {
			Set-PnPMicrosoft365Group -Identity $group -HideFromAddressLists:$hideFromAddressLists -HideFromOutlookClients:$HideFromOutlookClients -ErrorAction Stop
		}
		catch {
			Write-CatErrorCSV -funcCSVItem $team -funcAction "Configuring M365 Group" -funcException $_.Exception -funcDetails $_.ErrorDetails -funcStackTrace $_.ScriptStackTrace -funcFilePath $errorPath
			Continue teamLoop
		}
		Clear-Variable group
	} else {
		Write-CatErrorCSV -funcCSVItem $team -funcAction "No Group Found" -funcFilePath $errorPath
	}
} 
#endregion

Write-Host
Write-Host "Group Configuration Complete" -ForegroundColor Green
#endregion

#region Completed Script
# Disconnect from Services
Write-Host
Write-Host "Disconnecting from Services:" -ForegroundColor Cyan
Write-Host ("   Disconnecting from SharePoint Online Administration: {0}" -f $spoAdminUrl)
Disconnect-PnPOnline

# Output Script Info
Complete-CatScript $scriptStart[0]

# Stop logging
Stop-Transcript
#endregion