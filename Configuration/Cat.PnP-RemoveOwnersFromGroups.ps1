#region Documentation
<#  Cat.PnP-RemoveOwnersFromGroups.ps1

This script was created in relation to the Cat.PnP-AddUsersToGroups.ps1 script.  The client provided a file with incorrect owners and therefore
we needed to demote them from owners the groups.  In our case we then re-ran the Add Users script to give them member permissions.

Dependencies
	-	PnP PowerShell https://pnp.github.io/powershell/
	-	General Functions - ..\General\Cat.GeneralFunctions.psm1
	-	Input CSV File in ..\_Input Files\ - Example: Samples\Configuration\Cat.PnP-RemoveOwnersFromGroups-Client.csv

Call script with parameters
	cd C:\GitHub\CatPowerShell\Configuration\
	.\Cat.PnP-RemoveOwnersFromGroups.ps1 -inputFile "Samples\Configuration\Cat.PnP-RemoveOwnersFromGroups-Client.csv" -appID "caf6b11a-cf03-416c-b0f0-67a09a8e30ee" -certPath "C:\GitHub\NRPDev\AppRegistration\PowerShell App Registration.pfx" -clientTenant "nrpdev2"

Table of Contents
	1.	Snapin(s), Modules & Environment Connection(s)
	2.	Input & Output Files
	3.	Variables
	4.	Main Script
#>
#endregion

#region Script Parameters - Values are samples, remove the brackets for the actual values
Param(	[string] $inputFile = "Cat.PnP-ConfigMultipleLists-Client.csv",
		[string] $appID = "[ClientAppID]",
		[string] $certPath = "C:\GitHub\client\AppRegistration\ClientCertificate.pfx",
		[string] $clientTenant = "[ClientTenatPrefix]"
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
Start-Transcript -Path ("..\_Log Files\{0}.PnP-RemoveOwnersFromGroups-{1}.log" -f $clientTenant, $scriptStart[1])
#endregion

#region 1.	Snapin(s), Modules & Environment Connection(s)
# Load MS / PnP PowerShell Modules
Get-Module -ListAvailable | Where-Object{($_.Name -like "*SharePoint*") -or ($_.Name -like "*PnP*") } | Remove-Module
Import-Module PnP.PowerShell

# Load Custom Modules
Import-Module "..\Reporting\Cat.ReportingFunctions.psm1"
#endregion

#region 2.	Input & Output Files
# File Paths
$inPath = ("..\_Input Files\{0}" -f $inputFile)
$errorPath = ("..\_Output Files\{0}.PnP-RemoveOwnersFromGroups-Errors-{1}.csv" -f $client, $scriptStart[1])

# Import Input File(s)
# Load CSV file
If (Test-Path ($inPath)){
	$inFile = Import-Csv $inPath -Encoding Unicode
}
else {
	Write-Host
	Stop-Transcript
	throw ("The input file does not exist: {0}" -f $inPath)	
}
#endregion

#region 3.	Variables
# User Inputs
Write-Host
Write-Host "User Input Variables" -ForegroundColor Cyan
$certPass = ConvertTo-SecureString -String (Read-Host -Prompt "   Enter certificate password for the app registration") -AsPlainText -Force
$cert = Get-PnPAzureCertificate -Path $certPath -Password $certPass

# Script Variables
$tenant = ("{0}.onmicrosoft.com" -f $clientTenant)

# Output Parameters & Variables
Write-Host
Write-Host "Script Parameters" -ForegroundColor Cyan
Write-Host ("   Input File: {0}" -f $inputFile)
Write-Host ("   App ID: {0}" -f $appID)
Write-Host ("   Certificate Path: {0}" -f $certPath)
Write-Host ("   Tenant Domain: {0}" -f $clientTenant)
Write-Host
Write-Host "User Inputs" -ForegroundColor Cyan
Write-Host ("   Certificate Password: {0}" -f $certPass)
Write-Host
Write-Host "Script Variables" -ForegroundColor Cyan
Write-Host ("   Microsoft Tenant: {0}" -f $tenant)
#endregion
    
#region 4.	Main Script
# Connecting To SP Online
Write-Host
Write-Host ("{0} - Removing Owner Permisssions: {1}" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"), $inputFile) -ForegroundColor Cyan

# Looping CSV Records
:acctLoop ForEach ($acct in $inFile){
	try {
		# Connecting to the SharePoint site
		Write-Host
		Write-Host ("   {0} - Removing Account: {1}, from Group: {2}" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"),$acct.loginName, $acct.groupName) -ForegroundColor Magenta
		Write-Host "      Connecting to SharePoint site"
		$spSiteConnection = Connect-PnPOnline $acct.siteUrl -tenant $tenant -ClientId $appID -Thumbprint $cert.Thumbprint -ReturnConnection -ErrorAction Stop
			
		# Removing owner from M365 group
		Write-Host "      Removing Owner Permission"
		Remove-PnPMicrosoft365GroupOwner -Identity $acct.groupName -Users $acct.loginName

		# Disconning from Site & Clearing Variables
		Disconnect-PnPOnline -Connection $spSiteConnection
		Remove-Variable spSiteConnection
	}
	catch {Write-CatCSVErrorInfo -funcError $_ -funcObj $acct -funcOutputPath $errorPath}
}
#endregion

#region Completed Script
# Output Script Info
Complete-CatScript $scriptStart[0]

# Stop logging
Stop-Transcript
#endregion