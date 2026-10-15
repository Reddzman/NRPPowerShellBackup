<#  Cat.PnP-ODBSiteDeletion.ps1

**** DO NOT USE ON ACTIVE ODB ACCOUNTS ****

This script this script iterates through a csv and deletes the OneDrive site collection for users.  DO NOT USE ON ACTIVE ODB ACCOUNTS
It can be used for migration testing to delete migrated content that doesn't need to be retained or test the provisioning process.

Dependencies
	-	PnP.PowerShell Modules
	-	Input CSV File - Example: ..\_Input FIles\Cat.PnP-ODBSiteDeletion-Test.csv

Call script with parameters
	.\Cat.PnP-ODBSiteDeletion.ps1 -inputFile "Samples\Configuration\Cat.PnP-ODBSiteDeletion-Test.csv" -clientTenant "nrpdev2" -appID "57577043-3de0-40c0-9c68-7960e2cebe8b" -certPath "C:\GitHub\NRPDev\AppRegistration\PnP PowerShell Authentication.pfx"

Table of Contents
	1.	Input & Output Files
	2.	Variables
	3.	Snapin(s), Modules & Environment Connection(s)
	4.	Main Script
#>

# *******************************************
# **** DO NOT USE ON ACTIVE ODB ACCOUNTS ****
# *******************************************

#region Script Parameters
Param(
    [String] $inputFile = "Cat.SG-OD4BList-TEST.csv",
    [String] $clientTenant = "{ClientTenantPrefix}", # {ClientTenantPrefix}.sharepoint.com; No {}
    [String] $appID = "{AzureAppRegistrationID}", # No {}
    [String] $certPath = "C:\Temp\appReg\ClientAppRegisrationCertificate.pfx"
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
Start-Transcript -Path ("..\_Log FIles\{0}.PnP-ODBSiteDeletion-{1}.log" -f $clientTenant, $scriptStart[1])
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
# User Inputs
Write-Host
Write-Host "Enter Connection Information" -ForegroundColor Cyan
$certPass = ConvertTo-SecureString -String (Read-Host -Prompt "   Enter certificate password for the App Registration") -AsPlainText -Force 

# Script Variables
$spoAdminUrl = ("https://{0}-admin.sharepoint.com/" -f $clientTenant)
$tenant= ("{0}.onmicrosoft.com" -f $clientTenant)

# Output Parameters & Variables
Write-Host
Write-Host "Script Parameters" -ForegroundColor Cyan
Write-Host ("   Input File: {0}" -f $inputFile)
Write-Host ("   Client Tenant: {0}" -f $clientTenant)
Write-Host ("   App ID: {0}" -f $appID)
Write-Host ("   Certificate Path: {0}" -f $certPath)
Write-Host
Write-Host "Script Variables" -ForegroundColor Cyan
Write-Host ("   SharePoint Admin URL: {0}" -f $spoAdminUrl)
Write-Host ("   Tenant Domain: {0}" -f $tenant)
#endregion

#region 3.	Snapin(s), Modules & Environment Connection(s)
# Load MS / PnP PowerShell Modules
Import-Module PnP.PowerShell
#endregion
    
#region 4.	Main Script
# Get confirmation to remove permissions
Write-Host
Write-Host "*** Deleting existing OneDrive site Collections will caust the PERMINANT loss of that user's OneDrive Data ***" -ForegroundColor Yellow
$userProceed = Read-Host "   Are you sure you want to proceed with deleting existing OneDrive site collections (Yes or No)" 

# Exit if user doesn't want to continue
if ($userProceed.ToLower() -eq "yes") {
}

switch ($userProceed) {
	yes {
		# Connecting To SP Online
		Write-Host
		Write-Host "Connecting to Services:" -ForegroundColor Cyan
		Write-Host ("   Connecting to SharePoint Online: {0}" -f $spoAdminURL)
		$cert = Get-PnPAzureCertificate -CertificatePath $certPath -CertificatePassword $certPass
		$pnpConnection = Connect-PnPOnline $spoAdminURL -tenant $tenant -ClientId $appID -Thumbprint $cert.Thumbprint

		# Screen Output
		Write-Host
		Write-Host ("Deleting OneDrive Site Collections: {0}" -f $inputFile) -ForegroundColor Cyan

		# Looping through CSV
		ForEach ($odbSite in $inFile){
			Write-Host
			Write-Host ("   Removing OneDrive: {0}" -f $odbSite.upn) -ForegroundColor Magenta

			# Get Users OneDrive Site URL
			Write-Host "      Retrieving OneDrive Url"
			$odbURL = Get-PnPUserProfileProperty -Account $odbSite.upn | Select-Object PersonalUrl
			$odbURL = $odbURL.PersonalUrl.TrimEnd("/")

			# Deleting OneDrive Site Collection
			if ($odbURL) {
				Write-Host "      Deleting OneDrive Site Collection"
				Remove-PnPTenantSite -Url $odbURL -Connection $pnpConnection -Force        
			}
			else {
				Write-Host ("      OneDrive Site not found for user: {0}" -f $odbSite.upn) -ForegroundColor Yellow
			}
		}

		# Script Complete Output
		Write-Host "Successfully deleted OneDrive Site Collections" -ForegroundColor Green

		# Disconnect from Services
		Write-Host
		Write-Host "Disconnecting from Services:" -ForegroundColor Cyan
		Write-Host ("   Disconnecting from SharePoint Online: {0}" -f $spoAdminUrl)
		Disconnect-PnPOnline -Connection $pnpConnection
	}
	no {
		Write-Host
		Write-Host "Ending script to avoid deleting data" -ForegroundColor Yellow
	}
	Default {
		Write-Host
		Write-Host "Ending script because we didn't understand your reponse" -ForegroundColor Yellow
	}
}
#endregion

#region Completed Script
# Output Script Info
Complete-CatScript $scriptStart[0]

# Stop logging
Stop-Transcript
#endregion