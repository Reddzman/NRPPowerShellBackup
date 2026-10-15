#region Documentation
<#  Q.PnP-AddSiteCollectionAdmins.ps1

This script iterates a CSV file to configure site collection admins for each site. 

Dependencies
	-	PnP PowerShell https://pnp.github.io/powershell/
	-	General Functions - ..\General\Cat.GeneralFunctions.psm1
	-	Input CSV File in ..\_Input Files\ - Example: Samples\Configuration\Q.PnP-AddSiteCollectionAdmins-Client.csv

Call script with parameters
	cd C:\Quisitive\PowerShell\Configuration
	.\Q.PnP-AddSiteCollectionAdmins -inputFile "AdminsToAdd.csv" -appID "6bd4f453-6253-4f7a-be0c-717ecd464321" -certPath "C:\Quisitive\PowerShell\App Reg\SharePointMigration2024.pfx" -clientTenant "sanfordhealth"
	
Table of Contents
	1.	Snapin(s), Modules & Environment Connection(s)
	2.	Input & Output Files
	3.	Variables
	4.	Main Script
	5.	Adding Site Collection Admins
#>
#endregion

#region Script Parameters - Values are samples, remove the brackets for the actual values
Param(	[string] $inputFile = "Q.PnP-AddSiteCollectionAdmins-Client.csv",
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
Start-Transcript -Path ("..\_Log Files\{0}.PnP-AddSiteCollectionAdmins-{1}.log" -f $clientTenant, $scriptStart[1])

# Disabling PNP Upgrade checks
$env:PNPPOWERSHELL_UPDATECHECK="false" # Disable Update Check
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

# Script Variables
$tenant = ("{0}.onmicrosoft.com" -f $clientTenant)
$tenantRootSite = ("https://{0}.sharepoint.com" -f $clientTenant)
$m365GroupClaim = "c:0o.c|federateddirectoryclaimprovider|"

# Output Parameters & Variables
Write-Host
Write-Host "Script Parameters" -ForegroundColor Cyan
Write-Host ("   Input File: {0}" -f $inputFile)
Write-Host ("   App ID: {0}" -f $appID)
Write-Host ("   Certificate Path: {0}" -f $certPath)
Write-Host
Write-Host "User Inputs" -ForegroundColor Cyan
Write-Host ("   Certificate Password: {0}" -f $certPass)
Write-Host
Write-Host "Script Variables" -ForegroundColor Cyan
Write-Host ("   Microsoft Tenant: {0}" -f $tenant)
Write-Host ("   Site URL: {0}" -f $tenantRootSite)
#endregion
    
#region 4.	Main Script
# Connecting To SP Online
Write-Host
Write-Host ("{0} - Updating Site Collection Admin Permisssions: {1}" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"), $inputFile) -ForegroundColor Magenta

# Looping CSV Records
:acctLoop ForEach ($acct in $inFile){
	# Connecting to the SharePoint site
	Write-Host
	Write-Host ("   {0} - Adding Account, {1}, to Site Collection Admin Group on {2}" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"),$acct.loginName, $acct.siteUrl.Replace($tenantRootSite,""))
	Write-Host "      Connecting to SharePoint site" -ForegroundColor DarkGray
	$spSiteConnection = Connect-PnPOnline -ClientId $appID -Url $acct.siteUrl -Tenant $tenant -CertificatePath $certPath -CertificatePassword $certPass -WarningAction SilentlyContinue -ReturnConnection
		
    try {
        Add-PnPSiteCollectionAdmin -Owners $acct.loginName -Connection $spSiteConnection
    }
    catch {
        Write-Host("Unable to add {0}. Verify that this is a valid AD account" -f $acct.loginName) -ForegroundColor DarkRed
    }

	# Disconning from Site & Clearing Variables
	Disconnect-PnPOnline -Connection $spSiteConnection
	Remove-Variable spSiteConnection
}
#endregion

#region Completed Script
# Output Script Info
Complete-CatScript $scriptStart[0]

# Stop logging
Stop-Transcript
#endregion