<#  Cat.PnP-DeleteHiddenColumn.ps1

This script deletes instal Microsoft Columns. I recommend using a Microsoft ticket to confirm it is no longer in use as 
you do not want to break functionality. CSOM code was provided directly from Microsoft Tier 4 support.

*** May throw the following error in the script but use the UI to confirm if the columns removal was successful. ***
		MethodInvocationException: C:\GitHub\CatPowerShell\Configuration\Cat.PnP-DeleteHiddenColumn.ps1:124
		Line |
		124 |  $context.ExecuteQuery()
			|  ~~~~~~~~~~~~~~~~~~~~~~~
			| Exception calling "ExecuteQuery" with "0" argument(s): "One or more field types are not installed
			| properly. Go to the list settings page to delete these fields."

This url can be used to export column schema and confirm the internal name. 
    https://{clienttenant}.sharepoint.com/sites/{SitePath}/_vti_bin/owssvr.dll?Cmd=ExportList&List={ListGUID}

Dependencies
	-	PnP.PowerShell Modules

Call script with parameters
	.\Cat.PnP-DeleteHiddenColumn.ps1  -inputFile Clients\Cat.MT-DeployCrisisTeam-Client.xml -client "ClientNameNoSpacesOrAbriviation"

Table of Contents
	1.	Snapin(s), Modules & Environment Connection(s)
	2.	Input & Output Files
	3.	Variables
	4.	Main Script
	5.	CSV Loop Example
	6.	...
#>

#region Script Parameters
Param(	[string] $clientTenant = "{ClientTenant}", # IE. ClientTenant.sharepoint.com
		[string] $clientAppID = "{Client app registration ID from Azure}", # Leave Blank for interactive login
		[string] $clientCertPath = "{Path to client certificate file must be installed on local user}",# Leave Blank for interactive login
        [string] $siteURL = "https://{ClientTenant}.sharepoint.com/sites/{SitePath}",
        [string] $listName = "{ListPath}", # IE. https://{ClientTenant}.sharepoint.com/sites/{SitePath}/{ListPath}/
        [string] $columnName = "{InternalColumnName}" #Internal Name
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
Start-Transcript -Path ("..\_Log Files\{0}.SPO-DeleteHiddenColumn-{1}.log" -f $clientTenant, $scriptStart[1])
#endregion

#region 1.	Snapin(s), Modules & Environment Connection(s)
# Load MS / PnP PowerShell Modules
Get-Module -ListAvailable | Where-Object{($_.Name -like "*SharePoint*") -or ($_.Name -like "*PnP*") } | Remove-Module
Import-Module PnP.PowerShell
#endregion

#region 2.	Variables
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
Write-Host ("   Client Tenant Domain: {0}" -f $clientTenant)
if ($clientAppID) {
	Write-Host ("   App ID: {0}" -f $clientAppID)
	Write-Host ("   Certificate Path: {0}" -f $clientCertPath)
}
Write-Host ("   Site URL: {0}" -f $siteURL)
Write-Host ("   List Name: {0}" -f $listName)
Write-Host ("   Column Name: {0}" -f $columnName)
if ($clientAppID) {
	Write-Host
	Write-Host "User Inputs" -ForegroundColor Cyan
	Write-Host ("   Certificate Thumbprint: {0}" -f $cert.Thumbprint)
}
Write-Host
Write-Host "Script Variables" -ForegroundColor Cyan
Write-Host ("   Microsoft Tenant: {0}" -f $tenantDomain)
Write-Host ("   SPO admin URL: {0}" -f $spoAdminUrl)
#endregion
    
#region 3.	Main Script
#region Connecting To SP Online
Write-Host
Write-Host "Connecting to Services" -ForegroundColor Cyan
Write-Host ("   Connecting to SharePoint Online Administraton: {0}" -f $siteURL)
try {
	if ($clientAppID) {
		Connect-PnPOnline $siteURL -tenant $tenantDomain -ClientId $clientAppID -Thumbprint $cert.Thumbprint -WarningAction SilentlyContinue # PnP Modules w/ App Reg
	} else {
		Connect-PnPOnline -Url $spoAdminURL -Interactive -WarningAction SilentlyContinue # PnP Modules w/ App Reg
	} 
}
catch {
	Write-Host
	throw ("Unable to connect: {0}" -f $_.Exception)	
}
#endregion

#Get the CSOM Context
$context = Get-PnPContext

#Get the Field from List
$field = Get-PnPField -List $listName -Identity $columnName

#Update Column Settings
$field.ReadOnlyField = $false
$field.Sealed = $false
$field.Update()
$context.ExecuteQuery()

#Delete column
$field.DeleteObject()
$field.Update()
$context.ExecuteQuery()
#endregion

#region Completed Script
# Disconnect from Services
Write-Host
Write-Host "Disconnecting from Services:" -ForegroundColor Cyan
Write-Host ("   Disconnecting from SharePoint Online: {0}" -f $siteURL)
Disconnect-PnPOnline

# Output Script Info
Complete-CatScript $scriptStart[0]

# Stop logging
Stop-Transcript
#endregion