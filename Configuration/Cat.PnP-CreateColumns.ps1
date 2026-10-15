<#  Cat.PnP-CreateColumns.ps1

This script iterates through a CSV File to create columns in a SharePoint Online site collection.

Dependencies
	-	PnP PowerShell https://msdn.microsoft.com/en-us/pnp_powershell/pnp-powershell-overview
	-	Input CSV File in Input Files folder - inputFile - "Cat.SP13-CreateColumns-Test.csv"

Call script with parameters
	.\Cat.PnP-CreateColumns.ps1 -inputFile "..\..\NRPDev\Clients\AZDPS\SAFE\AZDPS-SAFE2.0-Columns.csv" -siteURL "https://nrpdev2.sharepoint.com/teams/Safe20Testing01" -clientAppID "caf6b11a-cf03-416c-b0f0-67a09a8e30ee" -clientCertPath "C:\GitHub\NRPDev\AppRegistration\PowerShell App Registration.pfx" -clientTenant "nrpdev2"

CSV Example
	displayName,internalName,groupName,type
	Test Text 01,TestText01,Test Columns,Text

Table of Contents
	1.	Snapin(s), Modules & Environment Connection(s)
	2.	Input & Output Files
	3.	Variables
	4.	Main Script
#>

#region Script Parameters
Param(	[String] $inputFile = "Cat.PnP-CreateColumns-Test.csv",
		[String] $siteURL = "https://client.sharepoint.com/sites/columnstesting",
		[String] $clientAppID = "[ClientAppID]",
		[String] $clientCertPath = "c:\clientcert\clientcert.pfx",
		[string] $clientTenant = "client"
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
Start-Transcript -Path ("..\_Log FIles\{0}-CreateColumns-{1}.log" -f $clientTenant, $scriptStart[1])
#endregion

#region 1.	Snapin(s), Modules & Environment Connection(s)
# Load MS / PnP PowerShell Modules
Get-Module -ListAvailable | Where-Object{($_.Name -like "*SharePoint*") -or ($_.Name -like "*PnP*") } | Remove-Module
Import-Module PnP.PowerShell
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
	throw ("The input file does not exist: {0}" -f $inPath)	
}
#endregion

#region 3.	Variables
# User Inputs
Write-Host
Write-Host "User Input Variables" -ForegroundColor Cyan
$certPass = ConvertTo-SecureString -String (Read-Host -Prompt "   Enter certificate password for the app registration") -AsPlainText -Force

# Script Variables
$tenantDomain = ("{0}.onmicrosoft.com" -f $clientTenant)
$cert = Get-PnPAzureCertificate -Path $clientCertPath -Password $certPass

# Output Parameters & Variables
Write-Host
Write-Host "Script Parameters" -ForegroundColor Cyan
Write-Host ("   Input File: {0}" -f $inputFile)
Write-Host ("   Site URL: {0}" -f $siteURL)
Write-Host ("   App ID: {0}" -f $clientAppID)
Write-Host ("   Certificate Path: {0}" -f $clientCertPath)
Write-Host ("   Client Tenant Domain: {0}" -f $clientTenant)
Write-Host
Write-Host "User Inputs" -ForegroundColor Cyan
Write-Host ("   Certificate Password: {0}" -f $certPass)
Write-Host
Write-Host "Script Variables" -ForegroundColor Cyan
Write-Host ("   Microsoft Tenant: {0}" -f $tenantDomain)
#endregion
    
#region 4.	Main Script
Write-Host
Write-Host ("Configuring Columns: {0}" -f $inputFile) -ForegroundColor Magenta
# Connecting To SP Online
Write-Host ("   Connecting to SharePoint Online site: {0}" -f $siteURL)
Connect-PnPOnline -URL $siteURL -tenant $tenantDomain -ClientId $clientAppID -Thumbprint $cert.Thumbprint

# Loop Through CSV File
foreach ($column in $inFile) {
	# Creating Column
	Write-Host ("   Creating Column: {0}" -f $column.displayName)
	$temp = Add-PnPField -DisplayName $column.displayName -InternalName $column.internalName -Group $column.groupName -Type $column.type
}
#endregion

#region Completed Script`
# Disconnect from Services
Write-Host ("   Disconnecting from SharePoint Online: {0}" -f $siteURL)
Disconnect-PnPOnline

# Output Script Info
Complete-CatScript $scriptStart[0]

# Stop logging
Stop-Transcript
#endregion