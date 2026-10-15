<#  Cat.PnP-ConfigSPOGroupMembership.ps1

This script iterates through a CSV file to add users to SharePoint Groups (Classic/Communication sites) in SPO. The deleteExistingMembership parameter can be leverage to have the script remove
all existing users from each unique Group/Web combination in the CSV file. The #TODO:Reporting\Cat.SP13-SiteGroupmembership.ps1 script can be leveraged to extract 
the CSV from a existing site.  This scripts initial purpose was to help configure permissions for UAT and production release testing.

Dependencies
	-	PnP.PowerShell
	-	General Functions - ..\General\Cat.PnP-GeneralFunctionsFunctions.psm1
	-	Input CSV File - Example: ..\_Input FIles\Cleint\Cat.Tech-Action-Client.csv

Call script with parameters
	.\Cat.PnP-ConfigSPOGroupMembership.ps1  -inputFile "Samples\Configuration\Cat.PnP-ConfigSPOGroupMembership-test.csv" -deletedExistingMembership $true -createMissingGroups $true -clientAppID "caf6b11a-cf03-416c-b0f0-67a09a8e30ee" -clientCertPath "C:\GitHub\NRPDev\AppRegistration\PowerShell App Registration.pfx" -clientTenant "nrpdev2"

Table of Contents
	1.	Input & Output Files
	2.	Variables
	3.	Snapin(s), Modules & Environment Connection(s)
	4.	Main Script
	5.	Delete Group Membership
	6.	Updating/Adding Group Membership
#>

#region Script Parameters
Param(	[string] $inputFile = "Samples\Configuration\Cat.PnP-ConfigSPGroupMembership-test.csv",
		[bool] $deleteExistingMembership = $true,
		[bool] $createMissingGroups = $true,
		[string] $clientAppID = "[ClientAppID]",
		[string] $clientCertPath = "c:\clientcert\clientcert.pfx",
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
Start-Transcript -Path ("..\_Log FIles\{0}.ConfigSPGroupMembership-{1}.log" -f $clientTenant, $scriptStart[1])
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
# User Input
if ($deleteExistingMembership) {
	# Get confirmation to remove permissions
	Write-Host
	Write-Host "*** Deleting existing group membership will permanantly remove all users permission that are being provided by the groups in this CSV. ***" -ForegroundColor Yellow
	$userProceed = Read-Host "   Are you sure you want to proceed with deleting existing group membership (Yes or No)" 
	if ($userProceed.ToLower() -eq "yes") {
		$deleteExistingMembership = $true
	}
	else {
		$deleteExistingMembership = $false
	}
}
Write-Host
Write-Host "User Input Variables" -ForegroundColor Cyan
$certPass = ConvertTo-SecureString -String (Read-Host -Prompt "   Enter certificate password for the app registration") -AsPlainText -Force

# Script Variables
if ($deleteExistingMembership) {$groupAction = "Updating"} Else {$groupAction = "Adding"}
$tenantDomain = ("{0}.onmicrosoft.com" -f $clientTenant)
$cert = Get-PnPAzureCertificate -Path $clientCertPath -Password $certPass

# Output Parameters & Variables
Write-Host
Write-Host "User Input" -ForegroundColor Cyan
Write-Host ("   Delete Membership Confirmation: {0}" -f $userProceed)
Write-Host
Write-Host "Script Parameters" -ForegroundColor Cyan
Write-Host ("   Input File: {0}" -f $inputFile)
Write-Host ("   Delete Existing Group Membership: {0}" -f $deleteExistingMembership)
Write-Host ("   Create Groups if Missing: {0}" -f $createMissingGroups)
Write-Host ("   App ID: {0}" -f $clientAppID)
Write-Host ("   App Certificate: {0}" -f $clientCertPath)
Write-Host ("   Client Tenant: {0}" -f $clientTenant)
Write-Host
Write-Host "Script Variables" -ForegroundColor Cyan
Write-Host ("   Action: {0}" -f $groupAction)
#endregion
    
#region 4.	Main Script
#region 5.	Delete Group Membership
if ($deleteExistingMembership) {
	# Screen Output
	Write-Host
	Write-Host "Deleting Group Membership" -ForegroundColor Magenta

	# Looping through all unique groups in CSV
	$groups = $inFile | Select-Object -Property groupName, siteURL -Unique
	foreach ($group in $groups) {
		Write-Host ("   Removing Users:   Site: {0}   Group: {1}" -f $group.siteURL, $group.groupName)

		# Connecting To SP Online
		Write-Host ("      Connecting to site: {0}" -f $group.siteURL) -ForegroundColor Blue
		Connect-PnPOnline $group.siteURL -tenant $tenantDomain -ClientId $clientAppID -Thumbprint $cert.Thumbprint

		# Retreiving Group
		Write-Host ("      Retrieviing Group: {0}" -f $group.groupName) -ForegroundColor Blue -NoNewline
		try {
			$tempGroup = Get-PnPGroup -Identity $group.groupName -ErrorAction Stop
			Write-Host " - Group Found"  -ForegroundColor Blue
		}
		catch {
			Write-Host " - Group Not Found" -ForegroundColor Yellow
		}

		# Get Group Membership
		if($tempGroup){
			$deleteUsers = Get-PnPGroupMember -Group $tempGroup.Id
			foreach ($user in $deleteUsers) {
				Write-Host ("         Removing User: {0}" -f $User.Email) -ForegroundColor DarkGray
				Remove-PnPGroupMember -LoginName $user.LoginName -Group $tempGroup.Id
			}
		}

		# Disconnecting from SP Online
		Disconnect-PnPOnline
	}

	# Screen Output
	Write-Host "   Deleting Membership Complete" -ForegroundColor Green
}
#endregion

#region 6.	Updating/Adding Group Membership
Write-Host
Write-Host ("{0} Group Membership" -f $groupAction) -ForegroundColor Magenta

# looping through CSV file
foreach ($user in $inFile) {
	# Connecting To SP Online
	Write-Host
	Write-Host ("   Connecting to site: {0}" -f $user.siteURL)
	Connect-PnPOnline $user.siteURL -tenant $tenantDomain -ClientId $clientAppID -Thumbprint $cert.Thumbprint

	# Retreiving Group
	Write-Host ("   Retrieviing Group: {0}" -f $user.groupName) -NoNewline
	try {
		$tempGroup = Get-PnPGroup -Identity $user.groupName -ErrorAction Stop
		$groupFound = $true
		Write-Host " - Group Found"
	}
	catch {
		$groupFound = $false
		Write-Host " - Group Not Found" -ForegroundColor Yellow
	}

	# Create Group if it doesn't exist
	if (-not($groupFound)) {
		try {
			Write-Host ("   Creating Group: {0}" -f $user.groupName)
			$tempGroup = New-PnPGroup -Title $user.groupName -Owner $user.groupOwner
			$groupFound = $true
		}
		catch {
			$groupFound = $false
		}
	}

	# Adding user to group
	if ($groupFound) {
		Write-Host ("   Adding User: {0}" -f $user.username)
		$tempUser = Add-PnPGroupMember -LoginName $user.username -Group $tempGroup.Id
	}

	# Disconnecting from SP Online
	Disconnect-PnPOnline
}

# Screen Output
Write-Host ("   {0} Membership Complete" -f $groupAction) -ForegroundColor Green
#endregion

#endregion

#region Completed Script
# Output Script Info
Complete-CatScript $scriptStart[0]

# Stop logging
Stop-Transcript
#endregion