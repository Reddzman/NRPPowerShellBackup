<#  Cat.SP13-UpdateSiteGroupMembership.ps1

This script iterates through a CSV file to add users to SharePoint Groups. The deleteExistingMembership parameter can be leverage to have the script remove
all existing users from each unique Group/Web combination in the CSV file. The Reporting\Cat.SP13-SiteGroupmembership.ps1 script can be leveraged to extract 
the CSV from a existing site.  The CSV is "|" delimited.  This scripts initial purpose was to help configure permissions for UAT and production release testing.

Dependencies
	-	SharePoint 2013 Module
	-	General Functions - ..\General\Cat.PnP-GeneralFunctionsFunctions.psm1
	-	Input CSV File - Example: ..\_Input FIles\Cleint\Cat.Tech-Action-Client.csv

Call script with parameters
	.\Cat.SP13-UpdateSiteGroupMembership.ps1  -inputFile Cat.SP13-SiteGroupmembership-Client-Environment.xml -client "ClientNameNoSpacesOrAbriviation" -deletedExistingMembership $false

Table of Contents
	1.	Input & Output Files
	2.	Variables
	3.	Snapin(s), Modules & Environment Connection(s)
	4.	Main Script
	5.	Delete Group Membership
	6.	Updating/Adding Group Membership
#>

#region Script Parameters
Param(	[string] $inputFile = "Testing\Cat.SP13-SiteGroupMembership-Test.xml",
		[string] $client = "Testing",
		[boolean] $deleteExistingMembership = $false
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
Start-Transcript -Path ("..\_Log FIles\{0}.SP13-UpdateSiteGroupMembership-{1}.log" -f $client, $scriptStart[1])
#endregion

#region 1.	Input & Output Files
# File Paths
$inPath = ("..\_Input Files\{0}" -f $inputFile)

# Import Input File(s)
# Load CSV file
If (Test-Path ($inPath)){
	$inFile = Import-Csv $inPath -Encoding Unicode -Delimiter "|"
}
else {
	Write-Host
	throw ("The input file does not exist: {0}" -f $inPath)	
}
#endregion

#region 2.	Variables
# User Input
if ($deleteExistingMembership) {
	# Get confirmation to remove permissions
	Write-Host
	Write-Host "*** Deleting existing group membership will permanantly remove all users permission that are being provided by the groups in this CSV. ***" -ForegroundColor Yellow
	$userProceed = Read-Host "   Are you sure you want to proceed with deleting existing group membership (Yes or No)" 

	# if user chose to proceed
	if ($userProceed.ToLower() -eq "yes") {
		$deleteExistingMembership = $true
	}
	else {
		$deleteExistingMembership = $false
	}
}

# Script Variables
if ($deleteExistingMembership) {$groupAction = "Updating"} Else {$groupAction = "Adding"}

# Output Parameters & Variables
Write-Host
Write-Host "User Input" -ForegroundColor Cyan
Write-Host ("   Delete Membership Confirmation: {0}" -f $userProceed)
Write-Host
Write-Host "Script Parameters" -ForegroundColor Cyan
Write-Host ("   Input File: {0}" -f $inputFile)
Write-Host ("   Client: {0}" -f $client)
Write-Host ("   Delete Existing Group Membership: {0}" -f $deleteExistingMembership)
Write-Host
Write-Host "Script Variables" -ForegroundColor Cyan
Write-Host ("   Action: {0}" -f $groupAction)
#endregion

#region 3.	Snapin(s), Modules & Environment Connection(s)
# Load MS / PnP PowerShell Modules
Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue
#endregion
    
#region 4.	Main Script
#region 5.	Delete Group Membership
if ($deleteExistingMembership) {
	# Screen Output
	Write-Host
	Write-Host "Deleting Group Membership" -ForegroundColor Magenta

	# Looping through all unique groups in CSV
	$groups = $inFile | Select-Object -Property groupLoginName, webURL -Unique
	foreach ($group in $groups) {
		Write-Host ("   Removing Users:   Web: {0}   Group: {1}" -f $group.webURL, $group.groupLoginName)

		# Get Group Membership
		$deleteUsers = Get-SPUser -Web $group.webURL -Group $group.groupLoginName
		foreach ($user in $deleteUsers) {
			Write-Host ("      Removing User: {0}" -f $User.DisplayName) -ForegroundColor DarkGray
			Remove-SPUser -Identity $user.UserLogin -Web $group.webURL -Group $group.groupLoginName -Confirm:$false
		}
	}

	# Screen Output
	Write-Host
	Write-Host "   Deleting Membership Complete" -ForegroundColor Green
}
#endregion

#region 6.	Updating/Adding Group Membership
Write-Host
Write-Host ("{0} Group Membership" -f $groupAction) -ForegroundColor Magenta

# looping through CSV file
foreach ($user in $inFile) {
	# Get Existing User
	Write-Host ("   Retrieving User: {0} - " -f $user.userDisplayName) -NoNewline
	try {
		$spUser = Get-SPUser -Web $user.webURL -Identity $user.userLoginName.Replace("\","|") -ErrorAction Stop
		$userExists = $true
		Write-Host "User Found - " -NoNewline
	}
	catch {
		$userExists = $false
		Write-Host "User Not Found - " -NoNewline
	}

	# Create New User or Add user to group
	if ($userExists) {
		# Updating existing user
		Set-SPUser -Web $user.webURL -Identity $user.userLoginName.Replace("\","|") -Group $user.groupLoginName
		Write-Host ("Added user to Group: {0}" -f $user.groupLoginName)
	}
	else {
		# Creating new users
		$spUser = New-SPUser -Web $user.webURL -UserAlias $user.userLoginName.Replace("\","|") -Group $user.groupLoginName
		Write-Host ("Created user in Group: {0}" -f $user.groupLoginName)
	}
}

# Screen Output
Write-Host
Write-Host ("   {0} Membership Complete" -f $groupAction) -ForegroundColor Green
#endregion
#endregion

#region Completed Script
# Output Script Info
Complete-CatScript $scriptStart[0]

# Stop logging
Stop-Transcript
#endregion