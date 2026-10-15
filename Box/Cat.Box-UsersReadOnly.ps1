<#  Cat.Box-UsersReadOnly.ps1

This script iterates through CSV to set a group of users read only in Box.

Dependencies
	-	../General/Cat.GeneralFunctions.psm1
    -	Input CSV File in Input Files folder - ..\_Input Files\Cat.Box-UserAccounts-Test.csv
    -   Box PowerShell SDK for Box API v2 - https://github.com/box-community/sdk-powershell-api-v2

Call script with parameters
	.\Cat.Box-UsersReadOnly.ps1 -usersFile "..\_Input Files\Cat.Box-ReadOnly-Test.csv"

CSV Example
	Site,siteURL,SPOUrl
	SiteName,https://sharepoint.client.com/sitedirectory/law,https://client.sharepoint.com/sites/law

Table of Contents
	1.	Input & Output Files
	2.	Variables
	3.	Snapin(s), Modules & Environment Connection(s)
	4.	Main Script
#>

#region Script Parameters
Param(	[String] $usersFile = "Cat.Box-ReadOnly-Test.csv",
		[boolean] $readonly = $true)
#endregion

#region Start Date / Time
# Remove Existing Catapult Modules to ensure PS session has most up to date functions
Get-Module Cat.* | Remove-Module

# Load General Functions
Import-Module "..\General\Cat.GeneralFunctions.psm1"

# Get Start Date / Time
$scriptStart = Start-CatScript

# Start Transcript
Start-Transcript -Path ("..\_Log FIles\BoxUsersReadOnly-{0}.log" -f $scriptStart[1])
#endregion

#region 1.	Input & Output Files
# File Paths
$inPath = ("..\_Input Files\{0}" -f $usersFile)

# Import Input File(s)
$inFile = Import-Csv $inPath -Encoding Unicode  # Importing CSV File
#endregion

#region 2.	Variables
# Script Variables
if ($readonly){
	$updateHash = @{id = "1234"; status = "cannot_delete_edit_upload";}
}
else {
	$updateHash = @{id = "1234"; status = "active";}
}

# Output Parameters & Variables
Write-Host
Write-Host "Script Parameters" -ForegroundColor Cyan
Write-Host ("   Users File: {0}" -f $usersFile)
Write-Host
Write-Host "Script Variables" -ForegroundColor Cyan
Write-Host ("   Status Update: {0}" -f $updateHash.status)
#endregion #
    
#region 4.	Main Script
# looping through CSV file
Write-Host
Write-Host ("Updated users in user file: {0}" -f $usersFile) -ForegroundColor Magenta
Write-Host "   Authenticating to Box"
connect-box
foreach ($user in $inFile) {
	Write-Host ("   User: {0}" -f $user.email)
	$boxUser = get-boxusers -SearchString $user.email
	$updateHash.id = $boxUser.Entries.ID
	$updateOutput = set-boxuser -updateinfo $updateHash
}

# Success Screen Output
Write-Host
Write-Host "Users successfully updated" -ForegroundColor Green
#endregion

#region Completed Script
# Stop Logging
Stop-Transcript

# Output Script Info
Complete-CatScript $scriptStart[0]
#endregion