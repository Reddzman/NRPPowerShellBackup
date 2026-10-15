<#  Cat.SP13-SiteGroupMembership.ps1

This script scans group membership for the provided site collection and outputs it to CSV.  

Dependencies
	-	SharePoint 2013 Module
	-	General Functions - ..\General\Cat.PnP-GeneralFunctionsFunctions.psm1

Call script with parameters
	.\Cat.SP13-SiteGroupMembership.ps1  -siteURL https://sp13.client.com -client "ClientNameNoSpacesOrAbriviation"

Table of Contents
	1.	Input & Output Files
	2.	Variables
	3.	Snapin(s), Modules & Environment Connection(s)
	4.	Main Script
#>

#region Script Parameters
Param(	[string] $siteURL = "https://sharepoint.client.com/sites/site",
		[string] $client = "Testing"
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
Start-Transcript -Path ("..\_Log FIles\{0}.SP13-SiteGroupMembership-{1}.log" -f $client, $scriptStart[1])
#endregion

#region 1.	Input & Output Files
# File Paths
$outPath = ("..\_Output Files\{0}-SiteGroupMembership-{1}.csv" -f $client, $scriptStart[1])

# Create Output File(s)
Out-File -FilePath $outPath -InputObject "siteURL|siteName|webURL|webName|groupName|groupLoginName|groupOwner|groupID|groupRoles|userDisplayName|userEmail|userLoginName|userRoles" -Encoding unicode
#endregion

#region 2.	Variables
# Output Parameters & Variables
Write-Host
Write-Host "Script Parameters" -ForegroundColor Cyan
Write-Host ("   Site URL: {0}" -f $siteURL)
Write-Host ("   Client Abreviation: {0}" -f $client)
#endregion

#region 3.	Snapin(s), Modules & Environment Connection(s)
# Load MS / PnP PowerShell Modules
Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue
#endregion
    
#region 4.	Main Script
# Screen Output
Write-Host
Write-Host ("Scanning site collection: {0}" -f $siteURL) -ForegroundColor Magenta

# Get Site Collection
Write-Host ("   Retrieving site and web objects: {0}" -f $siteURL)
$site = Get-SPSite -Identity $siteURL
$webs = $site.AllWebs

# Replacing "|" charecters where necessary in Site Info
If ($site.RootWeb.Title) {$siteName = $site.RootWeb.Title.Replace("|","\")} Else {$siteName = ""}

# Looping through Web objects
ForEach ($web in $webs) {
	# Get all Groups for web
	Write-Host ("   Retrieving groups for web: {0}" -f $web.Url)
	$groups = $web.groups

	# Replacing "|" charecters where necessary in Site Info
	If ($web.Title) {$webName = $web.Title.Replace("|","\")} Else {$webName = ""}

	# Looping through Groups
	ForEach ($group in $groups) {
		# Get all users for group
		Write-Host ("      Retrieving users for group: {0}" -f $group.Name) -ForegroundColor DarkGray
		$users = $group.users

		# Replacing "|" charecters where necessary in Group Info
		If ($group.Name) {$groupName = $group.Name.Replace("|","\")} Else {$groupName = ""}
		If ($group.LoginName) {$groupLoginName = $group.LoginName.Replace("|","\")} Else {$groupLoginName = ""}
		If ($group.Owner.LoginName) {$groupOwner = $group.Owner.LoginName.Replace("|","\")} Else {$groupOwner = ""}
		If ($group.Roles.Name) {$groupRoles = $group.Roles.Name.Replace("|","\") -join ","} Else {$groupRoles = ""}

		# Looping through Users
		ForEach ($user in $users) {
			# Screen Output
			Write-Host ("         Writing user record: {0}" -f $user.DisplayName) -ForegroundColor Blue

			# Replacing "|" charecters where necessary in User Info
			If ($user.DisplayName) {$userDisplayName = $user.DisplayName.Replace("|","\")} Else {$userDisplayName = ""}
			If ($user.Email) {$userEmail = $user.Email.Replace("|","\")} Else {$userEmail = ""}
			If ($user.LoginName) {$userLoginName = $user.LoginName.Replace("|","\")} Else {$userLoginName = ""}
			If ($user.Roles.Name) {$userRoles = $user.Roles.Name.Replace("|","\") -join ","} Else {$userRoles = ""}

			# Write user record
			Out-File -FilePath $outPath -Append -InputObject ("{0}|{1}|{2}|{3}|{4}|{5}|{6}|{7}|{8}|{9}|{10}|{11}|{12}" -f
				$site.Url,
				$siteName,
				$web.Url,
				$webName,
				$groupName,
				$groupLoginName,
				$groupOwner,
				$group.Id,
				$groupRoles,
				$userDisplayName,
				$userEmail,
				$userLoginName,
				$userRoles
			)
		}
	}

	# Screen Output
	Write-Host ("      Web Complete: {0}" -f $web.Url) -ForegroundColor Green
}
#endregion

#region Completed Script
# Output Script Info
Complete-CatScript $scriptStart[0]

# Stop logging
Stop-Transcript
#endregion