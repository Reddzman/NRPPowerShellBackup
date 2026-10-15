<#  Cat.SPO-AddSiteCollectionAdmin.ps1

This script iterates through all site collections in SharePoint 2013 and add a specified user to the Site Collection Administrators
for that site collection.  It requires that the user running the script be a site collection admin of ALL site collecitons.  If you don't
have a user with that permission, you can use Cat.SP13UpdateSecondaryOwner.ps1 to add yourself to those site collections.

Dependencies
	-	SharePoint Server Modules

Call script with parameters
	.\Cat.SP13-AddSiteCollectionAdmin.ps1 -client Testing -adminAcct sp13test\amansle -includeCentralAdmin $false -skipWebAppsURLs @("https://mysites.sp13test.local/";"https://other.client.com/";) -isSiteCollectionAdmin $true

Table of Contents
	1.	Input & Output Files
	2.	Variables
	3.	Snapin(s), Modules & Environment Connection(s)
	4.	Main Script
#>

#region Script Parameters
Param(	
		[string] $client = "{client}",
		[string] $adminAcct = "domain\username",
		[boolean] $includeCentralAdmin = $false,
		$skipWebAppsURLs = @("https://my.client.com";"https://other.client.com";),
		[boolean] $isSiteCollectionAdmin = $true

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
Start-Transcript -Path ("..\_Log FIles\{0}.SP13-AddSiteCollectionAdmin-{1}.log" -f $client, $scriptStart[1])
#endregion

#region 2.	Variables
# Script Variables
If($isSiteCollectionAdmin){$adminAction = "Adding"} Else {$adminAction = "Removing"}


# Output Parameters & Variables
Write-Host
Write-Host "Script Parameters" -ForegroundColor Cyan
Write-Host ("   Client: {0}" -f $client)
Write-Host ("   Admin User: {0}" -f $adminAcct)
Write-Host ("   Include Central Admin: {0}" -f $includeCentralAdmin)
Write-Host ("   Skip Web Applications: {0}" -f $skipWebAppsURLs)
Write-Host ("   Is Site Collection Admin: {0}" -f $isSiteCollectionAdmin)
Write-Host
Write-Host ("Skip Web Applications") -ForegroundColor Cyan
$skipWebAppsURLs
Write-Host
Write-Host "Script Varialbes" -ForegroundColor Cyan
Write-Host ("   Action: {0}" -f $adminAction)
#endregion

#region 3.	Snapin(s), Modules & Environment Connection(s)
# Load MS / PnP PowerShell Modules
Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue
#endregion
    
#region 4.	Main Script
# Screen Output
Write-Host
Write-Host ("{0} {1} as a Site Collection Administrator" -f $adminAction, $adminAcct) -ForegroundColor Cyan
# Get Web Application Objects
Write-Host "   Retrieving Web Applciations"	
$WebApps = Get-SPWebApplication -IncludeCentralAdministration:$includeCentralAdmin

# Loop through all web applications in the farm
ForEach ($WebApp in $WebApps)
{
	# Screen Output for web application
	Write-Host
	Write-Host ("   Web Application: {0}" -f $WebApp.URL) -ForegroundColor Magenta

	if ($WebApp.URL -notin $skipWebAppsURLs) {
		# Get site collection objects
		Write-Host "      Retrieving Site Collections"
		$Sites = Get-SPSite -WebApplication $WebApp.URL -limit all

		# Loop through site collections for web application
		ForEach ($Site in $Sites)
		{
			# Screen Output for Site
			Write-Host ("      {0} Site Collection Admin: {1}" -f $adminAction, $Site.URL)

			# Get the RootWeb
			$web = Get-SPWeb $Site.URL
			
			# Get the user acount
			$user = $web.EnsureUser($adminAcct)

			# Update Site collection Admin permissions
			$user.IsSiteAdmin = $isSiteCollectionAdmin
			$user.Update()
		} # ForEach ($Site in $Sites)
	}
	else {
		# Get site collection objects
		Write-Host "      Skipping web application" -ForegroundColor Yellow
	}
		
} # ForEach ($WebApp in $WebApps)

# Screen Output
Write-Host
Write-Host ("{0} {1} Complete" -f $adminAction, $adminAcct) -ForegroundColor Green
#endregion

#region Completed Script
# Output Script Info
Complete-CatScript $scriptStart[0]

# Stop logging
Stop-Transcript
#endregion