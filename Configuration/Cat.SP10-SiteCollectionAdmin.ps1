<#
Cat.SP10_PreMigScanWebAppADGroups.ps1

This script iterates through all site collections in a farm adding a single user to site collection administrators.

Table of Contents
	1.	Variables
	3.	Snapin(s), Modules & Environment Connection(s)
	4.	Main Script
#>

#region Start Date / Time
# Set PS location
$PSScriptRoot = Split-Path $MyInvocation.MyCommand.Path -Parent
Set-Location -Path $PSScriptRoot

# Get Script Parent Location
$LocSplit = $PSScriptRoot.Split("\")
$LocReplace = $LocSplit[($LocSplit.Count - 1)]
$Location = $PSScriptRoot.Replace($LocReplace,"")

# Remove Existing Catapult Modules to ensure PS session has most up to date functions
Get-Module Cat.* | Remove-Module

# Load General Functions
Import-Module ("{0}General\Cat.GeneralFunctions.psm1" -f $Location)

# Get Start Date / Time
$Start = Start-CatScript
#endregion

#region 1.	Variables
# User Input Variables
Write-Host
Write-Host "User Input Variables" -ForegroundColor DarkCyan
$user = Read-Host -Prompt "   Enter the user you would like to add (domain\username)" 
#endregion

#region 3.	Snapin(s), Modules & Environment Connection(s)
# Load SharePoint PowerShell Snap-In(s)
Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue
#endregion

#region 4.	Main Script
# Screen Output
Write-Host
Write-Host "Adding Site Collection Admins" -ForegroundColor DarkCyan

# Get Web Application Objects
Write-Host "   Retrieving Web Applciations" -ForegroundColor Magenta	
$WebApps = Get-SPWebApplication

# Loop through all web applications in the farm
ForEach ($WebApp in $WebApps)
{
	# Screen Output for web application
	Write-Host
	Write-Host ("   Web Application: {0}" -f $WebApp.URL) -ForegroundColor Magenta	
	
	# Get site collection objects
	Write-Host "      Retrieving Site Collections"
	$Sites = Get-SPSite -WebApplication $WebApp.URL -limit all

	# Loop through site collections for web application
	ForEach ($Site in $Sites)
	{
		# Screen Output for Site
		Write-Host ("      Adding Site Collection Admin: {0}" -f $Site.URL)

		# Add user to Site Collection
		Set-SPSite -Identity $Site.URL -SecondaryOwnerAlias $user
		
		# Dispose of site object
		$Site.Dispose()
	} # ForEach ($Site in $Sites)
} # ForEach ($WebApp in $WebApps)
# Screen Output for Farm
Write-Host "   Farm Configuration Complete" -ForegroundColor Green

#endregion

#region Completed Screen Output
Complete-CatScript $Start[0]
#endregion