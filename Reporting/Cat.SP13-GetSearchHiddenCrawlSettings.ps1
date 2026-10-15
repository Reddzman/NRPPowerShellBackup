<#
Cat.SP13-GetSearchHiddenCrawlSettings.ps1

This script loops through all sites in a web application to determine
if sites, Lists, or libraries have been removed from the search index.

Dependencies
	-	PnP PowerShell https://msdn.microsoft.com/en-us/pnp_powershell/pnp-powershell-overview
	-	SharePoint 2013 PowerShell Module

Table of Contents
	1.	Variables
	2.	Snapin(s), Modules & Environment Connection(s)
	3.	Main Script
	4.	Check Web Crawl Setting
	5.	Check List Crawl Setting
#>

#region Start Date / Time
# Set PS location
Set-Location $PSScriptRoot

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
Write-Host "   Enter credentials for the reporting site:"
$ReportCred = Get-Credential -Message "Enter your credentials for SharePoint Online reporting site:" 
$ReportURL = Read-Host -Prompt "   Enter URL for SharePoint Online reporting site"
$WebAppURL = Read-Host -Prompt "   Enter URL for SharePoint 2013 Web Application to scan"

# Script Variables
$SkipLists = @("Workflow Tasks";"Content and Structure Reports";"Reusable Content";"Style Library";"Customized Reports";)
#endregion 1.	Variables

#region 2.	Snapin(s), Modules & Environment Connection(s)
# Load SharePoint PowerShell Snap-In(s)
Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue

# Load Custom Modules
Import-Module ("{0}Reporting\Cat.ReportingFunctions.psm1" -f $Location)
#endregion

#region 3.	Main Script
Write-Host
Write-Host ("Scanning for Hidden Search Settings: {0}" -f $WebAppURL) -ForegroundColor DarkCyan

# Get site collection object
Write-Host "   Retrieving Site Collections"
$Sites = Get-SPSite -WebApplication $WebAppURL

# Looping through Sites in CSV Input
ForEach ($Site in $Sites)
{
	# Output Site to Screen
	Write-Host
	Write-Host ("   Scanning Site Collection: {0}" -f $Site.RootWeb.Title) -ForegroundColor Magenta

	#Loop through sub sites
	ForEach ($Web in $Site.AllWebs)
	{
		# Screen Output
		Write-Host ("      Sub Site: {0}" -f $Web.Title) -ForegroundColor DarkGray -NoNewline
		# 4.	Check Web Crawl Setting
		If ($Web.NoCrawl -eq $true)
		{
			Write-Host " - Hidden" -ForegroundColor Yellow

			# Add to SharePoint List
				$Values = @{"MigSCName"=$Site.RootWeb.Title;"MigSCURL"=$Site.Url;"MigWebName"=$Web.Title;"MigWebURL"=$Web.Url;}
				$Temp = Add-CatMigrationRecord -Url $ReportURL -Creds $ReportCred -List "Migration Hidden From Crawl" -ContentType "Migration Hidden Crawl Item" -Values $Values
		} # If ($Web.NoCrawl -eq $true)
		Else 
		{
			Write-Host
		}

		# Loop through all lists and libraries in a sub site
		$Lists = $Web.Lists | ?{($_.Hidden -eq $false) -and ($_.Title -notin $SkipLists) -and ($_.NoCrawl -eq $true)}
		ForEach ($List in $Lists)
		{				
			# Screen Output for List
			Write-Host ("         Hidden List: {0}" -f $List.Title)
			
			# Add to SharePoint List
			$Values = @{"MigSCName"=$Site.RootWeb.Title;"MigSCURL"=$Site.Url;"MigWebName"=$Web.Title;"MigWebURL"=$Web.Url;"MigListName"=$List.Title;"MigListURL"=$List.URL;}
			$Temp = Add-CatMigrationRecord -Url $ReportURL -Creds $ReportCred -List "Migration Hidden From Crawl" -ContentType "Migration Hidden Crawl Item" -Values $Values
		} # ForEach ($List in $CurWeb.Lists)
		
		# Screen Output & Dispose of Site
		Write-Host ("         Sub Site Complet: {0}" -f $Web.Title) -ForegroundColor Green
		$Web.Dispose()
	} # ForEach ($Web in $Site.AllWebs)
	# Screen Output & Dispose of Site
	Write-Host ("      Site Complete: {0}" -f $Site.RootWeb.Title) -ForegroundColor Green
	$Site.Dispose()
} # ForEach ($Site in $Sites)
#endregion 3.	Main Script

#region Completed Screen Output
Complete-CatScript $Start[0]
#endregion