<# Cat.SP10_PreMigScanWebAppADGroups.ps1

This script generates reports for an entire Farm to identify the following information:
	-	AD Group utilization

Potential Improvements
	-	Output Web Application information as part of CSV.

Table of Contents
	1.	Input & Output Files
	2.	Variables
	3.	Snapin(s), Modules & Environment Connection(s)
	4.	Main Script
	5.	Find AD Groups
	6.	Get AD groups for web
	7.	Get AD Groups for List
	8.	Get AD Groups for Item
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

# Start Transcript
Start-Transcript -Path ("..\_Log Files\ADGroupsScan-{0}.log" -f $Start[1])
#endregion

#region 1.	Input & Output Files
# File Paths
$ADGroupFile = ("{0}_Output Files\Client-ADGroups-{1}.txt" -f $Location, $Start[1])

# Create Output File(s)
Out-File -FilePath $ADGroupFile -InputObject "Site|SiteURL|Web|WebURL|List|ListURL|Item|Group"
#endregion

#region 2.	Variables
# Script Variables
$SkipLists = @("Content and Structure Reports";"Customized Reports";"Form Templates";"Master Page Gallery";"Reusable Content";"Style Library";"Workflow Tasks";)
#endregion

#region 3.	Snapin(s), Modules & Environment Connection(s)
# Load SharePoint PowerShell Snap-In(s)
Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue

# Load Custom Modules
Import-Module ("{0}Reporting\Cat.ReportingFunctions.psm1" -f $Location)
#endregion

#region 4.	Main Script
# Screen Output
Write-Host
Write-Host "Starting Farm Scan" -ForegroundColor DarkCyan

# Get Web Application Objects
Write-Host "   Retrieving Web Applciations" -ForegroundColor Magenta	
$WebApps = Get-SPWebApplication

# Loop through all web applications in the farm
ForEach ($WebApp in $WebApps)
{
	# Screen Output for web application
	Write-Host
	Write-Host ("   Scanning Web Application: {0}" -f $WebApp.URL) -ForegroundColor Magenta	
	
	# Get site collection objects
	Write-Host "      Retrieving Site Collections" -ForegroundColor DarkGreen
	$Sites = Get-SPSite -WebApplication $WebApp.URL -Limit All

	# Loop through site collections for web application
	ForEach ($Site in $Sites)
	{
		# Screen Output for Site
		Write-Host ("      Scanning all sites for site collection: {0}" -f $Site.URL) -ForegroundColor DarkGreen
		if ($Site.RootWeb.Title -eq $null){
			$siteTitle = ""
		}
		else {
			$siteTitle = $Site.RootWeb.Title.Replace("|","\")
		}
		Write-Host "         Retrieving Webs"
		
		# Loop through sub sites
		ForEach ($Web in $Site.AllWebs)
		{
			# Screen output for web
			Write-Host ("         Scanning Web: {0}" -f $Web.Title)
			if ($Web.Title -eq $null){
				$webTitle = ""
			}
			else {
				$webTitle = $Web.Title.Replace("|","\")
			}

			#region 5.	Find AD Groups
			# 6.	Get AD groups for web
			If ($Web.HasUniqueRoleAssignments)
			{
				Write-Host ("            Retrieving Web AD Groups: {0}" -f $Web.Title) -ForegroundColor DarkGray
				Get-CatSP10ADGroups -File $ADGroupFile -Assignments $Web.RoleAssignments -SiteName $siteTitle -SiteUrl $Site.Url -WebName $webTitle -WebUrl $Web.Url
			} # If ($Web.HasUniqueRoleAssignments)
			
			# Loop through all lists and libraries in a sub site
			$Lists = $Web.Lists | Where{($_.Hidden -eq $false) -and ($SkipLists -notcontains $_.Title)}
			ForEach ($List in $Lists)
			{
				if ($List -ne $null)
				{
					# Screen output for list
					Write-Host ("            Scanning List: {0}" -f $List.Title) -ForegroundColor DarkGray
					if ($List.Title -eq $null){
						$listTitle = ""
					}
					else {
						$listTitle = $List.Title.Replace("|","\")
					}
					$ListURL = ("{0}/{1}" -f $Web.Url, $List.RootFolder.Url)
								
					# 7.	Get AD Groups for List
					If ($List.HasUniqueRoleAssignments)
					{
						Write-Host ("               Retrieving List AD Groups: {0}" -f $List.Title) -ForegroundColor Yellow
						Get-CatSP10ADGroups -File $ADGroupFile -Assignments $Web.RoleAssignments -SiteName $siteTitle -SiteUrl $Site.Url -WebName $webTitle -WebUrl $Web.Url -ListName $listTitle -LURL $ListURL
					}

					# Loop through Items in List
					Write-Host "               Scanning Items" -ForegroundColor Yellow
					if ($Item.Name -eq $null){
						$itemTitle = ""
					}
					else {
						$itemTitle = $Item.Name.Replace("|","\")
					}
					$Items = $List.Items | ?{$_.HasUniqueRoleAssignments -eq $true}
					ForEach ($Item in $Items)
					{
						If ($Item -ne $null)
						{
							# 8.	Get AD Groups for Item
							Write-Host ("                  Retrieving Item AD Groups: {0}" -f $Item.Name) -ForegroundColor DarkRed
							Get-CatSP10ADGroups -File $ADGroupFile -Assignments $Web.RoleAssignments -SiteName $Site.RootWeb.Title -SiteUrl $Site.Url -WebName $Web.Title -WebUrl $Web.Url -ListName $List.Title -LURL $ListURL -ItemTitle $itemTitle
						}
					}
				}			
			} # ForEach ($List in $Web.Lists)
				
			# Screen Output for Web
			Write-Host "            Web scan complete" -ForegroundColor Green
			$Web.Dispose()
		} # ForEach ($Web in $Webs)

		# Dispose of site object
		Write-Host "         Site scan complete" -ForegroundColor Green
		$Site.Dispose()
	} # ForEach ($Site in $Sites)
	# Screen Output for Web App
	Write-Host "      Web Application Scan Complete" -ForegroundColor Green

} # ForEach ($WebApp in $WebApps)
# Screen Output for Farm
Write-Host "   Farm Scan Complete" -ForegroundColor Green
#endregion

#region Completed Screen Output
Complete-CatScript $Start[0]
Stop-Transcript
#endregion