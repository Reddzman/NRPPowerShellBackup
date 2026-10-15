<# Cat.SP10_PreMigScanInfoArch.ps1

This script generates reports for an entire Farm to identify the following information:
	-	Content Type
	-	Columns

Potential Improvements
	-	Output Web Application information as part of CSV.

Table of Contents
	1.	Input Files
	2.	Variables
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

# Start Transcript
Start-Transcript -Path ("..\_Log Files\InfoArchScan-{0}.log" -f $Start[1])
#endregion

#region 1.	Input & Output Files
# File Paths
$CTFile = ("{0}_Output Files\Client-ContentTypes-{1}.txt" -f $Location, $Start[1])
$ColumnFile = ("{0}_Output Files\Client-Columns-{1}.txt" -f $Location, $Start[1])

# Create Output File(s)
Out-File -FilePath $CTFile -InputObject "Site|SiteURL|Web|WebURL|List|ListURL|ContentType|Parent|Scope"
Out-File -FilePath $ColumnFile -InputObject "Site|SiteURL|Web|WebURL|List|ListURL|ContentType|Column|Type|Scope"
#endregion

#region 2.	Variables
# Script Variables
$SkipLists = @("Content and Structure Reports";"Customized Reports";"Form Templates";"Master Page Gallery";"Reusable Content";"Style Library";"Workflow Tasks";)
$SkipColumns = @("Item Child Count";"Folder Child Count";"App Created By";"App Modified By";"Check In Comment";"Source Version (Converted Document)";
			"Source Name (Converted Document)";"Check In Comment";"DisplayVersion";"Content Type")
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
	if ($Site.RootWeb.Title -eq $null){
		$siteTitle = ""
	}
	else {
		$siteTitle = $Site.RootWeb.Title.Replace("|","\")
	}

	# Loop through site collections for web application
	ForEach ($Site in $Sites)
	{
		# Screen Output for Site
		Write-Host ("      Scanning site collection: {0}" -f $Site.URL) -ForegroundColor DarkGreen
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

			# Loop through all lists and libraries in a sub site
			$Lists = $Web.Lists | Where{($_.Hidden -eq $false) -and ($SkipLists -notcontains $_.Title)}
			ForEach ($List in $Lists)
			{
				if ($List -ne $null)
				{
					# Screen output for web
					Write-Host ("            Scanning List: {0}" -f $List.Title) -ForegroundColor DarkGray
					if ($List.Title -eq $null){
						$listTitle = ""
					}
					else {
						$listTitle = $List.Title.Replace("|","\")
					}
					$ListURL = ("{0}/{1}" -f $Site.Url, $List.RootFolder.Url)

					# Get Content Types
					$ContentTypes = $List.ContentTypes

					#Loop Through all ContentTypes
					ForEach ($ContentType in $ContentTypes)
					{
						# Output Content Type Information
						Write-Host ("               Creating Content Type Record: {0}" -f $ContentType.Name) -ForegroundColor Yellow
						if ($ContentType.Name -eq $null){
							$ctName = ""
						}
						else {
							$ctName = $ContentType.Name.Replace("|","\")
						}
						if ($ContentType.Parent.Name -eq $null){
							$parentName = ""
						}
						else {
							$parentName = $ContentType.Parent.Name.Replace("|","\")
						}
						Out-File -FilePath $CTFile -Append -InputObject ("{0}|{1}|{2}|{3}|{4}|{5}|{6}|{7}|{8}" -f 
							$siteTitle, 
							$Site.Url, 
							$webTitle, 
							$Web.Url, 
							$listTitle, 
							$ListURL, 
							$ctName, 
							$parentName, 
							$ContentType.Scope)

						# Get Columns
						$Columns = $ContentType.Fields | Where{($_.Hidden -eq $false) -and ($SkipColumns -notcontains $_.Title )}
						
						# Loop Through all Columns
						ForEach ($Column in $Columns)
						{
							# Output Column Information
							Write-Host ("                  Creating Column Record: {0}" -f $Column.Title) -ForegroundColor Gray
							if ($Column.Title -eq $null){
								$columnName = ""
							}
							else {
								$columnName = $Column.Title.Replace("|","\")
							}
							Out-File -FilePath $ColumnFile -Append -InputObject ("{0}|{1}|{2}|{3}|{4}|{5}|{6}|{7}|{8}|{9}" -f 
								$siteTitle, 
								$Site.Url, 
								$webTitle, 
								$Web.Url, 
								$listTitle, 
								$ListURL, 
								$ctName, 
								$columnName, 
								$Column.TypeAsString, 
								$Column.Scope)
						} # ForEach ($Column in $Columns)
					} # ForEach ($ContentType in $ContentTypes)
				} # if ($List -ne $null)
			} # ForEach ($List in $Web.Lists)
				
			# Screen Output for Web
			Write-Host "         Web scan complete"
			$Web.Dispose()
		} # ForEach ($Web in $Webs)

		# Dispose of site object
		Write-Host "      Site scan complete" -ForegroundColor DarkGreen
		$Site.Dispose()
	} # ForEach ($Site in $Sites)
	# Screen Output for Web App
	Write-Host "   Web Application Scan Complete" -ForegroundColor Magenta

} # ForEach ($WebApp in $WebApps)
# Screen Output for Farm
Write-Host "Farm Scan Complete" -ForegroundColor Green
#endregion

#region Completed Screen Output
Complete-CatScript $Start[0]

Stop-Transcript
#endregion