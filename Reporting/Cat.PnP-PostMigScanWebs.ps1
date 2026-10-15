<#
Cat.Pnp-PostMigrationScanWeb.ps1

This script loops through a CSV file to Compare the following attributes post migration.  It collects all list
data first and then loops through at an item level.

Dependencies
	-	PnP PowerShell https://msdn.microsoft.com/en-us/pnp_powershell/pnp-powershell-overview
	-	BCT.PnP-ReportingFunctions.psm1
	-	Input CSV File in Input Files folder - Script Parameter - "Cat.ScanSites.Test.csv"

Call Script with Input File
	.\Cat.PnP-PostMigScanWebs.ps1 -InputFile Cat.PostMigWebs.Test.csv

Enhancements
	- Update to handle large lists

CSV Example
	Site,SP13Url,SPOUrl
	XYZ Site,https://sharepoint.client.com/sites/site/web,https://client.sharepoint.com/sites/site/web

Table of Contents
	1.	Input & Output Files
	2.	Variables
	3.	Snapin(s), Modules & Environment Connection(s)
	4.	Main Script
#>

#region Script Parameters
Param([string]$InputFile)
#endregion

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

#region 1.	Input & Output Files
# File Paths
$InPath = ("{0}_Input Files\{1}" -f $Location, $InputFile)

# Import Input File(s)
$InFile = import-csv $InPath -Encoding Unicode
#endregion

#region 2.	Variables
# User Inputs
Write-Host
Write-Host "User Input Variables" -ForegroundColor DarkCyan
Write-Host "   Enter credentials for SharePoint 2013 farm:"
$SP13Cred = Get-Credential -Message "Enter your credentials for SharePoint 2013 farm:" 
Write-Host "   Enter credentials for SharePoint Online:"
$SPOCred = Get-Credential -Message "Enter your credentials for SharePoint Online:" 
$ReportUrl = Read-Host -Prompt "   Enter the URL for the Reporting site"
Write-Host "   Enter credentials for Reporting Site:"
$ReportCred = Get-Credential -Message "Enter your credentials for Reporting Site:" 

# Script Variables
$WebTotal = (Measure-Object -InputObject $InFile -Line).Lines + 1
#endregion

#region 3.	Snapin(s), Modules & Environment Connection(s)
# Load Custom Modules
Import-Module ("{0}Reporting\Cat.ReportingFunctions.psm1" -f $Location)
#endregion
    
#region 4.	Main Script
# Screen Output
Write-Host
Write-Host ("Scanning sites in CSV file: {0}" -f $InputFile ) -ForegroundColor DarkCyan

#region Looping rows in CSV to scan SP 13 webs & lists
Write-Host
Write-Host "   Scanning SharePoint 2013 Webs and Lists" -ForegroundColor Magenta
$WebCount = 0
ForEach ($Web in $InFile)
{
	$WebCount++
	Get-CatPostMigWebListInfo -Env "SP13" -ScanURL $Web.SP13Url -ScanCreds $SP13Cred -ReportURL $ReportUrl -ReportCred $ReportCred -WCount $WebCount -WTotal $WebTotal
} # ForEach ($Web in $InFile)

# Screen Output for Web
Write-Host "      SharePoint 2013 Webs and Lists Scan Complete" -ForegroundColor Green
#endregion SP 13 webs & lists

#region Looping rows in CSV to scan SPO webs & lists
Write-Host
Write-Host "   Scanning SharePoint Online Webs and Lists" -ForegroundColor Magenta
$WebCount = 0
ForEach ($Web in $InFile)
{
	$WebCount++
	Get-CatPostMigWebListInfo -Env "SPO" -ScanURL $Web.SPOUrl -ScanCreds $SPOCred -ReportURL $ReportUrl -ReportCred $ReportCred -WCount $WebCount -WTotal $WebTotal
} # ForEach ($Web in $InFile)

# Screen Output for Web
Write-Host "      SharePoint Online Webs and Lists Scan Complete" -ForegroundColor Green
#endregion SPO webs & lists #>

#region Looping rows in CSV to scan SP13 Items
Write-Host
Write-Host "   Scanning SharePoint 2013 Items" -ForegroundColor Magenta
$WebCount = 0
ForEach ($Web in $InFile)
{
	$WebCount++
	Get-CatPostMigItemInfo -Env "SP13" -ScanURL $Web.SP13Url -ScanCreds $SP13Cred -ReportURL $ReportUrl -ReportCred $ReportCred -WCount $WebCount -WTotal $WebTotal
} # ForEach ($Web in $InFile)

# Screen Output for Web
Write-Host "      SharePoint 2013 Items Scan Complete" -ForegroundColor Green
#endregion SP13 Items

#region Looping rows in CSV to scan SPO Items
Write-Host
Write-Host "   Scanning SharePoint Online Items" -ForegroundColor Magenta
$WebCount = 0
ForEach ($Web in $InFile)
{
	$WebCount++
	Get-CatPostMigItemInfo -Env "SPO" -ScanURL $Web.SPOUrl -ScanCreds $SPOCred -ReportURL $ReportUrl -ReportCred $ReportCred -WCount $WebCount -WTotal $WebTotal
} # ForEach ($Web in $InFile)

# Screen Output for Web
Write-Host "      SharePoint Online Items Scan Complete" -ForegroundColor Green
#endregion SP13 Items #>

# Screen Output 
Write-Host ("   Scanning sites in CSV file Complete: {0}" -f $InputFile ) -ForegroundColor Green
#endregion Main Script

#region Completed Screen Output
Complete-CatScript $Start[0]
#endregion