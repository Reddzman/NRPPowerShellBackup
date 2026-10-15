<#
Cat.SPO-SiteReport.ps1

This script generates a list of all Site Collections and Sites within an SPO tenant.

Dependencies:
	-	SharePoint Online Management Shell
	-	PnP PowerShell https://msdn.microsoft.com/en-us/pnp_powershell/pnp-powershell-overview

Enhancements:
	-		

Table of Contents
	1.	Variables
	3.	Snapin(s), Modules & Environment Connection(s)
	4.	Main Script
	5.	Collect Site Collection Info
	6.	Collect Web Info
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

#region 1.	Input & Output Files
# File Paths
$OutPath = ("{0}_Output Files\Cat.SiteList-{1}.txt" -f $Location, $Start[1])

# Create Output File(s)
Out-File -FilePath $OutPath -InputObject "SiteCollection|SiteCollectionURL|SiteName|SiteURL"
#endregion

#region 1.	Variables
# User Input Variables
Write-Host
Write-Host "User Input Variables" -ForegroundColor DarkCyan
Write-Host "   Enter credentials for the SharePoint Online:"
$SPOCred = Get-Credential -Message "Enter your credentials for SharePoint Online:" 
$SPOURL = Read-Host -Prompt "   Enter URL for SharePoint Online Admin site"
$SCURLFilter = Read-Host -Prompt "   Enter filter for site collection URL (IE. /sites/)"
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
Write-Host "Starting SharePoint Online scan" -ForegroundColor DarkCyan

# Get site collection objects
Write-Host "   Retrieving site collections"
Connect-SPOService $SPOURL -Credential $SPOCred
$Sites = Get-SPOSite -Limit All | Where{$_.Url -like "*$SCURLFilter*"}

# Loop through site collections for web application
ForEach ($Site in $Sites)
{
	# Screen Output for Site
	Write-Host
	Write-Host ("   Scanning site collection: {0}" -f $Site.Title) -ForegroundColor Magenta

	# Get all webs in site collection accept app webs
	Write-Host ("      Retrieving all sub sites" -f $Web.Title)
	$Webs = Get-CatAllWebs -Creds $SPOCred -URL $Site.Url

	# Loop through sub sites
	ForEach ($Web in $Webs)
	{
		# Screen output for web
		Write-Host ("      site: {0}" -f $Web.Title)
		Out-File -FilePath $OutPath -Append -InputObject ("{0}|{1}|{2}|{3}" -f $Site.Title, $Site.Url, $Web.Title, $Web.Url)
	} # ForEach ($Web in $Webs)

	# Site Collection Complete Output
	Write-Host ("      Site collection complete: {0}" -f $Site.Title) -ForegroundColor Green
} # ForEach ($Site in $Sites)
#endregion

#region Completed Screen Output
Complete-CatScript $Start[0]
#endregion