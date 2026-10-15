<#	Cat.PnP-CreateMigrationSite.ps1

This script will create a site in SharePoint Online for the purpose of communicating migration data from
the Scan Environment scripts.

Dependencies
	-	PnP PowerShell https://msdn.microsoft.com/en-us/pnp_powershell/pnp-powershell-overview

Table of Contents
	1.	Input & Output Files
	2.	Variables
	3.	Snapin(s), Modules & Environment Connection(s)
	4.	Main Script

**** Last Updated 5/10/18 ****
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
# Input Files
$SiteTemplate = ("{0}_Input Files\MigrationSiteTemplate.xml" -f $Location)
#endregion

#region 2.	Variables
# User Inputs
Write-Host
Write-Host "User Input Variables" -ForegroundColor DarkCyan
Write-Host "   Enter credentials for SharePoint Online:"
$SPOCred = Get-Credential -Message "Enter your credentials for SharePoint Online:"
$SPOUrl = Read-Host -Prompt "   Enter URL for SharePoint Online site you want to create" 
$SPOAdminUrl = Read-Host -Prompt "   Enter URL for SharePoint Online Admin url"
$SiteName = Read-Host -Prompt "   Enter name for site you would like to create"
$SiteOwner = Read-Host -Prompt "   Enter UPN for the site owner"

# Script Variables
[Int]$TimeZone = 12  # 12 = Mountain Standard Time
#endregion

#region 3.	Snapin(s), Modules & Environment Connection(s)
# Load PowerShell Snap-In(s)
Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue

# Load Custom Modules
Import-Module ("{0}Configuration\Cat.ConfigurationFunctions.psm1" -f $Location)
#endregion
    
#region 4.	Main Script
Write-Host
Write-Host "Site Collection Creation" -ForegroundColor DarkCyan

# Connect to on SPO Admin Center
Write-Host "   Connecting to SharePoint Online Admin Center"
Connect-PnPOnline -Url $SPOAdminUrl -Credentials $SPOCred

# 4.	Create New Site Collection
Write-Host ("   Creating SharePoint Online site for {0}: {1}" -f $SiteName, $SPOUrl)
New-PnPTenantSite -Title $SiteName -Url $SPOUrl -Owner $SiteOwner -TimeZone $TimeZone -Wait

# Connect to new SPO site collection
Write-Host "      Waiting to contect to new SharePoint Online site collection" -NoNewline -ForegroundColor Yellow
Wait-CatSPOSiteCreation -Creds $SPOCred -URL $SPOUrl

# 5.	Applying Provisioning Template
Write-Host "   Applying Provisioning Template to new SharePoint Online site collection"
Apply-PnPProvisioningTemplate -Path $SiteTemplate

# Screen Output for completion
$SPOSite = Get-PnPSite -Includes RootWeb
Write-Host ("   {0} site collection created sucessfully" -f $SPOSite.RootWeb.Title) -ForegroundColor Green
#endregion

#region Completed Screen Output
Complete-CatScript $Start[0]
#endregion