#region Documentation
<#  Cat.SPO-GetSiteScriptFromWeb.ps1

This script this script extracts the initial Site Script (JSON) from a web that can be leveraged with a Site Design to create a SharePoint Template.

Dependencies
	-	SPO PowerShell
	-	General Functions - ..\General\Cat.GeneralFunctions.psm1

Call script with parameters
	.\Cat.SPO-GetSiteScriptFromWeb.ps1 -siteURL "https://nrpdev2.sharepoint.com/teams/SiteTemplateITPMOSCRUM" -siteScriptName "IT PMO SCRUM Template" -clientTenant "nrpdev2" -includeLists @("IT Project Documents", "SiteAssets", "Lists/Risk Log_New", "Lists/Test Cases  Results", "Lists/Announcements", "Lists/Communications Calendar", "Lists/Contacts", "Lists/Final Project Sign Off", "Lists/IT PMO Links", "Lists/Main goals and objectives", "Lists/Events", "Lists/IPM Scope Assessment Checklist", "Lists/Standard%20Links") -includeBranding $false -includeTheme $false -includeRegionalSettings $false -includeExternalSharingCapability $false -includeLinks $true

Table of Contents
	1.	Snapin(s), Modules & Environment Connection(s)
	2.	Input & Output Files
	3.	Variables
	4.	Main Script
#>
#endregion

#region Script Parameters
Param(	[string] $siteURL = "https://client.sharepoint.com/teams/SiteTemplate",
        [string] $siteScriptName = "Site Template",
        [string] $clientTenant = "{ClientTenant}",
        [array] $includeLists = @("IT Project Documents", "Lists/Risk Log_New", "Lists/Test Cases  Results", "Lists/Announcements", "Lists/Communications Calendar",
                "Lists/Contacts", "Lists/Final Project Sign Off", "Lists/IT PMO Links", "Lists/Main goals and objectives", "Lists/Events", "Lists/IPM Scope Assessment Checklist", "Lists/Standard%20Links"),
        [boolean] $includeBranding = $false,
        [boolean] $includeTheme = $false,
        [boolean] $includeRegionalSettings = $false,
        [boolean] $includeExternalSharingCapability = $false,
        [boolean] $includeLinks = $true
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
Start-Transcript -Path ("..\_Log Files\{0}.SPO-GetSiteScriptFromWeb-{1}.log" -f $clientTenant, $scriptStart[1])
#endregion

#region 1.	Snapin(s), Modules & Environment Connection(s)
# Load MS / PnP PowerShell Modules
Get-Module -ListAvailable | Where-Object{($_.Name -like "*SharePoint*") -or ($_.Name -like "*PnP*") } | Remove-Module
Import-Module Microsoft.Online.SharePoint.PowerShell -DisableNameChecking -erroraction SilentlyContinue
#endregion

#region 2.	Input & Output Files
# File Paths
$outPath = ("..\_Output Files\{0}-{1}-{2}.json" -f $clientTenant, $siteScriptName, $scriptStart[1])
#endregion

#region 3.	Variables
# Script Variables
$adminURL = ("https://{0}-admin.sharepoint.com" -f $clientTenant)

# Output Parameters & Variables
Write-Host
Write-Host "Script Parameters" -ForegroundColor Cyan
Write-Host ("   Site URL: {0}" -f $siteURL)
Write-Host ("   Site Script Name: {0}" -f $siteScriptName)
Write-Host ("   Client Tenant: {0}" -f $clientTenant)
Write-Host ("   Include Branding: {0}" -f $includeBranding)
Write-Host ("   Include Theme: {0}" -f $includeTheme)
Write-Host ("   Include Regional Settings: {0}" -f $includeRegionalSettings)
Write-Host ("   Include External Sharing Capability: {0}" -f $includeExternalSharingCapability)
Write-Host ("   Include Links: {0}" -f $includeLinks)
Write-Host
Write-Host "Script Variables" -ForegroundColor Cyan
Write-Host ("   Output File: {0}" -f $outPath)
Write-Host ("   Admin URL: {0}" -f $adminURL)
Write-Host
Write-Host "Included Lists" -ForegroundColor Cyan
$includeLists
#endregion
    
#region 4.	Main Script
# Connecting To SP Online
Write-Host
Write-Host "Creating Site Script JSON" -ForegroundColor Cyan
Write-Host ("   Connecting to SharePoint Online site: {0}" -f $adminURL)
Connect-SPOService -Url $adminURL
  
# Get the site schema
Write-Host ("   Extracting Site Script from Web: {0}" -f $siteURL)
$siteScript = Get-SPOSiteScriptFromWeb -WebURL $siteURL -IncludedLists $includeLists -IncludeBranding:$includeBranding -IncludeTheme:$includeTheme -IncludeRegionalSettings:$includeRegionalSettings -IncludeSiteExternalSharingCapability:$includeExternalSharingCapability -IncludeLinksToExportedItems:$includeLinks

# export to JSON
Write-Host ("   Exporting to JSON: {0}" -f $outPath)
Out-File -FilePath $outPath -InputObject $siteScript

# Disconnect from Services
Write-Host ("   Disconnecting from SharePoint Online: {0}" -f $adminURL)
Disconnect-SPOService
#endregion

#region Completed Script
# Output Script Info
Complete-CatScript $scriptStart[0]

# Stop logging
Stop-Transcript
#endregion