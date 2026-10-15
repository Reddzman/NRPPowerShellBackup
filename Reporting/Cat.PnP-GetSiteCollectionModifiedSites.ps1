<#
Cat.Tech-ScriptName.ps1

This script...

Dependencies
	-	PnP PowerShell https://msdn.microsoft.com/en-us/pnp_powershell/pnp-powershell-overview

Table of Contents
	1.	Input & Output Files
	2.	Variables
	3.	Snapin(s), Modules & Environment Connection(s)
	4.	Main Script
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

#region 2.	Variables
# User Inputs
Write-Host
Write-Host "User Input Variables" -ForegroundColor DarkCyan
Write-Host "   Enter credentials for SharePoint site:"
$SPCred = Get-Credential -Message "Enter your credentials for SharePoint site"
$SPURL = Read-Host -Prompt "   Enter URL for SharePoint Online reporting site" 

# Script Variables
$Days = 90
$Date = (Get-Date).AddDays(-($Days))
#endregion

#region 4.	Main Script
# Connect to Site
Connect-PnPOnline -Url $SPURL -Credentials $SPCred

# Get Site object and all web objects in site collection
$Site = Get-PnPSite -Includes Usage
$Webs = @()
$Webs = Get-PnPSubWebs -Recurse
$Webs += Get-PnPWeb

# Output Count for all webs
Write-Host
Write-host ("Total number of sites: {0}" -f $Webs.count) -foregroundcolor Magenta

# Output storage for site collection
Write-host ("Total amount of storage: {0}" -f ($Site.Usage.Storage/1024/1024/1024)) -foregroundcolor Magenta

# Get webs modified in last X days
$Webs = Get-PnPSubWebs -Recurse | ?{$_.LastItemModifiedDate -gt $Date}
$Webs += Get-PnPWeb | ?{$_.LastItemModifiedDate -gt $Date}

# Output count for modified webs
Write-host ("Total number of modified sites: {0}" -f $Webs.count ) -foregroundcolor Magenta
#endregion

#region Completed Screen Output
Complete-CatScript $Start[0]
#endregion