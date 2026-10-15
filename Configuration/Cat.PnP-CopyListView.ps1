<#
Cat.PnP-CopyListView.ps1

This script copies a sharepoint list view from one location to another.

Dependencies
	-	PnP PowerShell https://msdn.microsoft.com/en-us/pnp_powershell/pnp-powershell-overview

Call Script with Input File
	.\Cat.PnP-CopyListView.ps1

Table of Contents
	1.	Variables
	2.	Main Script
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
# User Inputs
Write-Host
Write-Host "User Input Variables" -ForegroundColor DarkCyan
Write-Host "   Enter credentials for SharePoint:"
$SPCred = Get-Credential -Message "Enter your credentials for SharePoint:" 
$SourceWeb = Read-Host -Prompt "   Enter the SharePoint source site URL"
$SourceList = Read-Host -Prompt "   Enter the SharePoint source list name"
$SourceView = Read-Host -Prompt "   Enter the SharePoint source view name"
$DestWeb = Read-Host -Prompt "   Enter the SharePoint destination site URL"
$DestList = Read-Host -Prompt "   Enter the SharePoint destination list name"
#endregion
  
#region 2.	Main Script
# Connecting To SP source site
Write-Host
Write-Host ("Connecting to source site: {0}" -f $SourceWeb) -ForegroundColor Magenta
Connect-PnPOnline -Url $SourceWeb -Credentials $SPCred

# Get Source View Configuration
Write-Host ("   Retrieving source view configuration: {0}" -f $SourceView)
$View = Get-PnPView -List $SourceList -Identity $SourceView -Includes ViewQuery, ViewType, RowLimit, Paged

# Concatenating Fields
Write-Host ("   Concatenating fields")
$ViewFields = @()
foreach ($Field in $View.ViewFields)
{
	$ViewFields += $Field
}

# output List Settings
Write-Host ("   View Settings")
Write-Host ("      Fields: {0}" -f $ViewFields)
Write-Host ("      Query: {0}" -f $View.ViewQuery)
Write-Host ("      Type: {0}" -f $View.ViewType)
Write-Host ("      Row Limit: {0}" -f $View.RowLimit)
Write-Host ("      Personal View: {0}" -f $View.PersonalView)
Write-Host ("      Default View: {0}" -f $View.DefaultView)
Write-Host ("      Paging Enabled: {0}" -f $View.Paged)

# Connect to SP destination site
Write-Host
Write-Host ("Connecting to destination site: {0}" -f $DestWeb) -ForegroundColor Magenta
Connect-PnPOnline -Url $DestWeb -Credentials $SPCred

# Create View in Destination
Write-Host ("   Creating view in destination list: {0}/{1}" -f $DestWeb, $DestList)
$Temp = Add-PnPView -List $DestList -Title $View.Title -Fields $ViewFields -Query $View.ViewQuery -ViewType $View.ViewType -RowLimit $View.RowLimit -Personal:$View.PersonalView -SetAsDefault:$View.DefaultView -Paged:$View.Paged
#endregion

#region Completed Screen Output
Complete-CatScript $Start[0]
#endregion