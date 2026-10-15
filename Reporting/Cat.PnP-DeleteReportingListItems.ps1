<#
Cat.PnP-DeleteReportingListItems.ps1

This script deletes all items from the lists listed in the $ListNames variable.  It assumes the migration site is at the root
of the URL provided and that all lists exist. 

Dependencies
	-	PnP PowerShell https://msdn.microsoft.com/en-us/pnp_powershell/pnp-powershell-overview

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
Write-Host "   Enter credentials for SharePoint Online reporting site:"
$ReportCred = Get-Credential -Message "Enter your credentials for SharePoint Online reporting site:"
$ReportURL = Read-Host -Prompt "   Enter URL for SharePoint Online reporting site"

# Script Variables
<# All Migration Lists
$ListNames = @("Migration Active Directory Groups";"Migration Checked Out Documents";"Migration Complex Columns";"Migration InfoPath";"Migration Large Lists";
    "Migration Site Collections";"Migration Solutions";"Migration Sub Sites";"Migration Workflows";"Migration Item Version Counts";"Migration List Item Counts";
    "Migration Permission Counts";) #>

# Post Migration Report
$ListNames = @("Migration Item Version Counts 3";"Migration List Item Counts 3";"Migration Permission Counts 3";) #>

#endregion
    
#region 2.	Main Script
# Screen Output
Write-Host
Write-Host "Deleting Reporting List Items" -ForegroundColor DarkCyan

# Connect to reporting site
Write-Host ("   Connecting to reporting site: {0}" -f $ReportURL)
Connect-PnPOnline -Url $ReportURL -Credentials $ReportCred

# Loop through list Names
ForEach ($ListName in $ListNames)
{
    # Screen Output for List
    Write-Host ("   Deleting items for: {0}" -f $ListName)
    
    # Loop through list items
    $List = Get-PnPList -Identity $ListName
    $Items = Get-PnPListItem -List $List
    ForEach ($Item in $Items)
    {
        # Removing List Item
        Remove-PnPListItem -List $ListName -Identity $Item.Id -Force
    } # ForEach ($Item in $Items)
    Write-Host ("      {0} Complete" -f $ListName) -ForegroundColor Green
} # ForEach ($List in $Lists)
#endregion

#region Completed Screen Output
Complete-CatScript $Start[0]
#endregion