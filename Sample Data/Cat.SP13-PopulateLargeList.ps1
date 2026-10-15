<#
Cat.SP10-PopulateLargeList.ps1

This script script will at items to a list to surpass a particular threshold.

Dependencies
	-	SharePoint 2013 Module

Call Script with Input File
	.\Cat.SP10-PopulateLargeList.ps1

Table of Contents
	1.	Input & Output Files
	2.	Variables
	3.	Snapin(s), Modules & Environment Connection(s)
	4.	Main Script
#>

#region Start Date / Time
# Set PS location
$PSScriptRoot = Split-Path $MyInvocation.MyCommand.Path -Parent
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
$SPUrl = Read-Host -Prompt "   Enter the SharePoint 2013 web URL"
$SPList = Read-Host -Prompt "   Enter the SharePoint 2013 list title"
$ItemCount = Read-Host -Prompt "   Enter the number of items to create"
#endregion

#region 3.	Snapin(s), Modules & Environment Connection(s)
# Load PowerShell Snap-In(s)
Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue
#endregion
    
#region 4.	Main Script
# Screen Output
Write-Host ("Adding Items to List: {0}" -f $SPList) -ForegroundColor DarkCyan

# Get SharePoint site
$Web = Get-SPWeb $SPUrl
Write-Host ("   Retrieving to Web: {0}" -f $Web.Title)

# Get List
$List = $Web.Lists[$SPList]
Write-Host ("   Retrieving to List: {0}" -f $List.Title)

# Create Items in list
Write-Host ("   Creating {0} Items" -f $ItemCount) -NoNewLine
For ($I = 1; $I -le $ItemCount; $I++)
{
	# Create new Item
	$NewItem = $List.AddItem()

	# Update properties
	$NewItem["Title"] = ("Item # {0}" -f $I)

	# Update item
	$NewItem.Update()

	# Screen Output
	Write-Host "." -NoNewLine
} # For ($I = 1; $I -le $ItemCount; $I++)

# Screen Output
Write-Host
Write-Host "Item Creation Complete" -ForegroundColor Green
#endregion

#region Completed Screen Output
Complete-CatScript $Start[0]
#endregion