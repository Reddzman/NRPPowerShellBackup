<#  Cat.SPO-RemoveAllThemes.ps1

This script removes all custom themes from a SharePoint Online Tenant.

Dependencies
	-	SharePoint Online Module

Table of Contents
	1.	Variables
	2.	Snapin(s), Modules & Environment Connection(s)
	3.	Main Script
#>

#region Script Parameters
Param(
    [String] $adminURL = "https://spnickp-admin.sharepoint.com"
)
#endregion

#region Start Date / Time
# Remove Existing Catapult Modules to ensure PS session has most up to date functions
Get-Module Cat.* | Remove-Module

# Load General Functions
Import-Module "..\General\Cat.GeneralFunctions.psm1"

# Get Start Date / Time
$scriptStart = Start-CatScript
#endregion

#region 1.	Variables
# User Inputs
Write-Host
Write-Host "User Input Variables" -ForegroundColor Cyan
Write-Host "   Enter credentials for SharePoint:"
$spCred = Get-Credential -Message "Enter your credentials for SharePoint:"
#endregion

#region Output Parameters & Variables
Write-Host
Write-Host "Script Parameters" -ForegroundColor Cyan
Write-Host ("   Admin URL: {0}" -f $adminURL)
Write-Host
Write-Host "Script Variables" -ForegroundColor Cyan
Write-Host ("   Credentials: {0}" -f $spCred.UserName)
#endregion

#region 2.	Snapin(s), Modules & Environment Connection(s)
# Load PowerShell Snap-In(s)
Import-Module Microsoft.Online.SharePoint.PowerShell -DisableNameChecking -erroraction SilentlyContinue
#endregion
    
#region 3.	Main Script
# Connecting To SP Online
Write-Host
Write-Host ("Connecting to SharePoint Online: {0}" -f $adminURL) -ForegroundColor Magenta 
Connect-SPOService -Url $adminURL -Credential $spCred

# Get all themes
Write-Host "   Getting all custom themes"
$allThemes = Get-SPOTheme

# Deleting Custom Themes
foreach ($theme in $allThemes){
    Write-Host ("   Deleting Theme: {0}" -f $theme.Name)
    Remove-SPOTheme -name $theme.Name
}

# Disconnect from SPO
Disconnect-SPOService
#endregion

#region Completed Screen Output
Complete-CatScript $scriptStart[0]
#endregion