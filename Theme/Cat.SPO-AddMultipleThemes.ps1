<# Cat.SPO-AddTheme.ps1

This script adds a theme to SharePoint Online based on a JSON imput file

Dependencies
	-	SharePoint Online Module
	-	Input JSON File in Input Files folder - Script Parameter - "Cat.SPO-ThemeTest.json"

Call Script with Input File
	.\Cat.SPO-AddTheme.ps1 -adminURL "https://nrpdev-admin.sharepoint.com" -themeFile "Cat.SPO-ThemeCat.ps1"

Theme PowerShell Example
$themePalette = @{
    "themePrimary" = "#c22613";
    "themeLighterAlt" = "#fdf5f3";
    "themeLighter" = "#f5d6d2";
    "themeLight" = "#edb4ad";
    "themeTertiary" = "#da7264";
    "themeSecondary" = "#c93c2a";
    "themeDarkAlt" = "#ae2411";
    "themeDark" = "#931e0f";
    "themeDarker" = "#6d160b";
    "neutralLighterAlt" = "#e2e2e2";
    "neutralLighter" = "#dedede";
    "neutralLight" = "#d5d5d5";
    "neutralQuaternaryAlt" = "#c6c6c6";
    "neutralQuaternary" = "#bdbdbd";
    "neutralTertiaryAlt" = "#b6b6b6";
    "neutralTertiary" = "#bdc3c6";
    "neutralSecondary" = "#80888c";
    "neutralPrimaryAlt" = "#4c5357";
    "neutralPrimary" = "#373d40";
    "neutralDark" = "#2a2e30";
    "black" = "#1f2224";
    "white" = "#e6e6e7"
  }

Table of Contents
	1.	Input & Output Files
	2.	Variables
	3.	Snapin(s), Modules & Environment Connection(s)
	4.	Main Script
#>

#region Script Parameters
Param(  [String] $adminURL = "https://nrpdev-admin.sharepoint.com",
        [String] $themesFile = "Cat.Themes.csv"
)
#endregion

#region Start Date / Time
# Remove Existing Catapult Modules to ensure PS session has most up to date functions
Get-Module Cat.* | Remove-Module

# Load General Functions
Import-Module ("..\General\Cat.GeneralFunctions.psm1")

# Get Start Date / Time
$scriptStart = Start-CatScript
#endregion

#region 1.	Input & Output Files
# File Paths
$inPath = ("..\_Input Files\Themes\{0}" -f $themesFile)

# Loading Files
$inFile = Import-Csv $inPath -Encoding Unicode
#endregion

#region 2.	Variables
# User Inputs
Write-Host
Write-Host "User Input Variables" -ForegroundColor Cyan
Write-Host "   Enter credentials for SharePoint:"
$spCred = Get-Credential -Message "Enter your credentials for SharePoint:"

#Output Parameters & Variables
Write-Host
Write-Host "Script Parameters" -ForegroundColor Cyan
Write-Host ("   Themes File: {0}" -f $themesFile)
Write-Host
Write-Host "Script Variables" -ForegroundColor Cyan
Write-Host ("   Credentials: {0}" -f $spCred.UserName)
#endregion

#region 3.	Snapin(s), Modules & Environment Connection(s)
# Load PowerShell Snap-In(s)
Import-Module Microsoft.Online.SharePoint.PowerShell -DisableNameChecking -erroraction SilentlyContinue
#endregion
    
#region 4.	Main Script
# Connecting To SP Online
Write-Host
Write-Host ("Connecting to SharePoint Online: {0}" -f $adminURL) -ForegroundColor Magenta
Connect-SPOService -Url $adminURL -Credential $spCred

# Loop through Themes CSV file
foreach ($theme in $inFile){
    # Loading Theme Palette
    $themePath = ("..\_Input Files\Themes\{0}" -f $theme.themeFile)
    Import-Module ($themePath)

    # Adding JSON Theme
    Write-Host ("   Adding Theme: {0}" -f $theme.themeName)
    Add-SPOTheme -Name $theme.themeName -Palette $themePalette -IsInverted $false #-Overwrite # User overwrite to update existing
}

# Disconnect from SPO
Disconnect-SPOService
#endregion

#region Completed Screen Output
Complete-CatScript $scriptStart[0]
#endregion