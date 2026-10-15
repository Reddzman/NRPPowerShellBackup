<#
Cat.SP10-WarmUpServer.ps1

This script makes web requests to site in order to for just in time compilation.

Dependencies
	-	SharePoint 2010 Module
	-	Input CSV File in Input Files folder - Script Parameter - "Cat.PerformAction.Test.csv"

Call Script with Input File
	.\Cat.Tech-ScriptName.ps1 -InputFile InputFileName.csv

CSV Example
	Site,SP13Url,SPOUrl
	SiteName,https://sharepoint.client.com/sitedirectory/law,https://client.sharepoint.com/sites/law

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
$InPath = ("{0}_Input Files\Input.csv" -f $Location)
$OutPath = ("{0}_Output Files\Output.txt" -f $Location)
$ErrorPath = ("{0}_Error Files\Error.csv" -f $Location)

# Create Output File(s)
Out-File -FilePath $OutPath -InputObject "Database|SiteTitle|SiteURL|Listname|ItemCount"
Out-File -FilePath $ErrorPath -InputObject "SiteName,SourceWeb,DestWeb"

# Import Input File(s)
$InFile = import-csv $InPath -Encoding Unicode
[xml]$InFile = get-content $InPath
#endregion

#region 2.	Variables
# User Inputs
Write-Host
Write-Host "User Input Variables" -ForegroundColor DarkCyan
Write-Host "   Enter credentials for SharePoint 2013 farm:"
$SP13Cred = Get-Credential -Message "Enter your credentials for SharePoint 2013 farm:" 
$SP13Url = Read-Host -Prompt "   Enter the SharePoint 2013 site URL"

# Script Variables
$SiteColectionURL = "https://sharepoint.client.com"
$MaxItemCount = 5000
#endregion

#region 3.	Snapin(s), Modules & Environment Connection(s)
# Load PowerShell Snap-In(s)
Import-Module Microsoft.Online.SharePoint.PowerShell -erroraction SilentlyContinue
Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue

# Load Custom Modules
Import-Module ("{0}FolderName\ModuleName.psm1" -f $Location)
#endregion
    
#region 4.	Main Script
# Connecting To SP Online
$Cred = Get-Credential -Message "Enter your credentials for SharePoint Online:" 
Connect-SPOService -Url $SPOAdminSite -Credential $Cred



#endregion

#region Completed Screen Output
Complete-CatScript $Start[0]
#endregion