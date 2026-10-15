<#  Cat.ADO-TestPipelineToM365.ps1

This script is designed to test script connections to SharePoint on premises an ADO Repo and Pipeline to execute the script. This must be executed on a SharePoint Server.
Powershell needs to be executed by a Farm Administrator and Site Collection administrator of the URL provided

Dependencies
    -   SharePoint Server Farm
    -   SharePoint Server PowerShell Modules

Call script with parameters
	.\Cat.ADO-TestPipelineToM365.ps1 -client "nrpdev" -url "https://portal.sp13test.local"

Table of Contents
	1.	Variables
	2.	Snapin(s), Modules & Environment Connection(s)
    3.	Main Script
#>

#region Script Parameters
Param(
    [string] $client = "nprdev",
    [string] $url = "https://portal.sp13test.local"
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
Start-Transcript -Path ("..\_Log Files\{0}.ADO-TestPipelineToSharePoint-{1}.log" -f $client, $scriptStart[1])
#endregion

#region 1.	Variables
# Variable Output
Write-Host
Write-Host "Script Parameters" -ForegroundColor Cyan
Write-Host ("   Client: {0}" -f $client)
Write-Host ("   URL: {0}" -f $url)
#endregion

#region 2.	Snapin(s), Modules & Environment Connection(s)
# Load MS / PnP PowerShell Modules
Get-Module -ListAvailable | Where-Object{($_.Name -like "*SharePoint*") -or ($_.Name -like "*PnP*") } | Remove-Module
Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue
#endregion
    
#region 3.	Main Script
# Connecting To SP Online
Write-Host
Write-Host "Testing Connection to SharePoint:" -ForegroundColor Cyan

# Retrieving Farm Information
Write-Host "   Retrieving Farm Information" -NoNewline
$farm = Get-SPFarm
Write-Host (" - Farm Config DB: {0}" -f $farm.Name)

# Testing Connection to SharePoint Online Root Site Collection
Write-Host ("   Retrieving Site: {0}" -f $url) -NoNewline
$site = Get-SPSite $url
Write-Host (" - Number of Webs: {0}" -f $site.AllWebs.Count)
#endregion

#region Completed Script
# Output Script Info
Complete-CatScript $scriptStart[0]

# Stop logging
Stop-Transcript
#endregion