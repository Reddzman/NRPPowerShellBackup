<#  Cat.SPO-ConfigSiteCollectionAdmin.ps1

This script iterates through all site collections in SharePoint Online and add a specified user to the Site Collection Administrators
for that site collection.

Dependencies
	-	SharePoint Online Module

Call script with parameters
	.\Cat.SPO-AddSiteCollectionAdmin.ps1  -spAdminURL "https://{tenant}-admin.sharepoint.com" -client "{File Abrivition}" -adminUPN = "admin@domain.com" -includePersonalSites $false -isSiteCollectionAdmin $true

Table of Contents
	1.	Input & Output Files
	2.	Variables
	3.	Snapin(s), Modules & Environment Connection(s)
	4.	Main Script
#>

#region Script Parameters
Param(	
		[string] $clientTenant = "nrpdev2",
		[string] $adminUPN = "nickp@nrpdev.com",
		[boolean] $includePersonalSites = $false,
		[boolean] $isSiteCollectionAdmin = $true

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
Start-Transcript -Path ("..\_Log FIles\{0}.SPO-AddSiteCollectionAdmin-{1}.log" -f $clientTenant, $scriptStart[1])
#endregion

#region 2.	Variables
# Script Variables
$spAdminURL = ("https://{0}-admin.sharepoint.com" -f $clientTenant)
If($isSiteCollectionAdmin){$adminAction = "Adding"} Else {$adminAction = "Removing"}


# Output Parameters & Variables
Write-Host
Write-Host "Script Parameters" -ForegroundColor Cyan
Write-Host ("   SPO Admin URL: {0}" -f $spAdminURL)
Write-Host ("   Client Abriviation: {0}" -f $clientTenant)
Write-Host ("   Admin User: {0}" -f $adminUPN)
Write-Host ("   Include Personal Sites: {0}" -f $includePersonalSites)
Write-Host ("   Is Site Collection Admin: {0}" -f $isSiteCollectionAdmin)
Write-Host
Write-Host "Script Varialbes" -ForegroundColor Cyan
Write-Host ("   Action: {0}" -f $adminAction)
#endregion

#region 3.	Snapin(s), Modules & Environment Connection(s)
# Load MS / PnP PowerShell Modules
Import-Module Microsoft.Online.SharePoint.PowerShell -DisableNameChecking -erroraction SilentlyContinue

# Load custom modules
Import-Module "..\Reporting\Functions\Cat.SPO-GetAllSites.psm1"
#endregion
    
#region 4.	Main Script
# Connecting To Services
Write-Host
Write-Host "Connecting to Services:" -ForegroundColor Cyan
Write-Host ("   Connecting to SharePoint Online: {0}" -f $spAdminURL)
Connect-SPOService -Url $spAdminURL

# Screen Output
Write-Host
Write-Host ("{0} {1} as a Site Collection Administrator" -f $adminAction, $adminUPN) -ForegroundColor Magenta

# Retrieving SPO Site collections
Write-Host "   Retrieving site collections"
$spSites = Get-CatSPOAllSites -personal $includePersonalSites | Where{$_.Template -ne "REDIRECTSITE#0"}

# Looping through all site collections
ForEach ($spSite in $spSites) {
	Write-Host ("   {0} {1}: {2}" -f $adminAction, $adminUPN, $spSite.Url)
    $temp = Set-SPOUser -Site $spSite.Url -LoginName $adminUPN -IsSiteCollectionAdmin $isSiteCollectionAdmin
}

# Screen Output
Write-Host ("{0} {1} Complete" -f $adminAction, $adminUPN) -ForegroundColor Green
#endregion

#region Completed Script
# Disconnect from SharePoint
Write-Host
Write-Host "Disconnecting from Services:" -ForegroundColor Cyan
Write-Host ("   Disconnecting from SharePoint Online: {0}" -f $spAdminURL)
Disconnect-SPOService

# Output Script Info
Complete-CatScript $scriptStart[0]

# Stop logging
Stop-Transcript
#endregion