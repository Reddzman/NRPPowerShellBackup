<#  Mitsubishi.SPO-ConfigSiteCollectionAdmin.ps1

This script iterates through all site collections in SharePoint Online and add a specified user to the Site Collection Administrators
for that site collection.

Dependencies
	-	SharePoint Online Module

Call script with parameters
	.\Mitsubishi-AddSiteCollectionAdmin.ps1  -spAdminURL "https://{tenant}-admin.sharepoint.com" -client "{File Abrivition}" -adminUPN = "admin@domain.com" -includePersonalSites $false -isSiteCollectionAdmin $true

Table of Contents
	1.	Variables
	2.	Snapin(s), Modules & Environment Connection(s)
	3.	Main Script
#>

#region Script Parameters
Param(	[string] $spAdminURL = "https://nrpdev-admin.sharepoint.com",
		[string] $client = "Testing",
		[string] $adminUPN = "nickp@nrpdev.com",
		[boolean] $includePersonalSites = $false,
		[boolean] $isSiteCollectionAdmin = $true
)
#endregion

#region Start Date / Time
# Collecting Start Time
$sDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
$fDate = Get-Date -Format "yyMMddHHmm"

# Return Start Time
$scriptStart = @()
$scriptStart += $sDate
$scriptStart += $fDate

# Start Transcript
Start-Transcript -Path (".\{0}.SPO-AddSiteCollectionAdmin-{1}.log" -f $client, $scriptStart[1])
#endregion

#region 1.	Variables
# Script Variables
If($isSiteCollectionAdmin){
	$adminAction = "Adding"
} 
Else {
	$adminAction = "Removing"
}

# Output Parameters & Variables
Write-Host
Write-Host "Script Parameters" -ForegroundColor Cyan
Write-Host ("   SPO Admin URL: {0}" -f $spAdminURL)
Write-Host ("   Client Abriviation: {0}" -f $client)
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
If ($includePersonalSites) {
	$spSites = Get-SPOSite -Limit All -IncludePersonalSite $includePersonalSites
}
Else {
	$spSites = Get-SPOSite -Limit All -IncludePersonalSite $includePersonalSites | Where{$_.URL -notlike "*-my.sharepoint.com*"}
}

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

# Finish Script Output
# Finish Time & Time Taken
$fTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
$tTime = New-TimeSpan -Start $scriptStart[0] -End $fTime

# Screen Output
Write-Host
Write-Host "-----------------------------------------------------------------------------------------------------" -ForegroundColor Cyan
Write-Host ("Started: {0}   Finished: {1}   Time Taken: {2}" -f $scriptStart[0], $fTime, $tTime) -ForegroundColor Cyan
Write-Host "-----------------------------------------------------------------------------------------------------" -ForegroundColor Cyan
Write-Host

# Stop logging
Stop-Transcript
#endregion