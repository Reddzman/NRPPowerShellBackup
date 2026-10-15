<# Cat.PnP-CreateClassicSite.ps1

This script creates a new site classic site collection and applys a PnP provisioning template.

Dependencies
	-	PnP PowerShell https://msdn.microsoft.com/en-us/pnp_powershell/pnp-powershell-overview

Command
	.\Cat.PnP-CreateClassicSite.ps1 -adminURL "https://nrpdev-admin.sharepoint.com" -siteURL "https://nrpdev.sharepoint.com/sites/test01" -siteName "Test Site" -siteOwner "nickp@nrpdev.onmicrosoft.com" -provTemplate "SPOAdmin.pnp" -siteTemplate "STS#0" -timezone 12 -resourceQuota 100 -storagequota 4048 -locale = 1033

Table of Contents
	1.	Input & Output Files
	2.	Variables
	3.	Main Script
	4.	Create New Site Collection
	5.	Applying Provisioning Template
#>

#region Script Parameters
Param(	[string] $adminURL = "https://nrpdev-admin.sharepoint.com",
        [string] $siteURL = "https://nrpdev.sharepoint.com/sites/test01",
        [string] $siteName = "Test Site",
        [string] $siteOwner = "nickp@nrpdev.onmicrosoft.com",
		[string] $provTemplate = "ClassicTeamTemplate.pnp",
		[string] $siteTemplate = "STS#0",
		[int] $timezone = 12,
		[int] $resourceQuota = 100,
		[int] $storagequota = 4048,
		[int] $locale = 1033
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

#region 1.	Input & Output Files
# File Paths
$provTemp = ("..\_Templates\{0}" -f $provTemplate)
#endregion

#region 2.	Variables
# User Input Variables
Write-Host
Write-Host "User Input Variables" -ForegroundColor Cyan
Write-Host "   Enter credentials for SharePoint 2013 on-premises:"
$spCred = Get-Credential -Message "Enter your credentials for SharePoint:" 

# Output Parameters & Variables
Write-Host
Write-Host "Script Parameters" -ForegroundColor Cyan
Write-Host ("   Admin URL: {0}" -f $adminURL)
Write-Host ("   Site URL: {0}" -f $siteURL)
Write-Host ("   Site Name: {0}" -f $siteName)
Write-Host ("   Site Owner: {0}" -f $siteOwner)
Write-Host ("   Timezone: {0}" -f $timezone)
Write-Host
Write-Host "Script Variables" -ForegroundColor Cyan
Write-Host ("   Credentials: {0}" -f $spCred.UserName)
#endregion
    
#region 3.	Main Script
try {
	#region 4.	Create New Site Collection
	Write-Host
	Write-Host ("Creating Site Collection: {0}" -f $siteURL) -ForegroundColor Magenta

	# Connect to on SPO Admin Center
	Write-Host "   Connecting to SharePoint Online Admin Center"
	Connect-PnPOnline -Url $adminURL -Credentials $spCred

	# Creating Site
	Write-Host ("   Creating SharePoint Online site: {0} " -f $siteURL) -NoNewline
	New-PnPTenantSite -Title $siteName -Url $siteURL -Owner $siteOwner -ResourceQuota $resourceQuota -StorageQuota $storageQuota -Lcid $locale -TimeZone $timezone -Wait -ErrorAction Stop
	Write-Host
	#endregion

	#region 5.	Applying Provisioning Template
	Write-Host
	Write-Host ("Applying Provisioning Template: {0}" -f $provTemp) -ForegroundColor Magenta

	# Connecting to new site
	Write-Host ("   Connecting to site: {0}" -f $siteURL)
	Connect-PnPOnline -Url $siteURL -Credentials $spCred

	# Applying provisioning template
	Write-Host ("   Applying provisioning template: {0}" -f $provTemp)
	Apply-PnPProvisioningTemplate -Path $provTemp
	#endregion

	# Screen Output for completion
	Write-Host
	Write-Host ("{0} created sucessfully" -f $siteName) -ForegroundColor Green	
}
catch {
	Write-Host "!! EXCEPTION OCCURRED !!" -ForegroundColor Red
	Write-Host "!! Exception Type: $($_.Exception.GetType().FullName)" -ForegroundColor Red
    Write-Host "!! Exception Message: $($_.Exception.Message)" -ForegroundColor Red
}
#endregion

#region Completed Screen Output
# Disconnect from SPO
Disconnect-PnPOnline
Complete-CatScript $scriptStart[0]
#endregion