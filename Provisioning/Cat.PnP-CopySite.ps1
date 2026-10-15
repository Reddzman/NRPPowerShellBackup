<#  Cat.PnP-CopySite.ps1

This script uses PnP PowerShell to copy a sites configuration to another site.

Dependencies
	-	PnP PowerShell https://docs.microsoft.com/en-us/powershell/sharepoint/sharepoint-pnp/sharepoint-pnp-cmdlets?view=sharepoint-ps

Call script with parameters
	.\Cat.PnP-CopyScript.ps1 -sourceURL "https://spnickp.sharepoint.com/sites/spoadmin" -destinationURL "https://nrpdev.sharepoint.com/sites/spoadmin" -adminURL "https://nrpdev-admin.sharepoint.com" -createSite false -siteOnwer "nickp@nrpsolutions.com" -timezone 12

Table of Contents
	1.	Variables
	2.	Main Script
#>

#region Script Parameters
Param(  [String] $sourceURL = "https://spnickp.sharepoint.com/sites/spoadmin",
        [String] $destinationURL = "https://nrpdev.sharepoint.com/sites/spoadmin",
        [String] $adminURL = "https://nrpdev-admin.sharepoint.com",
        [Boolean] $createSite = $true,
        [String] $siteOwner = "nickp@nrpdev.onmicrosoft.com",
        [String] $timezone = 12
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
Write-Host "   Enter credentials for the source URL:"
$sourceCred = Get-Credential -Message "Enter your credentials for the source URL:"
Write-Host "   Enter credentials for the destination URL:"
$destCred = Get-Credential -Message "Enter your credentials for the destination URL:"

# Output Parameters & Variables
Write-Host
Write-Host "Script Parameters" -ForegroundColor Cyan
Write-Host ("   Source URL: {0}" -f $sourceURL)
Write-Host ("   Destination URL: {0}" -f $destinationURL)
Write-Host ("   Admin URL: {0}" -f $adminURL)
Write-Host ("   Create Site: {0}" -f $createSite)
Write-Host
Write-Host "Script Variables" -ForegroundColor Cyan
Write-Host ("   Source Credentials: {0}" -f $sourceCred.UserName)
Write-Host ("   Destination Credentials: {0}" -f $destCred.UserName)
#endregion
    
#region 2.	Main Script
# Connecting to source site
Write-Host
Write-Host ("Connecting to source site: {0}" -f $sourceURL) -ForegroundColor Magenta
Connect-PnPOnline -Url $sourceURL -Credential $sourceCred

# 3.	Get Provisioning Template
Write-Host ("   Extracting provisioning template: {0}" -f $sourceURL)
$provTemp = Get-PnPProvisioningTemplate -OutputInstance -ExcludeContentTypesFromSyndication -PersistBrandingFiles -ErrorAction Stop

# Creating destination site
if ($createSite){
    # Get site details
    $site = Get-PnPSite -Includes RootWeb

    # Connect to on SPO Admin Center
    Write-Host
    Write-Host ("Connecting to Destination Admin Center: {0}" -f $adminURL) -ForegroundColor Magenta
    Connect-PnPOnline -Url $adminURL -Credentials $destCred

    # 4.	Create New Site Collection
    Write-Host ("   Creating destination site: {0}" -f $destinationURL)
    if (($provTemp.BaseSiteTemplate -like "GROUP*") -or ($provTemp.BaseSiteTemplate -like "SITEPAGES*")) {
        If ($provTemp.BaseSiteTemplate -contains "GROUP") {$type = "TeamSite"} Else {$type = "CommunicationSite"}
        New-PnPSite -Title $site.RootWeb.Title -Url $destinationURL -Type $type
    }
    Else {
        New-PnPTenantSite -Template $provTemp.BaseSiteTemplate -Title $site.RootWeb.Title -Url $destinationURL -Owner $siteOwner -TimeZone $timezone -Wait
    }
} # if ($createSite)

# Connect to on SPO Admin Center
Write-Host
Write-Host ("Connecting to destination site: {0}" -f $destinationURL) -ForegroundColor Magenta
Connect-PnPOnline -Url $destinationURL -Credentials $destCred

# 5.	Applying Provisioning Template
Write-Host ("   Applying Provisioning Template: {0}" -f $destinationURL)
Apply-PnPProvisioningTemplate -InputInstance $provTemp

# 6.	Disable Limited-access user permission lockdown mode feature
# 7.	Enable other desired features

# Screen Output for completion
$SPOSite = Get-PnPSite -Includes RootWeb
Write-Host ("   {0} site collection created sucessfully" -f $SPOSite.RootWeb.Title) -ForegroundColor Green
# Disconnect from SPO #>
Disconnect-PnPOnline
#endregion

#region Completed Screen Output
Complete-CatScript $scriptStart[0]
#endregion