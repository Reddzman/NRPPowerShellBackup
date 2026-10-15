<#
Cat.PnP-DisableModern.ps1

This script disables the modern UI Features on an classic team site.

Dependencies
	-	PnP PowerShell https://docs.microsoft.com/en-us/powershell/sharepoint/sharepoint-pnp/sharepoint-pnp-cmdlets?view=sharepoint-ps

Call Script with Input File
	.\Cat.Tech-ScriptName.ps1 -SiteURL "https://spnickp.sharepoint.com/sites/temp-classicteam"

	Table of Contents
	1.	Variables
	2.	Main Script
#>

#region Script Parameters
Param([string]$SiteUrl = "https://spnickp.sharepoint.com/sites/temp-classicteam")
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

#region 1.	Variables
# User Inputs
Write-Host
Write-Host "Input Variables" -ForegroundColor DarkCyan
Write-Host "   Enter credentials for SharePoint:"
$SPCred = Get-Credential -Message "Enter your credentials for SharePoint:"
$DisableFeatures = @("b6917cb1-93a0-4b97-a84d-7cf49975d4ec";"d95c97f3-e528-4da2-ae9f-32b3535fbb59")
$EnableFeatures = @("e3540c7d-6bea-403c-a224-1a12eafee4c4")
#endregion
    
#region 2.	Main Script
# Connecting To SP Online
Write-Host
Write-Host ("Connecting to Site: {0}" -f $SiteUrl) -ForegroundColor Magenta
Connect-PnPOnline -Url $SiteURL -Credentials $SPCred

# Disabling Features
Write-Host "   Disabling Features"
ForEach ($Feature in $DisableFeatures) {
	Disable-PnPFeature -identity $Feature -Scope Web
	Write-Host ("      Feature Disabled: {0}" -f $Feature)
}

# Enabling Features\
Write-Host "   Enabling Features"
ForEach ($Feature in $EnableFeatures){
	Enable-PnPFeature -Identity $Feature -Scope Site
	Write-Host ("      Feature Enabled: {0}" -f $Feature)
}
#endregion

#region Completed Screen Output
Complete-CatScript $Start[0]
#endregion
