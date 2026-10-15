<#
Cat.PnP-GetProvisioningTemplate.ps1

This script extracts a PnP Provisioning Template from an existing site or web.

Dependencies
	-	PnP PowerShell https://docs.microsoft.com/en-us/powershell/sharepoint/sharepoint-pnp/sharepoint-pnp-cmdlets?view=sharepoint-ps

Call Script with Input File
    .\Cat.Tech-ScriptName.ps1 -SiteURL "https://spnickp.sharepoint.com/sites/temp-classicteam" -Filename ClassicTeamSite.pnp
    
After creation:
    -   remove unnecessary lists.  IE. Site Assets
    -   remove unnecessary files. IE. Seattle.master

Table of Contents
	1.	Input & Output Files
	2.	Variables
	3.	Main Script
#>

#region Script Parameters
Param(  [string]$SiteUrl = "https://spnickp.sharepoint.com/sites/temp-classicteam",
        [string]$Filename = "ClassicTeamSite.pnp")
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
$OutPath = ("{0}_Templates\{1}" -f $Location, $Filename)
#endregion

#region 2.	Variables
# User Inputs
Write-Host
Write-Host "User Input Variables" -ForegroundColor DarkCyan
Write-Host "   Enter credentials for SharePoint:"
$SPCred = Get-Credential -Message "Enter your credentials for SharePoint:"

# Script Variables
#$ExcludeHandlers = @("Publishing";"PageContent") # Use to exlude certain features
#endregion
    
#region 3.	Main Script
try{
    # Connecting to Template Site
    Write-Host
    Write-Host ("Connecting to Template Site: {0}" -f $SiteURL) -ForegroundColor Magenta
    Connect-PnPOnline -Url $SiteURL -Credentials $SPCred

    # PnP Logging
    #Set-PnPTraceLog -On -LogFile ("$home\desktop\Cat.PnP-GetProvisioningTemplate-{0}.log" -f $Start[1]) -Level Debug

    # Extract Template
    Write-Host ("   Extracting Template: {0}" -f $OutPath)
    Get-PnPProvisioningTemplate -Out $OutPath <#-ExcludeHandlers $ExcludeHandlers#> -ExcludeContentTypesFromSyndication -PersistBrandingFiles -Force -ErrorAction Stop
    
    # PnP Logging
    #$error[0].Exception.Stacktrace
    #Set-PnPTraceLog -Off

    # Script Complete
    Write-Host ("   Template Creation Complete") -ForegroundColor Green
}
catch{
    Write-Host
	Write-Host "!! EXCEPTION OCCURRED !!" -ForegroundColor Red
	Write-Host "!! Exception Type: $($_.Exception.GetType().FullName)" -ForegroundColor Red
    Write-Host "!! Exception Message: $($_.Exception.Message)" -ForegroundColor Red
}
#endregion

#region Completed Screen Output
Complete-CatScript $Start[0]
#endregion