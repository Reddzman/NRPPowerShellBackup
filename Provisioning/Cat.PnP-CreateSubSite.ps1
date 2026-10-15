<#
Cat.PnP-CreateSubSite.ps1

This script creates a sub site based on a PnP Provisioning Template

Dependencies
    -	PnP PowerShell https://docs.microsoft.com/en-us/powershell/sharepoint/sharepoint-pnp/sharepoint-pnp-cmdlets?view=sharepoint-ps
    -   Template must be created at the web level not site level

Call Script with Input File
	.\Cat.Tech-ScriptName.ps1 -InputFile InputFileName.csv

Table of Contents
	1.	Input & Output Files
	2.	Variables
	3.	Main Script
#>

#region Script Parameters
Param(  [string]$SCURL = "https://spnickp.sharepoint.com/sites/temp-classicteam",
        [string]$WebTemplate = "STS#0",
        [int]$Locale = 1033,
        [string]$SiteTitle = "Test Site O1",
        [string]$SitePath = "testsite01",
        [string]$Template = "Template.pnp")
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
$InPath = ("{0}_Templates\{1}" -f $Location, $Template)
#endregion

#region 2.	Variables
# User Inputs
Write-Host
Write-Host "User Input Variables" -ForegroundColor DarkCyan
Write-Host "   Enter credentials for SharePoint:"
$SPCred = Get-Credential -Message "Enter your credentials for SharePoint:"

# Script Variables
$DisableFeatures = @() # Mobile Browser View; Site Pages
$HideLists = @()
#endregion
    
#region 3.	Main Script
try{
    # Connect to Parent Site
    Write-Host
    Write-Host ("Connecting to Site: {0}" -f $SCURL) -ForegroundColor Magenta
    Connect-PnPOnline -Url $SCURL -Credentials $SPCred

    # Create Site based on Publishing site template
    Write-Host ("   Creating Site: {0}" -f $SiteTitle)
    $NewWeb = New-PnPWeb -Title $SiteTitle -Url $SitePath -Locale $Locale -Template $WebTemplate

    # Apply Provisioning Template
    Write-Host ("   Applying Provisioing Template: {0}" -f $InPath)
    Apply-PnPProvisioningTemplate -Path $InPath -Web $NewWeb -ErrorAction Stop

    # Disable features
    Write-Host "   Disabling Features"
    ForEach ($Feature in $DisableFeatures){
        Disable-PnPFeature -Identity $Feature -Scope Web -Web $NewWeb
    }

    # Hide Libraries
    Write-Host "   Hiding Lists"
    ForEach ($List in $HideLists){
        $CurList = Get-PnPList -Identity $List -Web $NewWeb
        Set-PnPList -Identity $CurList -Hidden $true -Web $NewWeb
    }

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