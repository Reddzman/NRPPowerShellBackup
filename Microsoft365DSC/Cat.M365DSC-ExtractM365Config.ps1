<#  Cat.M365DSC-ExtractM365Config.ps1 Description

This script leverages the Micrsoft365 Desired State Configuration utility to extract configuration and settings from the specified M365 environment.

Dependencies
	-	Microsoft365DSC PowerShell Modules

Call script with parameters
	cd C:\GitHub\CatPowerShell\Microsoft365DSC
	.\Cat.M365DSC-ExtractM365Config.ps1 -exportName "SPAndTeamsConfig" -clientAppID "caf6b11a-cf03-416c-b0f0-67a09a8e30ee" -clientCertPath "C:\GitHub\NRPDev\AppRegistration\PowerShell App Registration.pfx" -clientTenant "nrpdev2"

Table of Contents
	1.	Snapin(s), Modules & Environment Connection(s)
	2.	Input & Output Files
	3.	Variables
	4.	Main Script
	5.	CSV Loop Example
	6.	...
#>

#region Script Parameters
Param(	[string] $exportName = "{Name used for configuration & export filename}",
		[string] $clientAppID = "{Client app registration ID from Azure}", # Leave Blank for interactive login
		[string] $clientCertPath = "{Path to client certificate file must be installed on local user}",# Leave Blank for interactive login
		[string] $clientTenant = "{Client tenant SharePoint prefix}" # IE. ClientTenant.sharepoint.com
)
#endregion Script Parameters

#region Start Date / Time
# Remove Existing Catapult Modules to ensure PS session has most up to date functions
Get-Module Cat.* | Remove-Module

# Load General Functions
Import-Module "..\General\Cat.GeneralFunctions.psm1"

# Get Start Date / Time
$scriptStart = Start-CatScript

# Start Transcript
Start-Transcript -Path ("..\_Log Files\{0}.M365DSC-ExtractM365Config-{1}.log" -f $clientTenant, $scriptStart[1])
#endregion Start Date / Time

#region 1.	Snapin(s), Modules & Environment Connection(s)
# Load MS / PnP PowerShell Modules
Import-Module Microsoft365DSC
#endregion Snapin(s), Modules & Environment Connection(s)

#region 2.	Input & Output Files
# File Paths
$outputPath = ("C:\GitHub\CatPowerShell\_Output Files" -f $clientTenant, $scriptStart[1])
$outputFile = ("{0}-{1}-{2}.ps1" -f $clientTenant, $exportName, $scriptStart[1])
$configName = ("{0}-{1}-{2}" -f $exportName, $clientTenant, $scriptStart[1])
#endregion Input & Output Files

#region 3.	Variables
# User Inputs
If ($clientAppID) {
	Write-Host
	Write-Host "User Input Variables" -ForegroundColor Cyan
	$certPass = ConvertTo-SecureString -String (Read-Host -Prompt "   Enter certificate password for the app registration") -AsPlainText -Force
	$cert = Get-PnPAzureCertificate -Path $clientCertPath -Password $certPass
	Remove-Variable certPass
	Remove-Variable clientCertPath
}

# Script Variables
$tenantId = ("{0}.onmicrosoft.com" -f $clientTenant)
# Baseline SharePoint, OneDrive, and Teams Components
$dscComponents = @("ODSettings", "SPOAccessControlSettings", "SPOApp", "SPOBrowserIdleSignout", "SPOHomeSite", "SPOHubSite", "SPOOrgAssetsLibrary", `
	"SPOSearchManagedProperty", "SPOSearchResultSource", "SPOSharingSettings", "SPOSiteDesign", "SPOSiteDesignRights", "SPOSiteScript", "SPOStorageEntity", `
	"SPOTenantCdnEnabled", "SPOTenantCdnPolicy", "SPOTenantSettings", "SPOTheme", "TeamsAppPermissionPolicy", "TeamsAppSetupPolicy", "TeamsChannel", `
	"TeamsChannelsPolicy", "TeamsClientConfiguration", "TeamsComplianceRecordingPolicy", "TeamsCortanaPolicy", "TeamsDialInConferencingTenantSettings", `
	"TeamsEnhancedEncryptionPolicy", "TeamsEventsPolicy", "TeamsFederationConfiguration", "TeamsFeedbackPolicy", "TeamsFilesPolicy", "TeamsGroupPolicyAssignment", `
	"TeamsGuestMessagingConfiguration", "TeamsMessagingPolicy", "TeamsMobilityPolicy", "TeamsNetworkRoamingPolicy", "TeamsOrgWideAppSettings", "TeamsShiftsPolicy", `
	"TeamsTemplatesPolicy", "TeamsTenantNetworkRegion", "TeamsTenantNetworkSite", "TeamsTenantNetworkSubnet", "TeamsTenantTrustedIPAddress", "TeamsTranslationRule", `
	"TeamsUnassignedNumberTreatment", "TeamsUpdateManagementPolicy", "TeamsUpgradeConfiguration", "TeamsUpgradePolicy", "TeamsUserPolicyAssignment", "TeamsVdiPolicy", `
	"TeamsWorkloadPolicy") #>

#region Output Parameters & Variables
Write-Host
Write-Host "Script Parameters" -ForegroundColor Cyan
Write-Host ("   Export Name: {0}" -f $exportName)
if ($clientAppID) {
	Write-Host ("   App ID: {0}" -f $clientAppID)
}
Write-Host ("   Client Tenant Prefix: {0}" -f $clientTenant)
Write-Host
Write-Host "Output Variables" -ForegroundColor Cyan
Write-Host ("   Output Path: {0}" -f $outputPath)
Write-Host ("   Output File: {0}" -f $outputFile)
If ($clientAppID) {
	Write-Host
	Write-Host "User Inputs" -ForegroundColor Cyan
	Write-Host ("   Certificate Thumbprint: {0}" -f "Only output for troubleshooting", $cert.Thumbprint)
}
Write-Host
Write-Host "Script Variables" -ForegroundColor Cyan
Write-Host ("   Tenant Id: {0}" -f $tenantId)
Write-Host
Write-Host "Components Array" -ForegroundColor Cyan
Write-Host $dscComponents
#endregion Output Parameters & Variables
#endregion
    
#region 4.	Main Script
Export-M365DSCConfiguration `
	-Components $dscComponents `
	-Path $outputPath `
	-ConfigurationName $configName `
	-FileName $outputFile `
	-ApplicationId $clientAppID `
	-TenantId $tenantId `
	-CertificateThumbprint $cert.Thumbprint

	New-M365DSCReportFromConfiguration -Type Excel -ConfigurationPath ("{0}\{1}" -f $outputPath, $outputFile) -OutputPath ("{0}\{1}" -f $outputPath, $outputFile.Replace(".ps1",".xlsx"))
#endregion Main Script

#region Completed Script
# Output Script Info
Complete-CatScript $scriptStart[0]

# Stop logging
Stop-Transcript
#endregion Completed Script