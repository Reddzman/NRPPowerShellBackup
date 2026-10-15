<#  Cat.Tech-ScriptName.ps1

This script configure group creation in your tenant to be limited to a specific M365 group.  

Dependencies
	-	AzureADPreview modules
	-	Run in Windows PowerShell instead of PowerShell core

Call script with parameters
	Run using App Reg	 -->	.\Cat.AAD-ConfigGroupCreation.ps1 -groupName "Site and Group Creation" -clientTenant "nrpdev" -clientTenantID "2811aa92-b6b5-458a-ab91-076e9ce5e8b3" -clientAppID "caf6b11a-cf03-416c-b0f0-67a09a8e30ee" -clientCertPath "C:\GitHub\NRPDev\AppRegistration\PowerShell App Registration.pfx" -allowGroupCreation $false
	Run using Modern Auth Prompt	-->	.\Cat.AAD-ConfigGroupCreation.ps1 -groupName "Site and Group Creation" -clientTenant "nrpdev" -clientTenantID $null -clientAppID $null -clientCertPath $null -allowGroupCreation $false

Table of Contents
	1.	Snapin(s), Modules & Environment Connection(s)
	2.	Variables
	3.	Main Script
#>

#region Script Parameters
Param(	[string] $groupName = "Group Name",
		[string] $clientTenant = "tenantprefix", # tenantprefix.sharepoint.com
		[string] $clientTenantID = "AAD Tenant ID",
		[string] $clientAppID = "Azure App Registration ID", # set to $null to be prompted for modern authentication
		[string] $clientCertPath = "Full Certificate Path",
		[bool] $allowGroupCreation = $false # $false disables group creation
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
Start-Transcript -Path ("..\_Log Files\{0}.AAD-ConfigGroupCreation-{1}.log" -f $clientTenant, $scriptStart[1])
#endregion

#region 1.	Snapin(s), Modules & Environment Connection(s)
# Load MS / PnP PowerShell Modules
Import-Module AzureADPreview #-UseWindowsPowerShell
#endregion

#region 2.	Variables
# User Inputs
# If using App Registration
If ($clientAppID){
	Write-Host
	Write-Host "User Input Variables" -ForegroundColor Cyan
	$certPass = ConvertTo-SecureString -String (Read-Host -Prompt "   Enter certificate password for the app registration") -AsPlainText -Force
	$cert = Get-PnPAzureCertificate -Path $clientCertPath -Password $certPass
}

# Output Parameters & Variables
Write-Host
Write-Host "Script Parameters" -ForegroundColor Cyan
Write-Host ("   Group Name: {0}" -f $groupName)
Write-Host ("   Tenant Domain {0}" -f $clientTenant)
Write-Host ("   AAD Tenant ID: {0}" -f $clientTenantID)
Write-Host ("   App ID: {0}" -f $clientAppID)
Write-Host ("   Certificate Path: {0}" -f $clientCertPath)
Write-Host ("   Allow Group Creation: {0}" -f $allowGroupCreation)
Write-Host
Write-Host "Certificate Thumbprint" -ForegroundColor Cyan
Write-Host $cert.Thumbprint
#endregion
    
#region 3.	Main Script
# Connecting To SP Online
Write-Host
Write-Host "Connecting to Services:" -ForegroundColor Cyan
Write-Host "   Connecting to Azure AD"
If ($clientAppID){
	# Connect to AAD via App Registration
	Connect-AzureAD -TenantId $clientTenantID -CertificateThumbprint $cert.Thumbprint -ApplicationId $clientAppID
} Else {
	# Connect to AAD via modern login
	Connect-AzureAD
}

# Retrieve unified group settings ID
Write-Host
Write-Host "Configuring Group Creation:" -ForegroundColor Cyan
Write-Host "   Retrieving unified group settings"
$settingsObjectID = (Get-AzureADDirectorySetting | Where-object -Property Displayname -Value "Group.Unified" -EQ).id

# Creating new settings if were not able to find unified group settings
if(!$settingsObjectID){
	Write-Host "   Creating new unified group settings"
    $template = Get-AzureADDirectorySettingTemplate | Where-object {$_.displayname -eq "group.unified"}
    $settingsCopy = $template.CreateDirectorySetting()
    $temp = New-AzureADDirectorySetting -DirectorySetting $settingsCopy
    $settingsObjectID = (Get-AzureADDirectorySetting | Where-object -Property Displayname -Value "Group.Unified" -EQ).id
}

# Updating unified group settings
Write-Host "   Copying Unified group Settings"
$settingsCopy = Get-AzureADDirectorySetting -Id $settingsObjectID

# Updating unified group settings
Write-Host ("   Setting allow group creation to {0}" -f $allowGroupCreation)
$settingsCopy["EnableGroupCreation"] = $allowGroupCreation

# Updating unified group settings
Write-Host ("   Configuring security group to {0}" -f $groupName)
if($groupName){
	$settingsCopy["GroupCreationAllowedGroupId"] = (Get-AzureADGroup -SearchString $groupName).objectid
} else {
	$settingsCopy["GroupCreationAllowedGroupId"] = $groupName
}

# Saving settings to Azure AD
Write-Host "   Saving settings to Azure AD"
Set-AzureADDirectorySetting -Id $settingsObjectID -DirectorySetting $settingsCopy

# Output final Settings
Write-Host
Write-Host "Current Group Creation Settings" -ForegroundColor Cyan
(Get-AzureADDirectorySetting -Id $settingsObjectID).Values | Format-Table Name, Value
#endregion

#region Completed Script
# Disconnect from Services
Write-Host
Write-Host "Disconnecting from Services:" -ForegroundColor Cyan
Write-Host "   Disconnecting from Azure AD"
Disconnect-AzureAD

# Output Script Info
Complete-CatScript $scriptStart[0]

# Stop logging
Stop-Transcript
#endregion