<#  Cat.M365DSC-TenantConfigExport.ps1

This script leveraged the Microsoft 365 Desired State Configurator to extract setting and configuration information from the a 
tenant.  This is best with an app registration.  User Accounts currently require basic authentication.

Dependencies
	-	Microsoft365DSC - https://microsoft365dsc.com/

Call script with parameters
    cd "C:\GitHub\CatPowerShell\Reporting\"
	.\Cat.M365DSC-TenantConfigExport.ps1 -clientAppID "caf6b11a-cf03-416c-b0f0-67a09a8e30ee" -clientCertPath "C:\GitHub\NRPDev\AppRegistration\PowerShell App Registration.pfx" -clientTenant "nrpdev" -configurationName "nrpdev2-M365DSC-230127" -outputFileName "nrpdev2-M365DSC-230127.ps1" -generateInfo $true -maxProcesses 10 -mode "Lite" -path "..\Output Files\nrpdev2-M365DSC-230127" -clientWorkloads @("OD", "SPO", "Teams")

Table of Contents
	1.	Snapin(s), Modules & Environment Connection(s)
	2.	Variables
	3.	Main Script
#>

#region Script Parameters
Param(	[string] $clientAppID = "{Client app registration ID from Azure}", # Leave Blank for interactive login
		[string] $clientCertPath = "{Path to client certificate file must be installed on local user}",# Leave Blank for interactive login
		[string] $clientTenant = "{Client tenant SharePoint prefix}", # IE. ClientTenant.sharepoint.com
        [string] $configurationName = "{Configuration Name}",
        [string] $outputFileName = "{Filename}.ps1", # output will be a PowerShell File (.ps1)
        [boolean] $generateInfo = $true,
        [int] $maxProcesses = 10,
        [string] $mode = "Full", # Options: Lite, Default, Full
        [string] $path = "{Output Path}",
        [array] $clientWorkloads = @("OD", "SPO", "Teams")
            <# Options: AAD – Azure Active Directory
                        EXO – Exchange Online
                        O365 – Office 365 administration
                        Intune – Intune
                        OD – OneDrive
                        Planner – Planner
                        PP – Power Platform
                        SC – Security and Compliance
                        SPO – SharePoint Online
                        Teams – Microsoft Teams
            #>
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
Start-Transcript -Path ("..\_Log Files\{0}.M365DSC-TenantConfigExport-{1}.log" -f $clientTenant, $scriptStart[1])
#endregion

#region 1.	Snapin(s), Modules & Environment Connection(s)
# Load MS / PnP PowerShell Modules
Import-Module Microsoft365DSC
#endregion

#region 2.	Variables
# User Inputs
Write-Host
Write-Host "User Input Variables" -ForegroundColor Cyan
$certPass = ConvertTo-SecureString -String (Read-Host -Prompt "   Enter certificate password for the app registration") -AsPlainText -Force
$cert = Get-PnPAzureCertificate -Path $clientCertPath -Password $certPass
Remove-Variable certPass
Remove-Variable clientCertPath

# Script Variables
$clientTenantDomain = ("{0}.onmicrosoft.com" -f $clientTenant)

# Output Parameters & Variables
Write-Host
Write-Host "Script Parameters" -ForegroundColor Cyan
Write-Host ("   App ID: {0}" -f <#"Show Only If Troublshooting",#> $clientAppID)
Write-Host ("   Certificate Path: {0}" -f $clientCertPath)
Write-Host ("   Tenant Abbreviation: {0}" -f $clientTenant)
Write-Host ("   Configuration Name: {0}" -f $configurationName)
Write-Host ("   File Name: {0}" -f $outputFileName)
Write-Host ("   Generate Information: {0}" -f $generateInfo)
Write-Host ("   Max Processes: {0}" -f $maxProcesses)
Write-Host ("   Mode: {0}" -f $mode)
Write-Host ("   Path: {0}" -f $path)
Write-Host
Write-Host "User Inputs" -ForegroundColor Cyan
Write-Host ("   Certificate Thumbprint: {0}" -f <#"Show Only If Troublshooting",#> $cert.Thumbprint)
Write-Host
Write-Host "Script Variables" -ForegroundColor Cyan
Write-Host ("   Microsoft Tenant: {0}" -f $clientTenantDomain)
Write-Host
Write-Host "Workloads" -ForegroundColor Cyan
Write-Host $clientWorkloads
#endregion
    
#region 3.	Main Script
# Exporting resources using credentials
Write-Host
Write-Host "Exporting Tenant Configuration" -ForegroundColor Cyan
Write-Host
Export-M365DSCConfiguration -ApplicationId $clientAppID `
    -TenantID $clientTenantDomain `
    -CertificateThumbprint $cert.Thumbprint `
    -ConfigurationName $configurationName `
    -FileName $outputFileName `
    -GenerateInfo:$generateInfo `
    -MaxProcesses 10 `
    -Mode $mode `
    -Workloads $clientWorkloads `
    -Path $path
#endregion

#region Completed Script
# Output Script Info
Complete-CatScript $scriptStart[0]

# Stop logging
Stop-Transcript
#endregion