<#  Cat.ADO-TestPipelineToM365.ps1

This script is designed to test script connections to M365 via an Azure App Registration while using an ADO Repo and Pipeline to execute the script. Should be executed on a ShareGate migration server to confirm capability.

Dependencies
    -   Azure App Registration with permissions to M365
    -   Azure certificate for authentication via app registration
    -   PnP.PowerShell Modules

Call script with parameters
	.\Cat.ADO-TestPipelineToM365.ps1 -tenant nrpdev2 -appID "caf6b11a-cf03-416c-b0f0-67a09a8e30ee" -certPath "C:\GitHub\NRPDev\AppRegistration\PowerShell App Registration.pfx" -certPass "Password"

Table of Contents
	1.	Variables
	2.	Snapin(s), Modules & Environment Connection(s)
    3.	Main Script
#>

#region Script Parameters
Param(
    [string] $tenant = "nprdev2",
    [String] $appID = "caf6b11a-cf03-416c-b0f0-67a09a8e30ee",
    [String] $certPath = "C:\GitHub\NRPDev\AppRegistration\PowerShell App Registration.pfx",
    [String] $certPass = ""
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
Start-Transcript -Path ("..\_Log Files\{0}.ADO-TestPipelineToM365-{1}.log" -f $tenant, $scriptStart[1])
#endregion

#region 1.	Variables
# Script Variables
$pass = ConvertTo-SecureString -String $certPass -AsPlainText -Force
Remove-Variable certPass
$adminUrl = ("https://{0}-admin.sharepoint.com/" -f $tenant)
$rootSiteUrl = ("https://{0}.sharepoint.com/" -f $tenant)
$tenantDomain = ("{0}.onmicrosoft.com" -f $tenant)
$cert = Get-PnPAzureCertificate -CertificatePath $certPath -CertificatePassword $pass

# Variable Output
Write-Host
Write-Host "Script Parameters" -ForegroundColor Cyan
Write-Host ("   Tenant: {0}" -f $tenant)
Write-Host ("   App ID: {0}" -f $appID)
Write-Host ("   Cert Path: {0}" -f $certPath)
Write-Host
Write-Host "Script Variables" -ForegroundColor Cyan
Write-Host ("   Admin Url: {0}" -f $adminUrl)
Write-Host ("   Root Site Url: {0}" -f $rootSiteUrl)
Write-Host ("   Tenant Url: {0}" -f $tenantDomain)
#endregion

#region 2.	Snapin(s), Modules & Environment Connection(s)
# Load MS / PnP PowerShell Modules
Get-Module -ListAvailable | Where-Object{($_.Name -like "*SharePoint*") -or ($_.Name -like "*PnP*") } | Remove-Module
Import-Module PnP.PowerShell
#endregion
    
#region 3.	Main Script
# Connecting To SP Online
Write-Host
Write-Host "Testing Connection to Services:" -ForegroundColor Cyan

# Testing connection to SharePoint Online Administration
Write-Host
Write-Host ("   Connecting to SharePoint Online Administration: {0}" -f $adminUrl) -ForegroundColor Magenta
Connect-PnPOnline -URL $adminUrl -tenant $tenantDomain -ClientId $appID -Thumbprint $cert.Thumbprint
Write-Host "      Retrieving Tenant Sites" -NoNewline
$sites = Get-PnPTenantSite
Write-Host (" - Site Count: {0}" -f $sites.Count)
Write-Host ("      Disconnecting from SharePoint Online Administration: {0}" -f $adminUrl)
Disconnect-PnPOnline

# Testing Connection to SharePoint Online Root Site Collection
Write-Host
Write-Host ("   Connecting to SharePoint Online root site: {0}" -f $rootSiteUrl) -ForegroundColor Magenta
Connect-PnPOnline $rootSiteUrl -tenant $tenantDomain -ClientId $appID -Thumbprint $cert.Thumbprint
Write-Host "      Retrieving Site Object" -NoNewline
$site = Get-PnPSite -Includes Id
Write-Host (" - Site ID: {0}" -f $site.Id)
Write-Host ("      Disconnecting from SharePoint Online root site: {0}" -f $rootSiteUrl)
Disconnect-PnPOnline
#endregion

#region Completed Script
# Output Script Info
Complete-CatScript $scriptStart[0]

# Stop logging
Stop-Transcript
#endregion