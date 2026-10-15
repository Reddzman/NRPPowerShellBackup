<#  Cat.PnP-CreateAppRegistration.ps1

This script leveraged PnP.PowerShell to create an app registration in a clients tenant. At the same time that it creates
the app registration, it also creates a certificate and stores it in the _Output Files folder. A global administrator will 
need to consent to the requested permissions.  You may need to add permissions in certain scenarios.

Dependencies
	-	PnP PowerShell https://pnp.github.io/powershell/
	-	General Functions - ..\General\Cat.GeneralFunctions.psm1

Call script with parameters
	.\Cat.PnP-CreateAppRegistration.ps1  -appName "Test App Registration 02" -clientTenant "nrpdev2" -certCommonName "testappreg02.nrpdev.com" -certStore "CurrentUser" -certValidYears 10 -certCountry "United States" -certState "Colorado" -certCity "Arvada" -certOrg "NRP Solutions" -certDept "Consulting"

Table of Contents
	1.	Variables
	3.	Snapin(s), Modules & Environment Connection(s)
	4.	Main Script
#>

#region Script Parameters
Param(	[string] $appName = "Catapult PowerShell App Registration",
		[string] $clientTenant = "[ClientTenantPrefix]",
		[string] $certCommonName = "CatPowerShell.[client].com",
		[string] $certStore = "CurrentUser",
		[int] $certValidYears = 1, # new certs should be created each year or on a schedule the client is comfortable with
		[string] $certCountry = "[ClientCountry]",
		[string] $certState = "[ClientState]",
		[string] $certCity = "[ClientCity]",
		[string] $certOrg = "[ClientName]",
		[string] $certDept = "Information Technology"
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
Start-Transcript -Path ("..\_Log Files\{0}.PnP-CreateAppRegistration-{1}.log" -f $clientTenant, $scriptStart[1])
#endregion

#region 1.	Variables
# User Inputs
Write-Host
Write-Host "User Input Variables" -ForegroundColor Cyan
$certPass = ConvertTo-SecureString -String (Read-Host -Prompt "   Enter certificate password for the app registration") -AsPlainText -Force

# Script Variables
$appGraphAppPerms = @(
	"Directory.ReadWrite.All",
	"Group.ReadWrite.All",
	"People.Read.All",
	"Sites.FullControl.All",
	"User.ReadWrite.All"
)
$appGraphDelegPerms = @(
	"offline_access",
	"openid",
	"profile"
)
$appSPAppPerms = @(
	"Sites.FullControl.All",
	"User.ReadWrite.All",
	"TermStore.ReadWrite.All"
)
$appSPDelegPerms = @(
	"AllSites.FullControl"
)
$tenantDomain = ("{0}.onmicrosoft.com" -f $clientTenant)
$certOutputFolder = "..\_Output Files\"

# Output Parameters & Variables
Write-Host
Write-Host "Script Parameters" -ForegroundColor Cyan
Write-Host ("   App Registration: {0}" -f $appName)
Write-Host ("   Client Tenant Domain: {0}" -f $clientTenant)
Write-Host ("   Certificate Common Name: {0}" -f $certCommonName)
Write-Host ("   Certificate Store: {0}" -f $certStore)
Write-Host ("   Certificate Country: {0}" -f $certCountry)
Write-Host ("   Certificate State: {0}" -f $certState)
Write-Host ("   Certificate City: {0}" -f $certCity)
Write-Host ("   Certificate Org: {0}" -f $certOrg)
Write-Host ("   Certificate Department: {0}" -f $certDept)
Write-Host ("   Certificate Valid for (years): {0}" -f $certValidYears)
Write-Host
Write-Host "User Inputs" -ForegroundColor Cyan
Write-Host ("   Certificate Password: {0}" -f $certPass)
Write-Host
Write-Host "Script Variables" -ForegroundColor Cyan
Write-Host ("   Microsoft Tenant: {0}" -f $tenantDomain)
Write-Host ("   Certificate Output Folder: {0}" -f $certOutputFolder)
Write-Host
Write-Host "MS Graph Application Permissions" -ForegroundColor Cyan
Write-Host $appGraphAppPerms
Write-Host
Write-Host "MS Graph Delegated Permissions" -ForegroundColor Cyan
Write-Host $appGraphDelegPerms
Write-Host
Write-Host "SharePoint Application Permissions" -ForegroundColor Cyan
Write-Host $appSPAppPerms
Write-Host
Write-Host "SharePoint Delegated Permissions" -ForegroundColor Cyan
Write-Host $appSPDelegPerms
#endregion

#region 2.	Snapin(s), Modules & Environment Connection(s)
# Load MS / PnP PowerShell Modules
Get-Module -ListAvailable | Where-Object{($_.Name -like "*SharePoint*") -or ($_.Name -like "*PnP*") } | Remove-Module
Import-Module PnP.PowerShell
#endregion
    
#region 3.	Main Script
# Creating the App Registration
Write-Host
Write-Host ("Creating App Registration: {0}" -f $appName)
Write-Host
$appReg = Register-PnPAzureADApp -ApplicationName $appName `
    -Tenant $tenantDomain `
    -Interactive `
    -CommonName $certCommonName `
	-CertificatePassword $certPass `
    -OutPath $certOutputFolder `
    -Store $certStore `
	-Country $certCountry `
	-State $certState `
	-Locality $certCity `
	-Organization $certOrg `
	-OrganizationUnit $certDept `
	-ValidYears $certValidYears `
	-GraphApplicationPermissions $appGraphAppPerms `
	-GraphDelegatePermissions $appGraphDelegPerms `
	-SharePointApplicationPermissions $appSPAppPerms `
	-SharePointDelegatePermissions $appSPDelegPerms #>

# Output App Registration Object
if ($null -ne $appReg) {
	Write-Host
	Write-Host ("App Registration Successfully Created: {0}" -f $appName) -ForegroundColor Green
	Format-List -InputObject $appReg
}
#endregion

#region Completed Script
# Output Script Info
Complete-CatScript $scriptStart[0]

# Stop logging
Stop-Transcript
#endregion