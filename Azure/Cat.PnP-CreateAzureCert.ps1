<#
Cat.PnP-CreateAzureCert.ps1

This script this script create self signed certificate for use in PnP Scripts.  This is a prerequisit to App Authentication.

Dependencies
	-	PnP PowerShell https://docs.microsoft.com/en-us/powershell/sharepoint/sharepoint-pnp/sharepoint-pnp-cmdlets?view=sharepoint-ps

Call Script with Input File
	.\Cat.PnP-CreateAzureCert.ps1 -commonName "Test Name" -fileName "filename.pfx" -etc.

Table of Contents
	1.	Input & Output Files
	2.	Main Script
#>

#region Script Parameters
Param(  [String] $commonName = "NRPDevSiteProvisioning",
        [String] $fileName = "NRPDevSiteProvisioning.pfx",
        [string] $filePass = "",
        [String] $orgCountry = "US",
        [String] $orgState = "Colorado",
        [String] $orgCity = "Denver",
        [String] $orgName = "Catapult",
        [String] $orgUnit = "Consulting",
        [int] $validYears = 10
)
#endregion

#region Start Date / Time
# Set PS locPath
Set-Location $PSScriptRoot

# Get Script Parent locPath
$localSplit = $PSScriptRoot.Split("\")
$localReplace = $localSplit[($localSplit.Count - 1)]
$locPath = $PSScriptRoot.Replace($localReplace,"")

# Remove Existing Catapult Modules to ensure PS session has most up to date functions
Get-Module Cat.* | Remove-Module

# Load General Functions
Import-Module ("{0}General\Cat.GeneralFunctions.psm1" -f $locPath)

# Get Start Date / Time
$scriptStart = Start-CatScript
#endregion

#region 1.	Input & Output Files
# File Paths
$outPath = ("{0}_Output Files\{1}" -f $locPath,$fileName)
#endregion

#region 2. Variables
$fileSecure = ConvertTo-SecureString $filePass -AsPlainText -Force
#endregion

#region Output Parameters & Variables
Write-Host
Write-Host "Script Parameters" -ForegroundColor Cyan
Write-Host ("   Common Name: {0}" -f $commonName)
Write-Host ("   File Name: {0}" -f $fileName)
Write-Host ("   Password: {0}" -f $filePass)
Write-Host ("   Country: {0}" -f $orgCountry)
Write-Host ("   State: {0}" -f $orgState)
Write-Host ("   City: {0}" -f $localCity)
Write-Host ("   Organization Name: {0}" -f $orgName)
Write-Host ("   Organization Unit: {0}" -f $orgUnit)
Write-Host ("   Years Valid: {0}" -f $validYears)
#endregion

#region 4.	Main Script
Write-Host
Write-Host "Creating Azure Certificate" -ForegroundColor Magenta
# Generate Certificate
New-PnPAzureCertificate -CommonName $commonName -CertificatePassword $fileSecure -ValidYears $validYears -Out $outPath -Country $orgCountry -State $orgState -Locality $orgCity -Organization $orgName -OrganizationUnit $orgUnit
#endregion

#region Completed Screen Output
Complete-CatScript $scriptStart[0]
#endregion