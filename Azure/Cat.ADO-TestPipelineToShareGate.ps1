<#  Cat.ADO-TestPipelineToShareGate.ps1

This script is designed to test script connections to the ShareGate migration tool using an ADO Repo and Pipeline to execute the script. Should be executed on a ShareGate migration server to confirm capability.
ShareGate is licensed to a specific Machine user. Powershell will also need to be executed as the licensed ShareGate Service account.

Dependencies
    -   ShareGate PowerShell Modules

Call script with parameters
    cd C:\GitHub\CatPowerShell\Azure\
    $source = Get-Credential # Domain\user & Password
    $dest = Get-Credential # M365 UPN & Password
	.\Cat.ADO-TestPipelineToShareGate.ps1 -client "nrpdev" -sourceURL "https://portal.sp13test.local/sites/migration" -destURL "https://nrpdev2.sharepoint.com/sites/OOtB-ModernCommunicationSite" -sourceCred $source -destCred $dest

Table of Contents
	1.	Variables
	2.	Snapin(s), Modules & Environment Connection(s)
    3.	Main Script
#>

#region Script Parameters
Param(
    [String] $client = "nrpdev",
    [String] $sourceURL = "https://portal.sp13test.local/sites/migration",
    [string] $destURL = "https://nrpdev2.sharepoint.com/sites/OOtB-ModernCommunicationSite",
    [pscredential] $sourceCred,
    [pscredential] $destCred
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
Start-Transcript -Path ("..\_Log Files\{0}.ADO-TestPipelineToShareGate-{1}.log" -f $client, $scriptStart[1])
#endregion

#region 1.	Variables
# Variable Output
Write-Host
Write-Host "Script Parameters" -ForegroundColor Cyan
Write-Host ("   Client: {0}" -f $client)
Write-Host ("   Source URL: {0}" -f $sourceURL)
Write-Host ("   Destination URL: {0}" -f $destURL)
Write-Host ("   Source Cred: {0}" -f $sourceCred.UserName)
Write-Host ("   Destination Cred: {0}" -f $destCred)
#endregion

#region 2.	Snapin(s), Modules & Environment Connection(s)
# Load Sharegate modules
Import-Module Sharegate	
#endregion
    
#region 3.	Main Script
# Testing Connection with ShareGate
Write-Host
Write-Host "Testing Connection with ShareGate:" -ForegroundColor Cyan

# Connecting to source site
Write-Host
Write-Host ("   Connecting to source site: {0}" -f $sourceURL) -ForegroundColor Magenta
$sgSourceSite = Connect-Site -Url $sourceURL -Credential $sourceCred
Write-Host ("      Source Site GUID: {0}" -f $sgSourceSite.Id)

# Connecting to source site
Write-Host
Write-Host ("   Connecting to destination site: {0}" -f $destURL) -ForegroundColor Magenta
$sgDestSite = Connect-Site -Url $destURL -Credential $destCred  
Write-Host ("      Destination Site GUID: {0}" -f $sgDestSite.Id) 
#endregion

#region Completed Script
# Output Script Info
Complete-CatScript $scriptStart[0]

# Stop logging
Stop-Transcript
#endregion