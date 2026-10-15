<#  Q.SG-AddSiteCollectionsToExplorer.ps1

This script iterates through a CSV list of sites and adds them to ShareGate's Explorer. 

*** This script was validated with ShareGate version 24.9.0

Dependencies
	-	ShareGate Module (must run on ShareGate server)
	-	Input CSV File with list of site collection URLs to add under the heading "URL"

Call script with parameters
	cd C:\GitHub\CatPowerShell\Configuration
	.\Cat.SG-AddSiteCollectionsToExplorer.ps1 -inputFile "Inside Sanford SP Master Inventory_URLs.csv"

CSV Notes:
	URL - URL of Site Collection to add

Table of Contents
	1.	Snapin(s), Modules & Environment Connection(s)
	2.	Input & Output Files
	3.	Variables
	4.	Main Script
#>

#region Script Parameters
Param(	
     [string] $inputFile = "Q.SG-AddSiteCollectionsToExplorer-Client.csv"
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
Start-Transcript -Path ("..\_Log Files\{0}.SG-AddSiteCollectionsToExplorer-{1}.log" -f $client, $scriptStart[1])
#endregion

#region 1.	Snapin(s), Modules & Environment Connection(s)
# Load Modules
Import-Module Sharegate
#endregion

#region 2.	Input & Output Files
# File Paths
$inPath = ("..\_Input Files\{0}" -f $inputFile)
$errorPath = ("..\_Output Files\{0}.SG-AddSiteCollectionsToExplorer-Errors-{1}.log" -f $client, $scriptStart[1])

# Import Input File(s)
# Load CSV file
If (Test-Path ($inPath)){
	$inFile = Import-Csv $inPath -Encoding Unicode
}
else {
	Write-Host
	throw ("The input file does not exist: {0}" -f $inPath)	
}
#endregion

#region 3.	Variables
# User Inputs
Write-Host
Write-Host "User Input Variables" -ForegroundColor Cyan
$userName = Read-Host -Prompt "   Enter Username"
$password = ConvertTo-SecureString -String (Read-Host -Prompt "   Enter password") -AsPlainText -Force

# Output Parameters & Variables
Write-Host
Write-Host "Script Parameters" -ForegroundColor Cyan
Write-Host ("   Input File: {0}" -f $inputFile)
#endregion

#region 4.	Main Script
:SiteLoop ForEach ($site in $inFile){
     try{
          Connect-Site -Url $site.URL -Username $userName -Password $password -AddToExplorer
     }
     catch {
          # Output Error Information
          Log-MigrationError -funcJob $job -funcAction "Adding Site to ShareGate Explorer" -funcException $_.Exception -funcDetails $_.ErrorDetails -funcStackTrace $_.ScriptStackTrace -funcLogPath $errorPath
          Continue SiteLoop
      }
     
}
#endregion

#region Completed Script
# Output Script Info
Complete-CatScript $scriptStart[0]

# Stop logging
Stop-Transcript
#endregion