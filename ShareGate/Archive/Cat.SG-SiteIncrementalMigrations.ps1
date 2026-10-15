<#  Cat.SG-SiteIncrementalMigrations.ps1

This script automates ShareGate migrations incremental list migrations at a Site Collection level.  Incremental migrations will be ran
for each Site collection in the CSV File. The script iterates through all webs in a site collection and migrates all lists per web.

Dependencies
    -	SharePoint 2013 Module
    -   SharePoint PnP PowerShell 2013 (installed on machine, usable in ShareGate Shell)
	-	ShareGate PowerShell Snapins
    -	..\General\Cat.PnP-GeneralFunctionsFunctions.psm1
    -   ..\Configuration\Cat.ConfigurationFunctions.psm1
    -	Input CSV File in Input Files folder - Script Parameter
    -   User mapping file export from ShareGate (UserMappingFile.sgum)

Call script with parameters
	.\Cat.SG-SiteMigrations.ps1 -inputFile Cat.SG-SiteList-Test.csv -userFile Cat.SG-UserMapping-Test.sgum

CSV Example
	sourceURL,destinationURL,siteTitle,siteOwner,secondOwner
    https://source.client.com/sites/testsite01,https://sharepoint.client.com/sites/testsite01,Migration Test 01,Client\user,client\sp_install

Table of Contents
	1.	Input & Output Files
	2.	Variables
	3.	Snapin(s), Modules & Environment Connection(s)
    4.	Main Script
    5.  Run Migration
#>

#region Script Parameters
Param(
        [String] $inputFile = "Cat.SG-SiteList-Test.csv",
        [String] $userFile = "Cat.SG-UserMapping-Test.sgum",
        [String] $srcURL = "https://sharepoint-old.client.com",
        [String] $dstURL = "https://sharepoint.client.com"
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
Start-Transcript -Path ("..\_Log Files\Cat.SG-SiteFullMigrations-{0}.log" -f $scriptStart[1])
#endregion

#region 1.	Input & Output Files
# File Paths
$inPath = ("..\_Input Files\{0}" -f $inputFile)
$userPath = ("..\_Input FIles\{0}" -f $userFile)

# Import Input File(s)
# Load CSV file
If (Test-Path ($inPath)){
	$inFile = Import-Csv $inPath -Encoding Unicode
}
else {
	Write-Host
	throw ("The jobs file does not exist: {0}" -f $inPath)	
}

# Load user mapping file
If (Test-Path ($userPath)){
	$userPath = Resolve-Path $userPath
}
else {
	Write-Host
	throw ("The jobs file does not exist: {0}" -f $userPath)	
}
#endregion

#region 2.	Variables
# User Inputs
Write-Host
Write-Host "User Input Variables" -ForegroundColor Cyan
Write-Host "   Enter credentials for the source SharePoint environment:"
$spCred = Get-Credential -Message "Enter your credentials for the source SharePoint environment:" 

# Output Parameters & Variables
Write-Host
Write-Host "Script Parameters" -ForegroundColor Cyan
Write-Host ("   Input File: {0}" -f $inputFile)
Write-Host ("   User Mapping File: {0}" -f $userFile)
Write-Host
Write-Host "User Inputs" -ForegroundColor Cyan
Write-Host ("   Credentials: {0}" -f $spCred.UserName)
#endregion

#region 3.	Snapin(s), Modules & Environment Connection(s)
# Load MS / PnP PowerShell Modules
Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue

# Load Sharegate modules
Import-Module Sharegate	

# Load Custom Modules
Import-Module "..\Reporting\Cat.ReportingFunctions.psm1"
#endregion
    
#region 4.	Main Script
Write-Host
Write-Host ("Starting Migration Jobs: {0}" -f $inputFile) -ForegroundColor Cyan

# Looping through input file
ForEach ($job in $inFile){
    Write-Host
    Write-Host ("   Migration Job: {0}" -f $job.sourceURL) -ForegroundColor Magenta

    #region 5.  Run Migrations    
    # Get source site collection
    Write-Host ("      Loading Source Site: {0}" -f $job.sourceURL)
    $siteWebs = Get-CatPnPAllWebs -Creds $spCred -URL $job.sourceURL

    # Loop Through All Webs in site collection
    ForEach ($curWeb in $siteWebs){
        # Determine if site exists
        $siteExists = Test-CatSP10SiteExists -url $curWeb.Url.Replace($srcURL,$dstURL)  -type "Web"
        If ($siteExists){
            # Setting up Job configuration
            Write-Host ("      Loading Source Web: {0}" -f $curWeb.Url)
            $srcSite = Connect-Site -Url $curWeb.Url -Username $spCred.UserName -Password $spCred.Password
            Write-Host ("         Loading Destination Web: {0}" -f $curWeb.Url.Replace($srcURL,$dstURL))
            $dstSite = Connect-Site -Url $curWeb.Url.Replace($srcURL,$dstURL)
            Write-Host ("         Loading Mappings file: {0}" -f $userFile)
            $copyMappings = New-MappingSettings
            $copyMappings = Import-UserAndGroupMapping -MappingSettings $copyMappings -Path $userPath.Path
            Write-Host "         Creating Copy Settings"
            $copySettings = New-CopySettings -OnContentItemExists IncrementalUpdate
            $taskName = ("{0} - {1} - Inc - {2}" -f $job.siteTitle, $curWeb.Title, (Get-Date -Format "MM/dd/yy HH:mm"))

            # Running ShareGate Migration Job
            Write-Host ("         Running Job: {0}" -f $taskName)
            Copy-List -ALL -SourceSite $srcSite -DestinationSite $dstSite -TaskName $taskName -CopySettings $copySettings -MappingSettings $copyMappings

            # Log complete
            Write-Host ("         Completed Successfully: {0}" -f $curWeb.Url) -foregroundcolor "green"   
        }
        Else {
            Write-Host ("      Site doesn't exist in destination: {0}" -f $curWeb.Url.Replace($srcURL,$dstURL)) -ForegroundColor Red
        }
    }
    #endregion #>
}
#endregion

#region Completed Script
# Output Script Info
Complete-CatScript $scriptStart[0]

# Stop logging
Stop-Transcript
#endregion