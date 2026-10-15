<#  Cat.SPO-ExportLists.ps1

This script exports the contents of the specified lists for each site in in the input CSV

Dependencies
	-	Source & Destination Service Account for ShareGate
    -   ..\General\Cat.GeneralFunctions.psm1
    -   ..\Reporting\Cat.ReportingFunctions.psm1
    -	Input CSV File in Input Files folder - Script Parameter

Call script with parameters
	.\Cat.SPO-ExportLists.ps1 -jobsFile "Cat.SPO-ExportLists-TEST.csv" -tenant "ClientTenant" -appID "Azure App ID" -certPath "C:\certs\Certificate.pfx"

CSV Example
	siteUrl,listName
    https://clienttenant.sharepoint.com/sites/TestSite,User Information List

Table of Contents
	1.	Input & Output Files
	2.	Variables
	3.	Snapin(s), Modules & Environment Connection(s)
    4.	Main Script
#>

#region Script Parameters
Param(
    [String] $jobsFile = "Cat.SPO-ListExport-Test.csv",
    [string] $tenant = "{SourceTenant}",
    [String] $appID = "57577043-3de0-40c0-9c68-7960e2cebe8b",
    [String] $certPath = "C:\certs\PnP PowerShell Authentication.pfx"
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
Start-Transcript -Path ("..\_Log Files\{0}.SPO-ExportUserInformationList-{0}.log" -f $destTenant, $scriptStart[1])
#endregion

#region 1.	Input & Output Files
# File Paths
$inPath = ("..\_Input Files\{0}" -f $jobsFile)

# Import Input File(s)
# Load CSV file
If (Test-Path ($inPath)){
	$inFile = Import-Csv $inPath -Encoding Unicode
}
else {
	Write-Host
	throw ("The jobs file does not exist: {0}" -f $inPath)	
}

#endregion

#region 2.	Variables
Write-Host "Enter Connection Information" -ForegroundColor Cyan
$certPass = ConvertTo-SecureString -String (Read-Host -Prompt "   Enter SOURCE certificate password for the app registration") -AsPlainText -Force
$fileListExport = ("..\_Output Files\{0}-UserInformationList-{1}.txt" -f$Tenant, $scriptStart[1])
Out-File -FilePath $fileListExport -InputObject "SiteURL|ListName|User"

$matchFilter = "Charlesmachine.works#EXT#"
# Script Variables
$tenantDomain = ("{0}.onmicrosoft.com" -f $tenant) 

# Output Parameters & Variables
Write-Host
Write-Host "Script Parameters" -ForegroundColor Cyan
Write-Host ("   Jobs File: {0}" -f $jobsFile)
Write-Host ("   Tenant: {0}" -f $tenant)
Write-Host
Write-Host "Script Variables" -ForegroundColor Cyan
Write-Host ("   Tenant Domain: {0}" -f $tenantDomain)
#endregion

#region 3.	Snapin(s), Modules & Environment Connection(s)
# Load Custom Modules
Import-Module "..\Reporting\Cat.ReportingFunctions.psm1"
#endregion
  
#region 4.	Main Script
Write-Host
Write-Host ("Exporting Lists: {0}" -f $jobsFile) -ForegroundColor Cyan

# Looping through input file
ForEach ($job in $inFile){
    $siteUrl = $job.siteUrl
    $listName = $job.listName

    Write-Host
    Write-Host ("   Connecting to SharePoint site: {0}" -f $siteUrl) 
    Connect-PnPOnline -ClientId $appID -Url $siteUrl -Tenant $tenantDomain -CertificatePath $certPath -CertificatePassword $certPass -ErrorAction SilentlyContinue
    
    Write-Host ("   Exporting List: {0}" -f $listName) -ForegroundColor Magenta
    $listItems = Get-PnPListItem -List $listName
    foreach($listItem in $listItems)  
    { 
        if ($listItem["Name"] -match $matchFilter) {
            Out-File -FilePath $fileListExport -Append -InputObject ("{0}|{1}|{2}" -f 
            $siteUrl, 
            $listName,
            $listItem["Name"])    
        }
    }

    Disconnect-PnPOnline
}

# Complete Output
Write-Host
Write-Host ("Lists have been exported to:{0}" -f $fileListExport) -ForegroundColor Cyan
#endregion

#region Completed Script
# Output Script Info
Complete-CatScript $scriptStart[0]

# Stop logging
Stop-Transcript
#endregion