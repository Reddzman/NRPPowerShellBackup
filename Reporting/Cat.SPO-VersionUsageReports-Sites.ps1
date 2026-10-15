<#  Cat.SPO-VersionUsageReports-Sites.ps1

#This PowerShell script automates running SharePoint Version Usage Rerports for a list of sites 
with a polling loop that ensures prior async report progress is completed before starting the next loop. 
The  Microsoft timer job restriction and only 1 version usage report can be running at a time. 
No parallel option available, jobs must be sequential.

Dependencies
    -   Service Account for SharePoint
    -	PowerShell version 7
    -   Microsoft.Online.SharePoint.PowerShell
    -	..\General\Cat.GeneralFunctions.psm1
    -   ..\Reporting\Cat.ReportingFunctions.psm1
    -	Input CSV File in Input Files folder
    -	Output Reports File in SPO library folder in client domain tenant

#region 1.	Snapin(s), Modules & Environment Connection(s)
# Load modules
#Import-Module Microsoft.Online.SharePoint.PowerShell
#Note in PS 7 must be loaded before calling script with specific parameters for user and/or machine 
#(i.e. Import-Module "Import-Module "C:\Program Files\PowerShell\Modules\Microsoft.Online.SharePoint.PowerShell\16.0.27215.12000\Microsoft.Online.SharePoint.PowerShell.psd1" -UseWindowsPowerShell -Force)

Call script with parameters 
    cd C:\Quisitive\PowerShell\Reporting
    .\Cat.SPO-VersionUsageReports-Sites.ps1 -jobFile "C:\Users\Isabel Liu\OneDrive - Quisitive\Clients\Crocs\Scripts\Top16-22sites.csv" -destTenant "crocsconnect"

CSV Example
    

Table of Contents
	1.	Snapin(s), Modules & Environment Connection(s)
	2.	Input & Output Files
	3.	Variables
    4.  Functions
    5.	Main Script
    6.  Running Microsoft SharePoint Version Usage Reports
#>

#region Script Parameters
Param(
    [string]$jobFile = "Cat.SiteVersionExpiration.csv",
    [string] $destTenant = "nrpdev2"
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
Start-Transcript -Path ("..\_Log Files\{0}.SPO-VersionUsageReports-{1}.log" -f $destTenant, $scriptStart[1])

#region 1.	Snapin(s), Modules & Environment Connection(s)
# Load modules
#Import-Module Microsoft.Online.SharePoint.PowerShell
#Note in PS 7 must be loaded before calling script with specific parameters for user and/or machine 
#(i.e. Import-Module "Import-Module "C:\Program Files\PowerShell\Modules\Microsoft.Online.SharePoint.PowerShell\16.0.27215.12000\Microsoft.Online.SharePoint.PowerShell.psd1" -UseWindowsPowerShell -Force)
# OR
Import-Module Microsoft.Online.SharePoint.PowerShell

#ByPass Execution Policy
# Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force

# Load Custom Modules
Import-Module "..\Reporting\Cat.ReportingFunctions.psm1"
#endregion

#region 2.	Input & Output Files
# File Paths
$inPath = ("..\_Input Files\{0}" -f $jobFile)

Write-Host ("   inPath: {0}" -f $inPath)

# Import Input File(s)
# Load CSV file
If (Test-Path ($inPath)){
	$inFile = Import-Csv $inPath -Encoding Unicode
    #$inFile # uncomment if inFile is empty to see what it has
}
else {
	Write-Host
	throw ("The jobs file does not exist: {0}" -f $inPath)	
}
#endregion

#region 3.	Variables
# Script Variables
$destAdminUrl = ("https://{0}-admin.sharepoint.com/" -f $destTenant)
$destRootSiteUrl = ("https://{0}.sharepoint.com/" -f $destTenant)
$destTenantDomain = ("{0}.onmicrosoft.com" -f $destTenant)
$specialCharacters = "[~``!@#$%^&*():_]"

# User Inputs
Write-Host
Write-Host "Connect to Tenant" -ForegroundColor Cyan
Connect-SPOService -Url $destAdminUrl #-UseSystemBrowser $true
 
# Variable Output
Write-Host
Write-Host "Script Parameters" -ForegroundColor Cyan
Write-Host ("   Job File: {0}" -f $jobFile)
Write-Host ("   Destination Tenant: {0}" -f $destTenant)
Write-Host
Write-Host "Script Variables" -ForegroundColor Cyan
Write-Host ("   Destination Admin Url: {0}" -f $destAdminUrl)
Write-Host ("   Destination Root Site Url: {0}" -f $destRootSiteUrl)
Write-Host ("   Destination Tenant Url: {0}" -f $destTenantDomain)
Write-Host ("   Special Characters {0}" -f $specialCharacters)
#endregion

#region 5.	Main Script
Write-Host
Write-Host ("Configuring & running reports for job list: {0}" -f $jobFile) -ForegroundColor Cyan

#$infile = Import-Csv -Path $jobFile
#test run check every 2 minutes (select small sites for testing), uncomment line below for test runs, comment line below for live runs
#$pollSeconds = 120 
#live run check every 15 minutes (adjust time polling period based on site sizes of sites being run, longer for large sites), comment line below for test runs, uncomment line below for live runs
$pollSeconds = 900 

:JobLoop foreach ($job in $inFile) {
try {
    #get sitecollection url substring if subsite
    $n = 5 # The nth instance
    $reportSiteUrl = $job.URL
    $slashCount = ([regex]::Matches($reportSiteUrl, "/")).Count

    if ($slashCount -ge $n){
        $index = [regex]::Matches($job.URL, "/")[$n-1].Index

        $reportSiteUrl = $job.URL.Substring(0, $index) 
    }

    $reportStartTimeStamp = Get-Date -Format "yyyyMMddHHmmss"
    $reportUrlFilePath = ("{0}/Shared Documents/_ITUSEONLY-DONOTDELETE-VersionUsageReport-{1}-{2}.csv" -f $reportSiteUrl, $job.Title,$reportStartTimeStamp)
    Write-Host ("   Site Url: {0}" -f $job.URL) -ForegroundColor Blue
    Write-Host ("   Report Url File Path: {0}" -f $reportUrlFilePath) -ForegroundColor Blue
    New-SPOSiteFileVersionExpirationReportJob -Identity $job.URL -ReportUrl $reportUrlFilePath -Confirm:$false

    do {
        $statusTimeStamp = Get-Date -Format "yy-MM-dd HH:mm:ss"
        Write-Host
        Write-Host ("Waiting to check status at {0} of job report: {1}" -f $statusTimeStamp, $job.Title) -ForegroundColor Cyan
    
        Start-Sleep -Seconds $pollSeconds

        $reportStatusObject = Get-SPOSiteFileVersionExpirationReportJobProgress -Identity $job.URL -ReportUrl $reportUrlFilePath
        $reportStatus = $reportStatusObject.Status
        $statusTimeStamp = Get-Date -Format "yy-MM-dd HH:mm:ss"

        Write-Host ("         Status at {0}: {1}" -f $statusTimeStamp, $reportStatus) -ForegroundColor Gray

    } while ($reportStatus -eq "in_progress")

    # Final status will be: completed or failed
    $reportStatusObject
    If ($reportStatus -eq "completed"){
        Write-Host ("SUCCESS - Completed at {0}: {1}" -f $statusTimeStamp, $job.Title) -ForegroundColor Green
    }
    else {
        Write-Host ("FAIL Status at {0}: {1} - {2}" -f $statusTimeStamp, $job.Title, $reportStatus) -ForegroundColor Red
    }

}
catch {
    Write-Host ("FAIL: {0} - {1}" -f $job.Title, $_.Exception.Message) -ForegroundColor Red

    # Output Error Information
    Write-Host ("         {0} - Error running Version Usage Report-Exception: {1}" -f $statusTimeStamp, $_.Exception ) -ForegroundColor Red
    Write-Host ("         ----- Error running Version Usage Report-Error Details: {0}" -f $_.ErrorDetails) -ForegroundColor Red
    Write-Host ("         ----- Error running Version Usage Report-Stack Tracke: {0}" -f $_.ScriptStackTrace) -ForegroundColor Red
    break JobLoop
}

}

# Script complete output
Write-Host
Write-Host "Version Usage Report runs completed. " -ForegroundColor Green
Write-Host "Please refer to the job log for warnings/errors that require reruns" -ForegroundColor Yellow
#endregion

#region Completed Script
# Output Script Info
Complete-CatScript $scriptStart[0]

# Stop logging
Stop-Transcript

#endregion



