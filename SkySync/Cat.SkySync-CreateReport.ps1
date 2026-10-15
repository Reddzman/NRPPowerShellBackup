<#  Cat.SkySync-CreateReports.ps1

This script creates a report with all jobs that meet search criteria.

Dependencies
	-	SkySync CLI
	-	..\General\Cat.PnP-GeneralFunctions.psm1

Call Script with Input File
	.\Cat.SkySync-CreateSPOConnections.ps1 -searchCriteria "TestWave-TestBU" -reportName "TestWave-TestBU"

Table of Contents
	1.	Variables
	2.	Main Script
#>

#region Script Parameters
Param(  
    [String] $searchCriteria = "TestWave-TestBU",
    [String] $reportName = "TestWave-TestBU",
    [string] $reportType = "Simulation",
    [String] $numJobs = 1000
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
Start-Transcript -Path ("..\_Log FIles\CreateReport-{0}.log" -f $scriptStart[1])
#endregion

#region 1.	Variables
# Script Variables
$offset = 0

# Output Parameters & Variables
Write-Host
Write-Host "Script Parameters" -ForegroundColor Cyan
Write-Host ("   Search Criteria: {0}" -f $searchCriteria)
Write-Host ("   Report Name: {0}" -f $reportName)
Write-Host ("   Offset Total: {0}" -f $numJobs)
Write-Host
Write-Host "Script Variables" -ForegroundColor Cyan
Write-Host ("   Offset: {0}" -f $offset)
#endregion

#region 2.	Main Script
Write-Host
Write-host ("Creating Report: {0}" -f $reportName) -ForegroundColor Magenta

# Get Jobs
$allJobs = @()
$counter = 1
If ($numJobs -le 242) {
    Write-Host ("   Search for Jobs: {0}" -f $searchCriteria)
    $reportJobs = skysync-cli jobs list --search $searchCriteria --limit 242 --offset $offset --json true | ConvertFrom-Json
    
    # Creating Report
    Write-Host ("   Adding Jobs to Report: {0}" -f $reportName)
    If ($reportType -eq "Simulation"){
        $reportTemp = skysync-cli reports add --name ("{0} - {1}" -f $reportName, $counter) --simulation --jobs $reportJobs.ID
    }
    else {
        $reportTemp = skysync-cli reports add --name ("{0} - {1}" -f $reportName, $counter) --jobs $reportJobs.ID
    }
}
Else {
    do {
        Write-Host ("   Search for Jobs: {0}   Offset: {1}" -f $searchCriteria, $offset) -ForegroundColor DarkGray
        $reportJobs = skysync-cli jobs list --search $searchCriteria --limit 242 --offset $offset --json true | ConvertFrom-Json
        
        # Creating Report
        Write-Host ("      Adding Jobs to Report: {0}" -f $reportName)
        If ($reportType -eq "Simulation"){
            $reportTemp = skysync-cli reports add --name ("{0} - {1}" -f $reportName, $counter) --simulation --jobs $reportJobs.ID
        }
        else {
            $reportTemp = skysync-cli reports add --name ("{0} - {1}" -f $reportName, $counter) --jobs $reportJobs.ID
        }
    
        # Incrementing Offset for paging
        $offset += 242
        $counter++
    }
    while ($offset -lt $numJobs)
}
#endregion

#region Completed Screen Output
Complete-CatScript $scriptStart[0]
Stop-Transcript
#endregion