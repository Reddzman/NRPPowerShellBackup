<#  Cat.SkySync-PauseJobWaves.ps1

This script pauses all jobs that meet search criteria.

Dependencies
	-	SkySync CLI
	-	..\General\Cat.PnP-GeneralFunctions.psm1
	-	.\Cat.SkySyncFunctions.psm1

Call Script with Input File
	.\Cat.SkySync-PauseJobWave.ps1 -searchCriteria "TestWave" -offsetTotal 5000

Table of Contents
	1.	Variables
	2.	Snapin(s), Modules & Environment Connection(s)
	3.	Main Script
#>

#region Script Parameters
Param(  
    [String] $JobID = "TestWave-TestBU-TestDepartment",
    [String] $offsetTotal = 999,
    [array] $csvFile = "..\_Input Files\JobsToPause.csv"
)
#endregion

#region Start Date / Time
# Remove Existing Catapult Modules to ensure PS session has most up to date functions
Get-Module Cat.* | Remove-Module

# Load General Functions
Import-Module "..\General\Cat.GeneralFunctions.psm1"

$csv = Import-Csv -Path $csvFile
$CSVJobIDs = $csv.JobID

# Get Start Date / Time
$scriptStart = Start-CatScript

# Start Transcript
Start-Transcript -Path ("..\_Log Files\PauseJobCSV-{0}.log" -f $scriptStart[1])
#endregion

#region 1.	Variables
# Script Variables
$accessToken = skysync-cli sessions token --output-json false
$authHeader = @{"Authorization"="Bearer $accessToken";}
$offset = 0
$serverURL = "http://localhost:9090/"
$baseURL = ("{0}v1" -f $serverURL)
$contentType = 'application/json'

# Output Parameters & Variables
Write-Host
Write-Host "Script Parameters" -ForegroundColor Cyan
Write-Host ("   JobID: {0}" -f $JobID)
Write-Host ("   Offset Total: {0}" -f $offsetTotal)
Write-Host
Write-Host "Script Variables" -ForegroundColor Cyan
Write-Host ("   Offset: {0}" -f $offset)
Write-Host ("   Server URL: {0}" -f $serverURL)
Write-Host ("   Base URL: {0}" -f $baseURL)
Write-Host ("   Content Type: {0}" -f $contentType)
#endregion

#region 2.	Snapin(s), Modules & Environment Connection(s)
# Load Custom Modules
Import-Module "..\SkySync\Cat.SkySyncFunctions.psm1"
#endregion

#region 3.	Main Script
foreach($JobID in $CSVJobIDs)
{
    Write-Host
    do {
        # Get Jobs
        $pauseJobs = skysync-cli jobs pause $JobID --limit 1000 --offset $offset --json true | ConvertFrom-Json
        If ($pauseJobs.count -gt 0) {
            Write-Host ("   Job: {0} | have been paused." -f $JobID, $offset)
        }
        Else {
            Write-Host ("   Job: {0} | FAILED to pause." -f $JobID, $offset) -ForegroundColor Yellow
        }

        # Incrementing Offset for paging
        $offset += 1000
    }
    while ($offset -lt $offsetTotal)
}
#endregion

#region Completed Screen Output
Complete-CatScript $scriptStart[0]
Stop-Transcript
#endregion