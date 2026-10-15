<#  Cat.SkySync-StartJobWaves.ps1

This script disables simulation mode, sets schedule, & starts all jobs that meet search criteria.

Dependencies
	-	SkySync CLI
	-	..\General\Cat.PnP-GeneralFunctions.psm1
	-	.\Cat.SkySyncFunctions.psm1
	-	Input CSV Files located in _Input Files
		-	Jobs File
		-	Template File

Call Script with Input File
	.\Cat.SkySync-StartJobWave.ps1 -searchCriteria "TestWave-TestBU-TestDepartment"

Table of Contents
	1.	Variables
	2.	Snapin(s), Modules & Environment Connection(s)
	3.	Main Script
#>

#region Script Parameters
Param(  
    [String] $searchCriteria = "TestWave-TestBU",
    [Int] $offsetTotal = 999
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
Start-Transcript -Path ("..\_Log FIles\StartJobWaves-{0}.log" -f $scriptStart[1])
#endregion

#region 1.	Variables
# Script Variables
$accessToken = skysync-cli sessions token --output-json false
$authHeader = @{"Authorization"="Bearer $accessToken";}
$offset = 0
$scheduleJSON = '{
    "kind": "transfer",
    "schedule": {
        "mode": "auto",
        "end_date": 1577775600,
        "repeat_interval": {
            "value": 900000,
            "unit": "ms"
        },
        "days": ["everyday"]
    }
}'
$serverURL = "http://localhost:9090/"
$baseURL = ("{0}v1" -f $serverURL)
$contentType = 'application/json'

# Output Parameters & Variables
Write-Host
Write-Host "Script Parameters" -ForegroundColor Cyan
Write-Host ("   Search Criteria: {0}" -f $searchCriteria)
Write-Host ("   Offset Total: {0}" -f $offsetTotal)
Write-Host
Write-Host "Script Variables" -ForegroundColor Cyan
Write-Host ("   Offset: {0}" -f $offset)
Write-Host ("   Server URL: {0}" -f $serverURL)
Write-Host ("   Base URL: {0}" -f $baseURL)
Write-Host ("   Content Type: {0}" -f $contentType)
<#Write-Host
Write-Host "Schedule JSON" -ForegroundColor Cyan 
Write-host $scheduleJSON #>
#endregion

#region 2.	Snapin(s), Modules & Environment Connection(s)
# Load Custom Modules
Import-Module "..\SkySync\Cat.SkySyncFunctions.psm1"
#endregion

#region 3.	Main Script
Write-Host
Write-host ("Configuring & Starting Jobs: {0}" -f $searchCriteria) -ForegroundColor Magenta
do {
    # Get Jobs
    $updateJobs = skysync-cli jobs list --search $searchCriteria --limit 1000 --offset $offset --json true | ConvertFrom-Json
    If ($updateJobs.count -gt 0) {
        Start-SkySyncMigration -jobs $updateJobs -restURL $baseURL -restHeader $authHeader -schedJSON $scheduleJSON -ct  $contentType
    }
    Else {
        Write-Host ("No jobs were found based on search criteria: {0}" -f $searchCriteria) -ForegroundColor Yellow
    }

    # Incrementing Offset for paging
    $offset += 1000
}
while ($offset -lt $offsetTotal)
#endregion

#region Completed Screen Output
Complete-CatScript $scriptStart[0]
Stop-Transcript
#endregion