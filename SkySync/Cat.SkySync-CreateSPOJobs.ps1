<#  Cat.SkySync-CreateSPOJobs.ps1

This script iterate through a CSV file to create Box to SharePoint Online migration jobs.

Dependencies
	-	SkySync CLI
	-	..\General\Cat.PnP-GeneralFunctions.psm1
	-	.\Cat.SkySyncFunctions.psm1
	-	Input CSV Files located in _Input Files
		-	Jobs File
		-	Template File

Call Script with Input File
	.\Cat.SkySync-CreateSPOConnections.ps1 -jobsFile "Cat.SkySync-SPOJobs-client-Test.csv" -templateFile "Cat.SkySync-SPOSimTemp-client.json" -categoryName "TestWave-TestBU-TestDepartment-BoxToSPO" -sourcePool "Box-Prod"

CSV Jobs Example
	sourceEmail,rootFolder,destinationUserID,destinationEmail,destinationlibrary,spoPool
	nick.patterson@client.com,Shared Test 1,nick.patterson,nick.patterson@client.com,migtest,SPO-TestBU-TestDept-TestSite-Pool
	nick.patterson@client.com,Shared Test 2,nick.patterson,nick.patterson@client.com,migtest,SPO-TestBU-TestDept-TestSite-Pool

JSON Template file example at end of Script

Table of Contents
	1.	Input & Output Files
	2.	Variables
	3.	Snapin(s), Modules & Environment Connection(s)
	4.	Main Script
	5.	Create simulation job
	6.	Start Similation Job
#>

#region Script Parameters
Param(  
	[String] $jobsFile = "Cat.SkySync-SPOJobs-client-Test.csv",
	[String] $templateFile = "Cat.SkySync-SPOSimTemp-client-ConflictLatest.json",
    [String] $categoryName = "TestWave-TestBU",
    [String] $deptName = "TestDept",
	[String] $sourcePool = "Box-Prod"
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
Start-Transcript -Path ("..\_Log FIles\CreateSPOJobs-{0}.log" -f $scriptStart[1])
#endregion

#region 1.	Input & Output Files
# File Paths
$jobsPath = ("..\_Input Files\{0}" -f $jobsFile)
$templatePath = ("..\_Input Files\{0}" -f $templateFile)
$acctPath = ("..\_Input Files\{0}" -f $accountsFile)

# Import Input File(s)
# Load jobs CSV
If (Test-Path ($jobsPath)){
	$jobsInput = Import-Csv $jobsPath -Encoding Unicode
}
else {
	Write-Host
	throw ("The jobs file does not exist: {0}" -f $jobsPath)	
}

# Load Template JSON
If (Test-Path ($templatePath)){
	$templateInput = Get-Content -Path $templatePath -Raw | ConvertFrom-Json
}
else {
	Write-Host
	throw ("The template file does not exist: {0}" -f $templatePath)	
}
#endregion

#region 2.	Variables
# Script Variables
$migrationFolder = "Box Migration"
$serverURL = "http://localhost:9090/"
$baseURL = ("{0}v1" -f $serverURL)
$contentType = "application/json"
$currentUser = "temp@temp.com"

# Output Parameters & Variables
Write-Host
Write-Host "Script Parameters" -ForegroundColor Cyan
Write-Host ("   Jobs File: {0}" -f $jobsFile)
Write-Host ("   Template File: {0}" -f $templateFile)
Write-Host ("   Category: {0}" -f $categoryName)
Write-Host ("   Source Connection Pool: {0}" -f $sourcePool)
Write-Host
Write-Host "Script Variables" -ForegroundColor Cyan
Write-Host ("   Migration Folder: {0}" -f $migrationFolder)
Write-Host ("   SkySync URL: {0}" -f $serverURL)
Write-Host ("   Base URL: {0}" -f $baseURL)
Write-Host ("   Content Type: {0}" -f $contentType)
Write-Host ("   Current User: {0}" -f $currentUser)
#endregion

#region 3.	Snapin(s), Modules & Environment Connection(s)
# Load Custom Modules
Import-Module "..\SkySync\Cat.SkySyncFunctions.psm1"
#endregion

#region 4.	Main Script
# Loop through jobs CSV file
foreach ($job in $jobsInput) {
    Write-Host
	Write-Host ("Configuring Job: {0} - {1}" -f $job.destinationEmail, $job.rootFolder) -ForegroundColor Magenta

    #region Refreshing Auth Token for long running scripts
    if ($currentUser -ne $job.sourceEmail){
        $currentUser = $job.sourceEmail
        $accessToken = skysync-cli sessions token --output-json false
        $authHeader = @{"Authorization"="Bearer $accessToken";}
    }
    #endregion

	#region 5.	Create simulation job
	# Determine if job already exists
    $jobName = ("SPO-{0}-{1}-{2}-{3}" -f
        $categoryName,
        $deptName,
        $job.destinationUserID,
        $job.RootFolder.Replace(" ","").Replace(" ","").Replace("(","").Replace(")","").Replace("&","").Replace("-","")
    )
	If (JobExists($jobName)){
		Write-Host ("  The job already exists: {0}" -f $jobName) -ForegroundColor Yellow
	}
	Else {
		Write-Host ("   Updating template JSON: {0}" -f $job.destinationEmail)
		# Update job template
        $template = $templateInput
        $template.name = $jobName
		$template.category.name = $categoryName
		$template.transfer.source.connection.id = GetConnectionID $sourcePool
		$template.transfer.source.impersonate_as.email = $job.sourceEmail
		$template.transfer.source.target.path = ("/{0}" -f $job.RootFolder)
		$template.transfer.destination.connection.id = GetConnectionID $job.spoPool
        $template.transfer.destination.target.path = ("/{0}/{1}/{2}" -f $job.destinationLibrary,
            $migrationFolder, 
            $job.RootFolder.Replace(":"," -").Replace('"',"").Replace("*","").Replace("<","").Replace(">","").Replace("?","").Replace("/","").Replace("\","").Replace("|",""))
		$template.transfer.account_map.id = Get-SkySyncMapID -mapType "account_map" -mapName $job.spoPool -restURL $baseURL -restHeader $authHeader -ct $contentType
		$template.transfer.group_map.id = Get-SkySyncMapID -mapType "group_map" -mapName $job.spoPool -restURL $baseURL -restHeader $authHeader -ct $contentType

		# Create Job
        Write-Host ("   Creating Job: {0}" -f $jobName)
        $spoJob = $template | ConvertTo-Json -Depth 10
        $createJobMethod = "$baseUrl/jobs"
        $createJobResponse = Invoke-RestMethod $createJobMethod -Headers $authHeader -Method Post -ContentType $contentType -Body $spoJob
		#$newJob = $template | ConvertTo-Json -Depth 10 | skysync-cli jobs add --name $jobName --manual --in --json | ConvertFrom-Json
		
		# 6.	Start Similation Job
		Write-Host ("   Starting Job: {0}" -f $jobName)
		$startJob = skysync-cli jobs start --search $jobName --json | ConvertFrom-Json
	}
	#endregion
}
#endregion

#region Completed Screen Output
Complete-CatScript $scriptStart[0]
Stop-Transcript
#endregion

<# JSON Template Sample
{
    "category": {"name": "{Category Name}"},
    "transfer": {
        "simulation_mode": true,
        "transfer_type":"Sync",
        "audit_level":"trace",
        "batch_mode":"always",
        "performance":{
            "parallel_writes":{
            "requested":6
            }
        },
        "account_map": {
            "name":"{account_map_name}",
            "type": "account_map",
            "unmapped_policy": "warn",
            "map_by": {
                "ID": true,
                "email": true,
                "external_passthrough": false
            }
        },
        "group_map": {
            "name": "{group_map_name}",
            "type": "group_map",
            "unmapped_policy": "warn",
            "map_by": {
                "ID": true
            }
         },
        "versioning":{
            "preserve":"native",
            "select":"all"
        },
        "timestamps":true,
        "preserve_owners":true,
        "permissions": {
            "policy": "add"
        },
        "conflict_resolution":"Copy",
        "delete_propagation":"mirror",
        "lock_propagation":"mirror_lock",
        "segment_transform":true,
        "filter": {
            "source": [
                {
                    "action": "exclude",
                    "rules": [
                        {
                            "type": "filter_shared"
                        }
                    ],
                    "type": "filter_rule"
                }
            ],
            "destination": [
                {
                    "action": "exclude",
                    "rules": [
                        {
                            "type": "filter_shared"
                        }
                    ],
                    "type": "filter_rule"
                }
            ]
        },
        "source": {
            "connection": {
                "id": "{Source Connection ID}"
            },
            "impersonate_as": {
                "email": "{source_email}"
            },
            "target":{
                "path":"{Replace Source Path}"
             }
        },
        "destination": {
            "connection": {
                "id": "{Destination Connection ID}"
            },
            "target":{
                "path":"{Replace Destination Path}"
             },
            "impersonate_as":{
                "username":"{destination_user_name}",
                "email":"{destination_user_name}"
            }
        }
    },
    "schedule": {
        "mode": "manual"
    },
    "stop_policy": {
        "on_failure": 96
    }
}
#>