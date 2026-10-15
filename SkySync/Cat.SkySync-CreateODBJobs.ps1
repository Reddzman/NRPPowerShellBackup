<#  Cat.SkySync-CreateODBJobs.ps1

This script iterate through a CSV file to create Box to OneDrive for Business migration jobs.

Dependencies
	-	SkySync CLI
	-	..\General\Cat.PnP-GeneralFunctions.psm1
	-	.\Cat.SkySyncFunctions.psm1
	-	Input Files located in _Input Files
		-	Jobs File
		-	Template File
		-	Accounts File

Call Script with Input File
	.\Cat.SkySync-CreateSPOConnections.ps1 -jobsFile "Cat.SkySync-ODBJobs-client-Test.csv" -templateFile "Cat.SkySync-ODBSimTemp-client.json" -accountsFile "Cat.SkySync-Accounts-client-Test.json" -categoryName "TestWave-TestBU-TestDepartment-BoxToODB" -sourcePool "Box-Prod" -desitnationPool "ODB-Pool01" -adminURL "https://client-admin.sharepoint.com" -onedriveURL "https://client-my.sharepoint.com"

CSV Example
	sourceEmail,rootFolder,destinationUserID,destinationEmail
	nick.patterson@client.com,Pre-Existing Test,nick.patterson,nick.patterson@client.com
	nick.patterson@client.com,Test 1,nick.patterson,nick.patterson@client.com
	nick.patterson@client.com,Test 2,nick.patterson,nick.patterson@client.com

Template File and Accounts File Example at end of script.

Table of Contents
	1.	Input & Output Files
	2.	Variables
	3.	Snapin(s), Modules & Environment Connection(s)
	4.	Main Script
	5.	Check OneDrive Provisioning
	6.	Add Service Account Permissions
	7.	Create simulation job
	8.	Start Simulation Job
#>

#region Script Parameters
Param(  
	[String] $jobsFile = "Cat.SkySync-ODBJobs-client-Test.csv",
    [String] $folderTemplate = "Cat.SkySync-ODBSimTemp-RootFolders-client-ConflictLatest.json",
    [String] $fileTemplate = "Cat.SkySync-ODBSimTemp-RootFiles-client-ConflictLatest.json",
	[String] $accountsFile = "Cat.SkySync-Accounts-client.json",
    [String] $categoryName = "TestWave-TestBU",
    [String] $deptName = "TestDept",
	[String] $sourcePool = "Box-Prod",
	[String] $desitnationPool = "ODB-Pool01",
	[String] $adminURL = "https://client-admin.sharepoint.com",
	[String] $onedriveURL = "https://client-my.sharepoint.com"
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
Start-Transcript -Path ("..\_Log FIles\CreateODBJobs-{0}.log" -f $scriptStart[1])
#endregion

#region 1.	Input & Output Files
# File Paths
$jobsPath = ("..\_Input Files\{0}" -f $jobsFile)
$folderTemplatePath = ("..\_Input Files\{0}" -f $folderTemplate)
$fileTemplatePath = ("..\_Input Files\{0}" -f $fileTemplate)
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

# Load Folder Template JSON
If (Test-Path ($folderTemplatePath)){
    $folderInput = Get-Content -Path $folderTemplatePath -Raw | ConvertFrom-Json
}
else {
	Write-Host
	throw ("The folder template file does not exist: {0}" -f $folderTemplatePath)	
}

# Load File Template JSON
If (Test-Path ($fileTemplatePath)){
    $fileInput = Get-Content -Path $fileTemplatePath -Raw | ConvertFrom-Json
}
else {
	Write-Host
	throw ("The file template file does not exist: {0}" -f $folderTemplatePath)	
}

# Load accounts JSON
If (Test-Path ($acctPath)){
	$acctInput = Get-Content -Path $acctPath -Raw | ConvertFrom-Json
}
else {
	Write-Host
	throw ("The accounts file does not exist: {0}" -f $conPath)	
}
#endregion

#region 2.	Variables
# User Inputs
Write-Host
Write-Host "User Input Variables" -ForegroundColor Cyan
Write-Host "   Enter credentials for SharePoint:"
$spCred = Get-Credential -Message "Enter your credentials for SharePoint:"

# Script Variables
$adminURL = "https://client-admin.sharepoint.com"
$migrationFolder = "Box Migration"
$sleep = 30
$currentUser = "temp@temp.com"
$serverURL = "http://localhost:9090/"
$baseURL = ("{0}v1" -f $serverURL)
$contentType = "application/json"

# Output Parameters & Variables
Write-Host
Write-Host "Script Parameters" -ForegroundColor Cyan
Write-Host ("   Jobs File: {0}" -f $jobsFile)
Write-Host ("   Folder Template: {0}" -f $folderTemplate)
Write-Host ("   File Template: {0}" -f $fileTemplate)
Write-Host ("   Accounts File: {0}" -f $accountsFile)
Write-Host ("   Category: {0}" -f $categoryName)
Write-Host ("   Department: {0}" -f $deptName)
Write-Host ("   Source Connection Pool: {0}" -f $sourcePool)
Write-Host ("   Destination Connection Pool: {0}" -f $desitnationPool)
Write-Host ("   Admin URL: {0}" -f $adminURL)
Write-Host ("   OneDrive URL: {0}" -f $onedriveURL)
Write-Host
Write-Host "Script Variables" -ForegroundColor Cyan
Write-Host ("   Admin URL: {0}" -f $adminURL)
Write-Host ("   SPO Credentials: {0}" -f $spCred.UserName)
Write-Host ("   Migration Folder: {0}" -f $migrationFolder)
Write-Host ("   Provisioning Sleep: {0}" -f $sleep)
<#Write-Host
Write-Host "Folder Template JSON" -ForegroundColor Cyan
$folderInput | ConvertTo-Json
Write-Host
Write-Host "Folder Template JSON" -ForegroundColor Cyan
$fileInput | ConvertTo-Json #>
#endregion

#region 3.	Snapin(s), Modules & Environment Connection(s)
# Load Custom Modules
Import-Module Microsoft.Online.SharePoint.PowerShell -DisableNameChecking
Import-Module "..\SkySync\Cat.SkySyncFunctions.psm1"
#endregion

#region 4.	Main Script
# Loop through jobs CSV file
foreach ($job in $jobsInput) {
	Write-Host
	Write-Host ("Configuring Job: {0} - {1}" -f $job.destinationEmail, $job.rootFolder) -ForegroundColor Magenta

    if ($currentUser -ne $job.sourceEmail){
        $currentUser = $job.sourceEmail

        #region Refreshing Auth Token for long running scripts
        $accessToken = skysync-cli sessions token --output-json false
        $authHeader = @{"Authorization"="Bearer $accessToken";}
        #endregion
        
        #region 5.	Check OneDrive Provisioning
        $personalProv = $false
        $personalURL = ("{0}/personal/{1}" -f $onedriveURL, $job.destinationEmail.Replace(".","_").Replace("@","_").Replace("\","_").Replace("/","_"))
        Try {
            Write-Host
            Write-Host ("   Checking OneDrive Site: {0}" -f $job.destinationEmail) -NoNewline -ForegroundColor DarkGray
            Connect-PnPOnline -Url $personalURL -Credentials $spCred
            Write-Host " - Provisioned" -ForegroundColor Green
        } # Try
        Catch {
            #Checking alternate Personal URL
            try {
                Connect-PnPOnline -Url ("{0}1" -f $personalURL) -Credentials $spCred
                Write-Host " - Provisioned" -ForegroundColor Green
                $personalProv = $false
                $personalURL = ("{0}1" -f $personalURL)
            }
            catch {
                Write-Host " - Not Provisioned" -ForegroundColor Red
                Write-Host
                $personalProv = $true
            }
        } # Catch

        # Provision OneDrive
        If ($personalProv){
            Write-Host ("   Provisioning OneDrive: {0}" -f $job.destinationEmail) -NoNewline  -ForegroundColor DarkGray
            Connect-PnPOnline -Url $adminURL -Credentials $spCred
            New-PnPPersonalSite -Email $job.destinationEmail

            # Waiting for provisioning to complete
            Write-Host " - Working on it." -NoNewline -ForegroundColor DarkGray
            $complete = $false
            $trys = 0
            While ((-not $complete) -and ($trys -le 30)) {
                Try
                {
                    Connect-PnPOnline -Url $personalURL -Credentials $spCred
                    $complete = $true
                    Write-Host " - Provisioned" -ForegroundColor Green
                } # Try
                Catch
                {
                    $complete = $false
                    $trys++
                    Write-Host "." -NoNewline -ForegroundColor DarkGray
                    #Write-Host ". Exception Message: $($_.Exception.Message)" # use to troubleshoot connection issues
                    Start-Sleep -Seconds $sleep
                } # Catch
            } # While (-not $complete)
        } # If ($personalProv)
        Disconnect-PnPOnline
        #endregion

        #region 6.	Add Service Account Permissions
        Write-Host
        Write-Host "   Adding Permissions for Service Accounts" -ForegroundColor DarkGray
        Connect-SPOService -Url $adminURL -Credential $spCred
        foreach	($acct in $acctInput){
            Write-Host ("      {0}" -f $acct.username)
            $tempUser = Set-SPOUser -Site $personalURL -LoginName $acct.username -IsSiteCollectionAdmin $true
        } # foreach	($acct in $acctInput)
        #endregion
        
        #region 7.	Create file job
        # Determine if job already exists
        $jobName = ("ODB-{0}-{1}-{2}-{3}" -f $categoryName, $deptName, $job.destinationUserID, "RootFiles")
        If (JobExists($jobName)){
            Write-Host ("  The job already exists: {0}" -f $jobName) -ForegroundColor Yellow
        }
        Else {
            Write-Host
            Write-Host ("   Creating root files job: {0}" -f $job.destinationEmail) -ForegroundColor DarkGray
            Write-Host ("      Updating template JSON: {0}" -f $job.destinationEmail)
            
            # Update job template
            $template = $fileInput
            $template.name = $jobName
            $template.category.name = $categoryName
            $template.transfer.source.connection.id = GetConnectionID $sourcePool
            $template.transfer.source.impersonate_as.email = $job.sourceEmail
            $template.transfer.source.target.path = "/"
            $template.transfer.destination.connection.id = GetConnectionID $desitnationPool
            $template.transfer.destination.impersonate_as.username = $job.destinationEmail
            $template.transfer.destination.impersonate_as.email = $job.destinationEmail
            $template.transfer.destination.target.path = ("/{0}" -f $migrationFolder)

            # Create Job
            Write-Host ("      Creating Job: {0}" -f $jobName)
            $fileJob = $template | ConvertTo-Json -Depth 10
            $createJobMethod = "$baseUrl/jobs"
		    $createJobResponse = Invoke-RestMethod $createJobMethod -Headers $authHeader -Method Post -ContentType $contentType -Body $fileJob
            
            # 8.	Start Simulation Job
            Write-Host ("      Starting Job: {0}" -f $jobName)
            $fileStart = skysync-cli jobs start --search $jobName --json | ConvertFrom-Json
        }
        #endregion
    }

	#region 8.	Create folder job
	# Determine if job already exists
    $jobName = ("ODB-{0}-{1}-{2}-{3}" -f 
        $categoryName,
        $deptName,
        $job.destinationUserID, 
        $job.RootFolder.Replace(" ","").Replace(" ","").Replace("(","").Replace(")","").Replace("&","").Replace("-","").Replace(":","").Replace('"',"").Replace("*","").Replace("<","").Replace(">","").Replace("?","").Replace("/","").Replace("\","").Replace("|","").Replace(".","")
    )
	If (JobExists($jobName)){
		Write-Host ("  The job already exists: {0}" -f $jobName) -ForegroundColor Yellow
	}
	Else {
        Write-Host
		Write-Host ("   Updating template JSON: {0}" -f $job.destinationEmail) -ForegroundColor DarkGray
        
        # Update job template
        $template = $folderInput
        $template.name = $jobName
		$template.category.name = $categoryName
		$template.transfer.source.connection.id = GetConnectionID $sourcePool
		$template.transfer.source.impersonate_as.email = $job.sourceEmail
		$template.transfer.source.target.path = ("/{0}" -f $job.RootFolder)
		$template.transfer.destination.connection.id = GetConnectionID $desitnationPool
		$template.transfer.destination.impersonate_as.username = $job.destinationEmail
		$template.transfer.destination.impersonate_as.email = $job.destinationEmail
        $template.transfer.destination.target.path = ("/{0}/{1}" -f 
            $migrationFolder, 
            $job.RootFolder.Replace(":"," -").Replace('"',"").Replace("*","").Replace("<","").Replace(">","").Replace("?","").Replace("/","").Replace("\","").Replace("|",""))

        # Create Job
        Write-Host
        Write-Host ("   Creating Job: {0}" -f $jobName) -ForegroundColor DarkGray
        $folderJob = $template | ConvertTo-Json -Depth 10
        $createJobMethod = "$baseUrl/jobs"
        $createJobResponse = Invoke-RestMethod $createJobMethod -Headers $authHeader -Method Post -ContentType $contentType -Body $folderJob
		
        # 8.	Start Simulation Job
        Write-Host
		Write-Host ("   Starting Job: {0}" -f $jobName) -ForegroundColor DarkGray
		$folderStart = skysync-cli jobs start --search $jobName --json | ConvertFrom-Json
	}
	#endregion
}
#endregion

#region Completed Screen Output
Complete-CatScript $scriptStart[0]
Disconnect-SPOService
Stop-Transcript
#endregion

<# Template File
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
            "id": "2c2c0aceeb324e60abe757059fa280f7",
            "type": "account_map"
        },
        "group_map": {
            "id": "1aba77cdd51747ce9b0a3600931810d1",
            "type": "group_map"
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
                "id": "66a56d049e77420095094606be91d119"
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
                "id": "3cc743ca15c64ed08dac3365f466ce06"
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

<# Accounts File
[
    {
        "name":  "SPO-SA1-",
        "platform":  "office365",
        "type":  "automatic",
        "username":  "o365-skysync-odb-sa1@client.com",
        "password":  "TempPassword",
        "uri":  "TempURL"
    },
    {
        "name":  "SPO-SA2-",
        "platform":  "office365",
        "type":  "automatic",
        "username":  "o365-skysync-odb-sa2@client.com",
        "password":  "TempPassword",
        "uri":  "TempURL"
    },
    {
        "name":  "SPO-SA3-",
        "platform":  "office365",
        "type":  "automatic",
        "username":  "o365-skysync-odb-sa3@client.com",
        "password":  "TempPassword",
        "uri":  "TempURL"
    },
    {
        "name":  "SPO-SA4-",
        "platform":  "office365",
        "type":  "automatic",
        "username":  "o365-skysync-odb-sa4@client.com",
        "password":  "TempPassword",
        "uri":  "TempURL"
    },
    {
        "name":  "SPO-SA5-",
        "platform":  "office365",
        "type":  "automatic",
        "username":  "o365-skysync-odb-sa5@client.com",
        "password":  "TempPassword",
        "uri":  "TempURL"
    },
    {
        "name":  "SPO-SA6-",
        "platform":  "office365",
        "type":  "automatic",
        "username":  "o365-skysync-spo-sa1@client.com",
        "password":  "TempPassword",
        "uri":  "TempURL"
    },
    {
        "name":  "SPO-SA7-",
        "platform":  "office365",
        "type":  "automatic",
        "username":  "o365-skysync-spo-sa2@client.com",
        "password":  "TempPassword",
        "uri":  "TempURL"
    },
    {
        "name":  "SPO-SA8-",
        "platform":  "office365",
        "type":  "automatic",
        "username":  "o365-skysync-spo-sa3@client.com",
        "password":  "TempPassword",
        "uri":  "TempURL"
    },
    {
        "name":  "SPO-SA9-",
        "platform":  "office365",
        "type":  "automatic",
        "username":  "o365-skysync-spo-sa4@client.com",
        "password":  "TempPassword",
        "uri":  "TempURL"
    },
    {
        "name":  "SPO-SA10-",
        "platform":  "office365",
        "type":  "automatic",
        "username":  "o365-skysync-spo-sa5@client.com",
        "password":  "TempPassword",
        "uri":  "TempURL"
    }
]
#>