<#
Cat.CM-NightlyIncMigWeb.ps1

This script runs an incremental migration for Site Collections listed in a CSV file.  It breaks up
the jobs at a web level and runs lists and libraries only.  It will only create/run jobs for sites
modified in the last X days.

Dependencies
	-	Metalogix PowerShell Snapins
	-	BCT.PnP-MigrationFunctions.psm1
	-	SQL Server Command Utility - https://www.microsoft.com/en-us/download/details.aspx?id=53591
	-	Input CSV File in Input Files folder - Script Parameter - "CM-SmallMigrations.csv"
	-	Cat.CMFunctions should be updated with data from "create powershell" feature in Content Matrix

Call Script with Input File
	.\BCT.CM-LargeIncrementalJob.ps1 -InputFile CM-LargeMigrations-ITS.csv -SPUser "client\npatterson" -SPPass "client018" -DaysBack 90 -JobSleep 60 -MaxSessions 6 -SQLServer "MSGNM-SPSTG01" -SQLDB "SPOMig_JobDB" $SQLUser "SPOMig" -SQLPass "Sh4r3Po!nt0nl!n3"

CSV Example
	Site,SP13Url,SPOUrl,ContentType
	Site Name,https://sharepoint.client.com/sites/name,https://client.sharepoint.com/sites/name,Client Document

Table of Contents
	1.	Input & Output Files
	2.	Variables
	3.	Snapin(s), Modules & Environment Connection(s)
	4.	Main Script
	5.	Run Metalogix Job
#>

#region Script Parameters
Param(
		[string]$InputFile,
		[string]$SPUser,
		[string]$SPPass,
		[int]$DaysBack,
		[int]$JobSleep,
		[int]$MaxSessions,
		[string]$SQLServer,
		[string]$SQLDB,
		[string]$SQLUser,
		[string]$SQLPass
	)
#endregion

#region Start Date / Time
# Set PS location
Set-Location $PSScriptRoot

# Get Script Parent Location
$LocSplit = $PSScriptRoot.Split("\")
$LocReplace = $LocSplit[($LocSplit.Count - 1)]
$Location = $PSScriptRoot.Replace($LocReplace,"")

# Remove Existing Catapult Modules to ensure PS session has most up to date functions
Get-Module Cat.* | Remove-Module

# Load General Functions
Import-Module ("{0}General\Cat.GeneralFunctions.psm1" -f $Location)

# Get Start Date / Time
$Start = Start-CatScript
#endregion

#region 1.	Input & Output Files
# File Paths
$InPath = ("{0}_Input Files\{1}" -f $Location, $InputFile)

# Import Input File(s)
$InFile = import-csv $InPath -Encoding Unicode
#endregion

#region 2.	Variables
# Create PS Credentials for SPO site
$SPSecPass = ConvertTo-SecureString $SPPass -AsPlainText -Force
$SPCred = New-Object System.Management.Automation.PSCredential ($SPUser, $SPSecPass)

#SQL Variables
$MigrationDate = (Get-Date).AddDays(-$DaysBack)
#endregion

#region 3.	Snapin(s), Modules & Environment Connection(s)
# Load Custom Modules
Import-Module ("{0}Content Matrix\Cat.CMFunctions.psm1" -f $Location)
#endregion
    
#region 4.	Main Script
Write-Host
Write-Host "Running Nightly Incremental Migrations" -ForegroundColor DarkCyan

# Looping through Sites in CSV Input
ForEach ($Job in $InFile)
{
	# Output Site to Screen
	Write-Host
	Write-Host ("   Job Site: {0}" -f $Job.Site) -ForegroundColor Magenta

	# 5.	Run Migration Jobs
	Wait-SystemResourcesFree
	Wait-PSSession -MaxSessions $MaxSessions
	$scriptBlock = 
		{
			param($MPath, $SUrl, $TUrl, $CT, $JobName, $SSvr, $LogDB, $SUser, $SPass)
			Import-Module $MPath
			Migrate-CMAllChildLists -SourceURL $SUrl -TargetURL $TUrl -ContentType $CT -JobName $JobName -Server $SSvr -Database $LogDB -User $SUser -Pass $SPass
		}	

	$PSJob = Start-Job -ScriptBlock $scriptBlock -ArgumentList ("{0}Content Matrix\Cat.CMFunctions.psm1" -f $Location), $Job.SP13Url, $Job.SPOUrl, $Job.ContentType, ("{0} - Nightly Incremental - {1}" -f $Job.Site, $FileDate), $SQLServer, $SQLDB, $SQLUser, $SQLPass
	Write-Host ("      Starting Job for {0}: {1}" -f $Job.Site, $PSJob.Name)
	Start-Sleep $JobSleep

} # ForEach ($Job in $InFile

# Waiting for jobs to complete
Write-Host
Wait-BackgroundJobs

# Screen Output for site complete
Write-Host
Write-Host "Incremental Migration Complete" -ForegroundColor Green
#endregion

#region Completed Screen Output
Complete-CatScript $Start[0]
#endregion