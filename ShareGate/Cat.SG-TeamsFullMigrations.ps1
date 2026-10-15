<#  Cat.SG-TeamsFullMigrations.ps1

This script automates ShareGate Teams migrations.  If the destination TEAMS site doesn't exist
it will be created. 

NOTE: ShareGate only supports FULL migrations for TEAMS - 1/8/21

Dependencies
	-	Source & Destination Service Account for ShareGate
    -	ShareGate PowerShell Modules; Doesn't support PS Core 7 - Install-Module PnP.PowerShell -RequiredVersion 1.12 -Force
    -   ..\General\Cat.GeneralFunctions.psm1
    -   ..\Reporting\Cat.ReportingFunctions.psm1
    -	Input CSV File in Input Files folder - Script Parameter
    -   User mapping file export from ShareGate (UserMappingFile.sgum)

Call script with parameters
	.\Cat.SG-TeamsFullMigrations.ps1 -sourceTenant "ClientTenant" -destination "ClientTenant" -jobFile Cat.SG-TeamsList-Test.csv -userFile Cat.SG-UserMapping-Test.sgum

CSV Example
	sourceTeamName,destinationTeamName
    Source Team Site Title,Destionation Team Site Title

Table of Contents
	1.	Input & Output Files
	2.	Variables
	3.	Snapin(s), Modules & Environment Connection(s)
    4.	Main Script
#>

#region Script Parameters
Param(
    [String] $jobFile = "Cat.SG-TeamsList-Test.csv",
    [String] $userFile = "Cat.SG-UserMapping-Test.sgum",
    [string] $srcTenant = "{Source Tenant Name}",
    [string] $destTenant = "{Destination Tenant Name}"
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
Start-Transcript -Path ("..\_Log Files\{0}.SG-TeamsFullMigrations-{0}.log" -f $destTenant, $scriptStart[1])
#endregion

#region 1.	Input & Output Files
# File Paths
$inPath = ("..\_Input Files\{0}" -f $jobFile)
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
	throw ("The user mapping file does not exist: {0}" -f $userPath)	
}
#endregion

#region 2.	Variables
# User Inputs
Write-Host
Write-Host "Enter Connection Information" -ForegroundColor Cyan
Write-Host "   Enter ShareGate credentials for the SOURCE tenant:"
$spSrcCred = Get-Credential -Message "Enter ShareGate credentials for the SOURCE tenant:" 
Write-Host "   Enter ShareGate credentials for the DESTINATION tenant:"
$spDestCred = Get-Credential -Message "Enter ShareGate credentials for the DESTINATION tenant:" 

# Script Variables
$srcTenantDomain = ("{0}.onmicrosoft.com" -f $srcTenant) 
$destTenantDomain = ("{0}.onmicrosoft.com" -f $destTenant) 

# Output Parameters & Variables
Write-Host
Write-Host "Script Parameters" -ForegroundColor Cyan
Write-Host ("   Jobs File: {0}" -f $jobFile)
Write-Host ("   User Mapping File: {0}" -f $userFile)
Write-Host ("   Source Tenant: {0}" -f $srcTenant)
Write-Host ("   Destination Tenant: {0}" -f $destTenant)
Write-Host
Write-Host "User Inputs" -ForegroundColor Cyan
Write-Host ("   Source ShareGate Credentials: {0}" -f $spSrcCred.UserName)
Write-Host ("   Destination ShareGate Credentials: {0}" -f $spDestCred.UserName)
Write-Host
Write-Host "Script Variables" -ForegroundColor Cyan
Write-Host ("   Source Tenant Domain: {0}" -f $srcTenantDomain)
Write-Host ("   Destination Tenant Domain: {0}" -f $destTenantDomain)
#endregion

#region 3.	Snapin(s), Modules & Environment Connection(s)
# Load Sharegate modules
Import-Module Sharegate
#endregion
  
#region 4.	Main Script
Write-Host
Write-Host ("Configuring & running migrations for job list: {0}" -f $jobFile) -ForegroundColor Cyan

# Connecting to SOURCE and DESTINATION with ShareGate
Write-Host ("   Connecting to Source: {0}" -f $srcTenantDomain)
$srcTenantConnection = Connect-Tenant -Domain $srcTenantDomain -Credential $spSrcCred
Write-Host ("   Connecting to Destination: {0}" -f $destTenantDomain)
$destTenantConnection = Connect-Tenant -Domain $destTenantDomain -Credential $spDestCred

# Importing User Mappings from mapping file
Write-Host ("   Configuring mappings file: {0}" -f $userFile)
$copyMappings = New-MappingSettings
$copyMappings = Import-UserAndGroupMapping -MappingSettings $copyMappings -Path $userPath.Path

# Looping through input file
ForEach ($job in $inFile){
    # Connecting to Source Team
    Write-Host
    Write-Host ("   Migrating Team: {0}" -f $job.sourceTeamName) -ForegroundColor Magenta
    Write-Host ("      Connecting to Source Team: {0}" -f $job.sourceTeamName)
    $srcTeam =  Get-Team -Name $job.sourceTeamName -Tenant $srcTenantConnection    

    # Coping Team between tenants
    If ($srcTeam) {
        Write-Host ("      Copying Team: {0} into Team: {1}" -f $job.sourceTeamName, $job.destinationTeamName)
        $result = Copy-Team -Team $srcTeam -TeamTitle $job.destinationTeamName -DestinationTenant $destTenantConnection -MappingSettings $copyMappings -WarningAction SilentlyContinue
        Write-Host ("         Teams Migration Completed: {0}" -f $job.destinationTeamName) -foregroundcolor "green"   
    }
    else {
        Write-Host ("         Source Team does not exist! Please check the input file: {0}" -f $srcTeamName) -ForegroundColor Red
    }         
}

# Complete Output
Write-Host
Write-Host ("Teams Migration completed. Please refer to the ShareGate Job Logs for any warnings or errors that require remediation") -ForegroundColor Cyan
#endregion

#region Completed Script
# Output Script Info
Complete-CatScript $scriptStart[0]

# Stop logging
Stop-Transcript
#endregion