<#  Cat.SG-BoxFolderMigrations.ps1

This script automates the migration of specific folders from ox to OneDrive and SharePoint Online. It can migrate to the root of a document library or to a
folder created in the destination. Ensure site, library, and folder existence in advance using the Cat.SG-BoxPreMigrationConfiguration.ps1. Both scripts leverage the 
same CSV format. 

*** Although the sample CSV below shows them together, OneDrive and SharePoint jobs need to be ran in seperate csv files/job runs. ***

Dependencies
    -   Destination Service Account for ShareGate
    -	ShareGate PowerShell Modules (Doesn't support PS Core 7) - https://documentation.sharegate.com/hc/en-us/articles/115000343993-Import-from-Box-com 
    -	..\General\Cat.GeneralFunctions.psm1
    -	Input CSV File in Input Files folder - ..\_Input Files\Samples\ShareGate\Cat.SG-BoxFolderMigrations-Test1.csv

Call script with parameters
    cd C:\Quisitive\PowerShell\ShareGate
	.\Cat.SG-BoxFolderMigrations.ps1 -boxAdmin "Box.Admin@SailPoint.com" -jobFile "Cat.SG-BoxFolderMigrations-ODB.csv" -userFile "Cat.SG-UserMapping-Test.sgum" -dstTenant "sailpoint-my" -dstTech "ODB"

CSV Sample
    _____________________________ Source & Destination Information _______________________________---------------------------------------- Site Creation -----------------------------------------------------~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ ShareGate Settings ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    srcPath,dstSiteTitle,dstURL,dstGroupAlias,dstLibrary,dstLibraryPath,dstFolder,dstType,dstEmail,siteTemplate,CommSiteDesign,siteOwner,modernTeamSiteIsPublic,isHubSite,hubParentURL,classicSubSiteParentURL,onWarning,onError,onSiteObjectExists,onContentItemExists,versionComment,noCustomPermissions,versionLimit
    /Directory/Department Folders/Moddern Comm Topic Hub,Migration - Modern Comm Topic Hub Test16,https://nrpdev2.sharepoint.com/sites/migrationmoderncommtopichubTest16,,Documents,Shared Documents,,SPO,,CommunicationSite,Topic,nickp@nrpdev.com,FALSE,TRUE,,,Continue,Skip,Merge,IncrementalUpdate,,TRUE,2
    /Directory/Department Folders/Modern Team Private Documents,Migration - Modern Team Private Test16,https://nrpdev2.sharepoint.com/teams/migrationmodernteamprivTest16,migrationmodernteamprivTest16,Documents,Shared Documents,,SPO,,TeamSite,,nickp@nrpdev.com,FALSE,FALSE,https://nrpdev2.sharepoint.com/sites/migrationmoderncommtopichubTest16,,Continue,Skip,Merge,IncrementalUpdate,,TRUE,2
    /Directory/Department Folders/OneDrive Nick Patterson,Nick Patterson,https://nrpdev2-my.sharepoint.com/personal/nickp_nrpdev_com,,Documents,Documents,,ODB,,,,nickp@nrpdev.com,FALSE,FALSE,,,Continue,Skip,Merge,IncrementalUpdate,,TRUE,2

Table of Contents
	1.	Snapin(s), Modules & Environment Connection(s)
    2.	Input & Output Files
	3.	Variables
    4.	Main Script
    5.  Authentication
    6.  Running Migration jobs
    7.  Configure ShareGate Settings
    8.  Run ShareGate Job
#>

#region Script Parameters
Param(
    [String] $boxAdmin = "Box.Admin@SailPoint.com",# Must be box admin, co-admin isn't sufficient
    [String] $jobFile = "Cat.SG-SPOnPremListMigrations-Client.csv", # must exist in the root of the _Input Files folder
    [String] $userFile = "Cat.SG-UserMapping-CLIENT.sgum",
    [string] $dstTenant = "sailpoint", # $dstTenant.sharepoint.com; include "-my" for OneDrive migrations
	[String] $dstTech = "SPO" # must be "SPO" or "ODB"
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
Start-Transcript -Path ("..\_Log Files\{0}.SG-BoxFolderMigrations-{1}.log" -f $dstTenant, $scriptStart[1])
#endregion

#region 1.	Snapin(s), Modules & Environment Connection(s)
# Load Sharegate modules
Import-Module Sharegate	
#endregion

#region 2.	Input & Output Files
# File Paths
$inPath = ("..\_Input Files\{0}" -f $jobFile)
$userPath = ("..\_Input FIles\{0}" -f $userFile)
$errorPath = ("..\_Output Files\{0}.SG-BoxFolderMigrationErrors-{1}.log" -f $dstTenant, $scriptStart[1])

# Import Input File(s)
# Load Job file
If (Test-Path ($inPath)){
	$inFile = Import-Csv $inPath -Encoding Unicode
}
else {
	Write-Host
	throw ("The jobs file does not exist: {0}" -f $inPath)	
}

# Load User Mapping file
If (Test-Path ($userPath)){
	$userPath = Resolve-Path $userPath
}
else {
	Write-Host
	throw ("The user mapping file does not exist: {0}" -f $userPath)	
}
#endregion

#region 3.	Variables
# Script Variables
$dstRootSiteUrl = ("https://{0}.sharepoint.com/" -f $dstTenant)
Set-Variable dstSite, dstList
 
#region Variable Output
Write-Host
Write-Host "Script Parameters" -ForegroundColor Cyan
Write-Host ("   Box Admin: {0}" -f $boxAdmin)
Write-Host ("   Job File: {0}" -f $jobFile)
Write-Host ("   User File: {0}" -f $jobFile)
Write-Host ("   Destination Tenant URL: {0}" -f $dstTenant)
Write-Host ("   Destination Technology (SPO/ODB): {0}" -f $dstTech)
Write-Host
Write-Host "Script Variables" -ForegroundColor Cyan
Write-Host ("   Root Site URL: {0}" -f $dstRootSiteUrl)
#endregion
#endregion

#region 4.	Main Script
Write-Host
Write-Host ("Migrating Box Folders to {0}: {1}" -f $dstTech, $jobFile) -ForegroundColor Cyan

#region 5.  Authentication
# Box.com Connection process
Write-Host ("   {0} - Connecting to Box" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"))
try {
	$boxAuth = Connect-Box -Email $boxAdmin -Admin
}
catch {
	Write-CatErrorCSV -funcCSVItem $job -funcAction "Connecting to Box" -funcException $_.Exception -funcDetails $_.ErrorDetails -funcStackTrace $_.ScriptStackTrace -funcFilePath $errorPath
	throw ("Error Connecting to Box")
}

# M365 Connection Process
Write-Host ("   {0} - Connecting to {1}" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"),$dstTech)
try {
	$dstAuth = Connect-Site -Url $dstRootSiteUrl -Browser -DisableSSO
}
catch {
	Write-CatErrorCSV -funcCSVItem $job -funcAction ("Connecting to {0}" -f $dstTech) -funcException $_.Exception -funcDetails $_.ErrorDetails -funcStackTrace $_.ScriptStackTrace -funcFilePath $errorPath
	throw ("Error Connecting to {0}" -f $dstTech)
}
#endregion

#region 6.  Running Migration jobs
foreach ($job in $inFile) {
    Write-Host
    Write-Host ("{0} - Running Migration: {1}" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"), $job.srcPath) -ForegroundColor Magenta
 
    # Connect to Destination site/library
    Write-Host ("   {0} - Loading Destination Library" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"))
    $dstSite = Connect-Site -Url $job.dstURL -UseCredentialsFrom $dstAuth
    $dstList = Get-List -Name $job.dstLibrary -Site $dstSite

    #region 7.  Configure ShareGate Settings
    # Create ShareGate Mapping settings
    Write-Host ("   {0} - Creating Mapping Settings and Adding User Mapping: {1}" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"), $userFile)
    $sgMappings= New-MappingSettings
    $sgMappings= Import-UserAndGroupMapping -MappingSettings $sgMappings -Path $userPath.Path

    # Create ShareGate Job Settings
    Write-Host ("   {0} - Creating Copy Settings" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"))
    $sgCopy = New-CopySettings `
        -OnWarning $job.onWarning `
        -OnError $job.onError `
        -OnSiteObjectExists $job.onSiteObjectExists `
        -OnContentItemExists $job.onContentItemExists `
        -VersionOrModerationComment $job.versionComment

    # Settings based on CSV 
    $sgTaskName = ("Box to {0} - {1}" -f $dstTech, $job.srcPath)
    Write-Host ("   {0} - Retrieving Job Settings" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"))
    if ($job.noCustomPermissions.ToLower() -eq "true") {$sgNoCustomPermissions = $true} else {$sgNoCustomPermissions = $false}

    # Create Property Template 
    Write-Host ("   {0} - Creating Property Template" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"))
    $sgProperty = New-PropertyTemplate -AuthorsAndTimestamps:$true -VersionHistory:$true -FlattenFolders:$false -Permissions:$sgNoCustomPermissions -VersionLimit $job.versionLimit
    #endregion
   
    #region 8.   Run ShareGate Job
   Write-Host ("   {0} - Running Box Folder Migration: {1}" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"), $sgTaskName)
   If ($job.dstFolder) {
        Import-BoxDocument `
            -Box $boxAuth `
            -SourceFolder $job.srcPath `
            -DestinationList $dstList `
            -DestinationFolder $job.dstFolder `
            -CopySettings $sgCopy `
            -MappingSettings $sgMappings `
            -Template $sgProperty `
            -TaskName $sgTaskName
    } else {
        Import-BoxDocument `
            -Box $boxAuth `
            -SourceFolder $job.srcPath `
            -DestinationList $dstList `
            -CopySettings $sgCopy `
            -MappingSettings $sgMappings `
            -Template $sgProperty `
            -TaskName $sgTaskName
    }
    #endregion

    #region Clearing loop variables
	Clear-Variable dstSite
	Clear-Variable dstList
    Clear-Variable sgMappings
    Clear-Variable sgCopy
    Clear-Variable sgTaskName
    Clear-Variable sgNoCustomPermissions
    Clear-Variable sgProperty
    #endregion
}
#endregion

#Cleaning up Variables
Clear-Variable boxAuth
Clear-Variable dstAuth

#endregion

#region Completed Script
# Variable Clean Up
Clear-Variable boxAdmin
Clear-Variable jobFile
Clear-Variable dstTenant
Clear-Variable dstTech
Clear-Variable inPath
Clear-Variable errorPath
Clear-Variable inFile
Clear-Variable dstRootSiteUrl

# Output Script Info
Complete-CatScript $scriptStart[0]
Clear-Variable scriptStart

# Stop logging
Stop-Transcript

# Output Script Info
Complete-CatScript $scriptStart[0]
Clear-Variable scriptStart

# Stop logging
Stop-Transcript
#endregion