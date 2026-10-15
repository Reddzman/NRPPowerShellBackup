<#  Cat.SG-OD4BProvisioning.ps1

This script Does ShareGate migrations for the Documents libary in an accounts OneDrive. This script leverages
ShareGate service account and App Registrations for authentication.

Enhancements:   Would be great to test this with MS Lists.  Right now it will only migrate the documents library.
                Could we leverage the copy site command to get MS lists?

Dependencies
    -   Source & Destination Azure App Registrations
    -   Source & Destination Service Account for ShareGate
    -	ShareGate PowerShell Modules; Doesn't support PS Core 7
    -   SharePointPnPPowerShellOnline Module
    -	..\General\Cat.PnP-GeneralFunctionsFunctions.psm1
    -   ..\Reproting\Cat.ReportingFunctions.psm1
    -	Job CSV Input file

Call script with parameters
	.\Cat.SG-OD4BMigrations.ps1 -jobFile "TTC.SG-OD4BList-ModifiedTest.csv" -userFile "TTC.SG-UserMapping-CMW-ALL.sgum" -srcTenant "charlesmachineworks" -srcAppID "5d97d419-3845-4ee0-8705-2dc9c9ae01a1" -srcCertPath "C:\MigrationFiles\TORO SharePoint Migration.pfx" -destTenant "toroco" -destAppID "73f0b78b-1545-424d-b9b8-7e73dc3fe6c3" -destCertPath "C:\AppRegCerts\CMWSharePointMigration.pfx" -skipNotModified $true -lastMigrationDate "6/9/2021"

CSV Example
	sourceUserEmail,destUserEmail
    john.doe@client1.com,john.doe@client2.com

Table of Contents
	1.	Snapin(s), Modules & Environment Connection(s)
	2.	Input & Output Files
	3.	Variables
    4.	Main Script
    5.  ODB Provisioning
    6.  Retrieving ODB Properties
    7.  Adding ShareGate Permissions
    8.  ShareGate Migration
    9.  Removing ShareGate Permissions
#>

#region Script Parameters
Param(
    [string] $jobFile = "Cat.SG-OD4BMigrations-Test.csv", # file should exist in the _Input Files directory in the CatPowerShell repo. Please don't commit
    [string] $userFile = "Cat.SG-UserMapping-NRP-WIL.sgum", # file should exist in the _Input Files directory in the CatPowerShell repo. Please don't commit
    [string] $srcTenant = "{Source Tenant Abbreviation}", # tenant name from name.sharepoint.com
    [string] $srcAppID = "{Source App Registration ID}",
    [string] $srcCertPath = "C:\Certs\SourceCertificateFileName.pfx", # Will be prompted for cert password
    [string] $destTenant = "{Destination Tenant Abbreviation}", # tenant name from name.sharepoint.com
    [string] $destAppID = "{Destination App Registration ID}",
    [string] $destCertPath =  "C:\Certs\DestinationCertificateFileName.pfx", # Will be prompted for cert password
    [bool] $skipNotModified = $false, # This will skip accounts that have not modified their content since the $lastMigrationDate
    [datetime] $lastMigrationDate = "1/1/2000", # Date of the last successful migration. This is ignored if $skipNotModified is set to False
    [bool] $basicAuthEnabled = $false # If $false, the script with leverage browser authentication with ShareGate. This uses modern authentication. The security tokens can potenitally time out.
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
Start-Transcript -Path ("..\_Log Files\{0}.SG-OD4BProvisioningMigration-{1}.log" -f $destTenant, $scriptStart[1])
#endregion

#region 1.	Snapin(s), Modules & Environment Connection(s)
# Load Sharegate modules
Import-Module Sharegate	

# Load Custom Modules
Import-Module "..\Reporting\Cat.ReportingFunctions.psm1"
#endregion

#region 2.	Input & Output Files
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

#region 3.	Variables
# Script Variables
$srcAdminUrl = ("https://{0}-admin.sharepoint.com/" -f $srcTenant)
$destAdminUrl = ("https://{0}-admin.sharepoint.com/" -f $destTenant)
$srcTenantDomain = ("{0}.onmicrosoft.com" -f $srcTenant) 
$destTenantDomain = ("{0}.onmicrosoft.com" -f $destTenant)
$srcRootSiteUrl = ("https://{0}.sharepoint.com/" -f $srcTenant)
$destRootSiteUrl = ("https://{0}.sharepoint.com/" -f $destTenant)
$siteWaitPause = 15
$siteWaitMaxCount = 40

# User Inputs
Write-Host
Write-Host "Enter Connection Information" -ForegroundColor Cyan
$srcCertPass = ConvertTo-SecureString -String (Read-Host -Prompt "   Enter certificate password for the SOURCE App Registration") -AsPlainText -Force 
Write-Host "   Enter your credentials for the SOURCE ShareGate Service Account:"
$srcSGCreds = Get-Credential -Message "Enter your credentials for the SOURCE ShareGate Service Account:"
if (-not $basicAuthEnabled) {
    Write-Host "   *** Using browser based authentication, although using modern auth, can potentially result in security token timeouts on long running scripts. ***" -ForegroundColor Yellow
    Write-Host "   Authenticating to M365"
    $srcSGConnection = Connect-Site -Url $srcRootSiteUrl -Browser -DisableSSO
}
$destCertPass = ConvertTo-SecureString -String (Read-Host -Prompt "   Enter certificate password for the DESTINATION App Registration") -AsPlainText -Force 
Write-Host "   Enter your credentials for the DESTINATION ShareGate Service Account:"
$destSGCred = Get-Credential -Message "Enter your credentials for the DESTINATION ShareGate Service Account:" 
if (-not $basicAuthEnabled) {
    Write-Host "   *** Using browser based authentication, although using modern auth, can potentially result in security token timeouts on long running scripts. ***" -ForegroundColor Yellow
    Write-Host "   Authenticating to M365"
    $destSGConnection = Connect-Site -Url $destRootSiteUrl -Browser -DisableSSO
}

# Output Parameters & Variables
Write-Host
Write-Host "Script Parameters" -ForegroundColor Cyan
Write-Host ("   Input File: {0}" -f $jobFile)
Write-Host ("   Source Tenant: {0}" -f $srcTenant)
Write-Host ("   Source App ID: {0}" -f $srcAppID)
Write-Host ("   Source Cert Path: {0}" -f $srcCertPath)
Write-Host ("   Destination Tenant: {0}" -f $destTenant)
Write-Host ("   Destination App ID: {0}" -f $destAppID)
Write-Host ("   Destination Cert Path: {0}" -f $destCertPath)
Write-Host ("   Source Root Site URL: {0}" -f $srcRootSiteUrl)
Write-Host ("   Skip Not Modified Accounts: {0}" -f $skipNotModified)
Write-Host ("   Last Successful Migration Date: {0}" -f $lastMigrationDate)
Write-Host
Write-Host "User Inputs" -ForegroundColor Cyan
Write-Host ("   Source ShareGate Service Account: {0}" -f $srcSGCreds.UserName)
Write-Host ("   Destination ShareGate Service Account: {0}" -f $destSGCred.UserName)
Write-Host
Write-Host "Script Variables" -ForegroundColor Cyan
Write-Host ("   Source Root Site URL: {0}" -f $srcRootSiteUrl)
Write-Host ("   Source SP Admin URL: {0}" -f $srcAdminUrl)
Write-Host ("   Source Tenant Domain: {0}" -f $srcTenantDomain)
Write-Host ("   Destination Root Site URL: {0}" -f $destRootSiteUrl)
Write-Host ("   Destination SP Admin URL: {0}" -f $destAdminUrl)
Write-Host ("   Destination Tenant Domain: {0}" -f $destTenantDomain)
Write-Host ("   Provisioning Wait (Seconds): {0}" -f $siteWaitPause)
Write-Host ("   Provisioing Max Attempts: {0}" -f $siteWaitMaxCount)
#endregion

#region 4.	Main Script
Write-Host
Write-Host ("Configuring & running migrations for job list: {0}" -f $jobFile) -ForegroundColor Cyan

# Connecting to the Source and Destination Tenants
Write-Host "   Connecting to Admin Sites"
Write-Host ("      Connecting to Source: {0}" -f $srcAdminUrl) -ForegroundColor DarkGray
if ($basicAuthEnabled) {
    $srcAdminSite = Connect-Site -Url $srcAdminUrl -Credential $srcSGCreds -ErrorAction Stop
} else {
    $srcAdminSite = Connect-Site -Url $srcAdminUrl -UseCredentialsFrom $srcSGConnection -ErrorAction Stop
}
Write-Host ("      Connecting to Destination: {0}" -f $destAdminUrl) -ForegroundColor DarkGray
if ($basicAuthEnabled) {
    $dstAdminSite = Connect-Site -Url $destAdminUrl -Credential $destSGCred -ErrorAction Stop
} else {
    $dstAdminSite = Connect-Site -Url $destAdminUrl -UseCredentialsFrom $destSGConnection -ErrorAction Stop
}

# Create ShareGate copy settings
Write-Host "      Creating Copy Settings"
Write-Host "         Creating Copy Settings"
$copysettings = New-CopySettings -OnContentItemExists IncrementalUpdate
Write-Host ("         Loading User Mappings file: {0}" -f $userFile) -ForegroundColor DarkGray
$copyMappings = New-MappingSettings
$copyMappings = Import-UserAndGroupMapping -MappingSettings $copyMappings -Path $userPath.Path

#region - job file iteration
ForEach ($job in $inFile){
    Write-Host
    Write-Host ("   {0} - Configuring & running migration job: {1}" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"), $job.sourceUserEmail) -ForegroundColor Magenta
    
    #region 5.  ODB Provisioning
    Write-Host ("      Querying Source OneDrive Url for: {0}" -f $job.sourceUserEmail)
    $srcOneDriveUrl = Get-OneDriveUrl -Tenant $srcAdminSite -Email $job.sourceUserEmail
    Write-Host ("      Querying Destination OneDrive Url for: {0}" -f $job.destUserEmail)
    $destOneDriveUrl = Get-OneDriveUrl -Tenant $dstAdminSite -Email $job.destUserEmail -ProvisionIfRequired

    # Wait on ODB Provisioning
    $counter = 0
    Write-Host "         Waiting for ODB Provisioning" -NoNewline -ForegroundColor DarkGray
    while (($destSiteExists -eq $false) -and $counter -lt $siteWaitMaxCount) {
        $counter++
        $destSiteExists = Test-CatPnPSiteExistsAppReg -fnUrl $destOneDriveUrl -fnAppID $destAppID -fnTenant $destTenantDomain -fnCertPath $destCertPath -fnCertPass $destCertPass
        if (-not $destSiteExists) {
            Write-Output (".") -foregroundcolor DarkGray
            start-sleep -Seconds $siteWaitPause
        }
    }
    Write-Host "." -ForegroundColor DarkGray
    #endregion

    #region 6.  Retrieving ODB Properties
    Write-Host "      Checking Source Modified Date: "  -NoNewline
    $pnpSiteConnection = Connect-PnPOnline -ClientId $srcAppID -Url $srcOneDriveUrl -Tenant $srcTenantDomain -CertificatePath $srcCertPath -CertificatePassword $srcCertPass -WarningAction SilentlyContinue
    $currentList = Get-PnPList "Documents"
    Write-Host $currentList.LastItemModifiedDate
    Disconnect-PnPOnline -Connection $pnpSiteConnection
    #endregion

    # deteriming if ODB should migrated
    If ($skipNotModified -and ($currentList.LastItemModifiedDate -le $lastMigrationDate)) {
        Write-Host ("      Skipping OneDrive: {0}" -f $job.sourceUserEmail) -ForegroundColor Yellow
    } Else {
        #region 7.  Adding ShareGate Permissions
        Write-Host ("      Setting Source OneDrive Permissions")
        $srcOD4BSiteConnection = Connect-PnPOnline -ClientId $srcAppID -Url $srcOneDriveUrl -Tenant $srcTenantDomain -CertificatePath $srcCertPath -CertificatePassword $srcCertPass -WarningAction SilentlyContinue    
        Add-PnPSiteCollectionAdmin -Owners $srcSGCreds.UserName -Connection $srcOD4BSiteConnection
        Disconnect-PnPOnline -Connection $srcOD4BSiteConnection -WarningAction SilentlyContinue

        # Adding ShareGate Account as an administrator to Destination
        Write-Host ("      Setting Destination OneDrive Permissions")
        $destOD4BSiteConnection = Connect-PnPOnline -ClientId $destAppID -Url $destOneDriveUrl -Tenant $destTenantDomain -CertificatePath $destCertPath -CertificatePassword $destCertPass -WarningAction SilentlyContinue    
        Add-PnPSiteCollectionAdmin -Owners $destSGCred.UserName -Connection $destOD4BSiteConnection
        Disconnect-PnPOnline -Connection $destOD4BSiteConnection -WarningAction SilentlyContinue
        #endregion
                
        #region 8.  ShareGate Migration
        Write-Host ("      Connecting to Source OD4B Account: {0}" -f $job.sourceUserEmail) 
        if ($basicAuthEnabled) {
            $srcSite = Connect-Site -Url $srcOneDriveUrl -Credential $srcSGCreds -ErrorAction Stop
        } else {
            $srcSite = Connect-Site -Url $srcOneDriveUrl -UseCredentialsFrom $srcSGConnection -ErrorAction Stop
        }
        Write-Host ("      Connecting to Destination OD4B Account: {0}" -f $job.destUserEmail)
        if ($basicAuthEnabled) {
            $dstSite = Connect-Site -Url $destOneDriveUrl -Credential $destSGCred -ErrorAction Stop
        } else {
            $dstSite = Connect-Site -Url $destOneDriveUrl -UseCredentialsFrom $destSGConnection -ErrorAction Stop
        }

        # Loading Document Libiries
        $jobTaskName = ("OneDrive Migration - {0} to {1}" -f $job.sourceUserEmail, $job.destUserEmail)
        Write-Host ("      {0}" -f $jobTaskName)
        Write-Host ("         Loading Documents Libraries")
        $srcList = Get-List -Site $srcSite -Name "Documents"
        $dstList = Get-List -Site $dstSite -Name "Documents"

        # Running Migration Job
        $now = Get-Date
        Write-Host ("         Migration Started: {0}" -f $now)
        Copy-Content -SourceList $srcList -DestinationList $dstList -TaskName $jobTaskName -CopySettings $copySettings -MappingSettings $copyMappings -WaitForImportCompletion 
        $now = Get-Date
        Write-Host ("         Migration Compeleted: {0}" -f $now) -foregroundcolor Green

        # Clearing ShareGate Objects
        Clear-Variable srcSite
        Clear-Variable dstSite
        Clear-Variable srcList
        Clear-Variable dstList    
        #endregion

        #region 9.   Removing ShareGate Permissions
        Write-Host ("      Removing Source OneDrive Permissions")
        $srcOD4BSiteConnection = Connect-PnPOnline -ClientId $srcAppID -Url $srcOneDriveUrl -Tenant $srcTenantDomain -CertificatePath $srcCertPath -CertificatePassword $srcCertPass -WarningAction SilentlyContinue    
        Remove-PnPSiteCollectionAdmin -Owners $srcSGCreds.UserName -Connection $srcOD4BSiteConnection        
        Disconnect-PnPOnline -Connection $srcOD4BSiteConnection -WarningAction SilentlyContinue

        # Removing ShareGate Account as an administrator to Destination
        Write-Host ("      Removing Destination OneDrive Permissions")
        $destOD4BSiteConnection = Connect-PnPOnline -ClientId $destAppID -Url $destOneDriveUrl -Tenant $destTenantDomain -CertificatePath $destCertPath -CertificatePassword $destCertPass -WarningAction SilentlyContinue    
        Remove-PnPSiteCollectionAdmin -Owners $destSGCred.UserName -Connection $destOD4BSiteConnection
        Disconnect-PnPOnline -Connection $destOD4BSiteConnection -WarningAction SilentlyContinue
        #endregion
    }
}

Write-Host
Write-Host ("OneDrive Provisioning and Migration completed.") -ForegroundColor Green
Write-Host ("Please refer to the ShareGate Job Logs for any warnings or errors that require remediation") -ForegroundColor Yellow
#endregion     

#region Completed Script
# Output Script Info
Complete-CatScript $scriptStart[0]

# Stop logging
Stop-Transcript
#endregion