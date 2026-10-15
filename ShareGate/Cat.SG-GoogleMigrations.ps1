<#  Cat.SG-FileMigrations.ps1

This script automates ShareGate File Share/Google Drive migration to SharePoint, Teams and OneDrive document libraries.

Dependencies
    -   Destination Azure App Registrations
    -   Destination Service Account for ShareGate
    -	ShareGate PowerShell Modules; Doesn't support PS Core 7
    -   SharePointPnPPowerShellOnline Module
    -	..\General\Cat.GeneralFunctions.psm1
    -   ..\Reporting\Cat.ReportingFunctions.psm1
    -	Input CSV File in Input Files folder
    -   User mapping file export from ShareGate

Call script with parameters
	.\Cat.SG-FileMigrations.ps1 -destTenant "tenantname" -destAppID = "7f0fd640-e811-44cc-a19f-e9bcd0e28d95" -destCertPath = "C:\Temp\appReg\PnPAppRegPowerShell.pfx" -inputFile "Cat.SG-FileShareMigrations-Client.csv" -userFile Cat.SG-UserMapping-Test.sgum

CSV Example
sourceType,googleDriveAccount,sourceFilePath,destinationURL,migrationType,documentLibrary,subFolder,createDestFolders
Google,neilwilgus@gmail.com,Migration,https://wilguscatdev.sharepoint.com/sites/ContosoITTeamSite,FULL,GoogleDrive,,FALSE
FileShare,,C:\Temp\migrate\,https://wilguscatdev-my.sharepoint.com/personal/neilwilgus_wilguscatdev_onmicrosoft_com/,FULL,Documents,Folder1\SubFolder1\SubFolder2,TRUE
FileShare,,C:\Temp\migrate\,https://wilguscatdev.sharepoint.com/sites/ContosoITTeamSite,Incremental,FileShare,SubFolder,FALSE
FileShare,,C:\Temp\migrate\,https://wilguscatdev.sharepoint.com/sites/ContosoITTeamSite,Incremental,NoFolders,,FALSE

CSV Notes
    The source can be either File Shares or Google Drive 
    The destination site can be either a SharePoint, Teams or OneDrive for Business site. 
    The script will assign migration account SC Admin permissions before migration and revoke after.
    When migrating from G Suite, be sure you grant access before running these scripts
    https://support-desktop.sharegate.com/hc/en-us/articles/360037473391
    Any Google Drive source MUST be a G-Suite member and not just a standard google drive user

Table of Contents
	1.	Snapin(s), Modules & Environment Connection(s)
	2.	Input & Output Files
	3.	Variables
    4.	Main Script
    5.  Adding ShareGate Permissions
    6.  Create Destination Folder Structure
    7.  Configuring ShareGate job
    8.  Running ShareGate Migration
    9.  Removing ShareGate Permissions
#>

#region Script Parameters
Param(
    [String] $jobFile = "Cat.SG-ModernSiteList-CLIENT.csv",
    [String] $userFile = "Cat.SG-UserMapping-CLIENT.sgum",
    [string] $destTenant = "{DestinationTenant}",
    [String] $destAppID = "{Azure App Registration ID}",
    [String] $destCertPath = "C:\certs\PnPAppRegPowerShell.pfx",
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
Start-Transcript -Path ("..\_Log Files\{0}.Cat.SG-FileShareMigrations-{1}.log" -f $tenant, $scriptStart[1])
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
# User Inputs
Write-Host
$destCertPass = ConvertTo-SecureString -String (Read-Host -Prompt "   Enter DESTINATION certificate password for the app registration") -AsPlainText -Force
Write-Host "   Enter ShareGate credentials for the DESTINATION SharePoint environment:"
$destSGCreds = Get-Credential -Message "Enter ShareGate credentials for the DESTINATION environment:" 
if ($inFile -match "google") {
    $googleDriveAdmin = Read-Host "   Enter Google Drive Admin email address:"
    $googleDrive = Connect-GoogleDrive -Email $googleDriveAdmin -Admin
}
if (-not $basicAuthEnabled) {
    Write-Host "*** Using browser based authentication, although using modern auth, can potentially result in security token timeouts on long running scripts. ***" -ForegroundColor Yellow
    Write-Host "   Authenticating to M365"
    $destSGConnection = Connect-Site -Url $destRootSiteUrl -Browser -DisableSSO
}

# Output Parameters & Variables
# Script Variables
$destAdminUrl = ("https://{0}-admin.sharepoint.com/" -f $destTenant)
$destRootSiteUrl = ("https://{0}.sharepoint.com/" -f $destTenant)
$destTenantDomain = ("{0}.onmicrosoft.com" -f $destTenant) 

# Variable Output
Write-Host
Write-Host "Script Parameters" -ForegroundColor Cyan
Write-Host ("   Job File: {0}" -f $jobFile)
Write-Host ("   User Mapping File: {0}" -f $userFile)
Write-Host ("   Destination Tenant: {0}" -f $destTenant)
Write-Host ("   Destination App ID: {0}" -f $destAppID)
Write-Host ("   Destination Cert Path: {0}" -f $destCertPath)
Write-Host
Write-Host "User Inputs" -ForegroundColor Cyan
Write-Host ("   Destination Credentials: {0}" -f $destSGCreds.UserName)
Write-Host
Write-Host "Script Variables" -ForegroundColor Cyan
Write-Host ("   Destination Admin Url: {0}" -f $destAdminUrl)
Write-Host ("   Destination Root Site Url: {0}" -f $destRootSiteUrl)
Write-Host ("   Destination Tenant Url: {0}" -f $destTenantDomain)
#endregion
    
#region 4.	Main Script
Write-Host
Write-Host ("Configuring & running migrations for job list: {0}" -f $jobFile) -ForegroundColor Cyan

# Looping through input file
:JobLoop ForEach ($job in $inFile){
    
    $destURL = $job.DestinationUrl
    $srcFolderPath = $job.sourceFilePath
    $destSubFolders = $job.subFolder -split "\\"
    $createDestFolders = $job.createDestFolders
    $docLibName = $job.documentLibrary

    Write-Host
    Write-Host ("   Configuring & running migration job: {0}" -f $job.sourceFilePath) -ForegroundColor Magenta
    
    # Determine if file path exists
    if ($job.sourceType -eq "FileShare") {
        if ((Test-Path -Path $job.sourceFilePath) -eq $false) {
            Write-Host ("         Invalid Source File Path specified: {0} - Migration Canceled" -f $job.sourceFilePath) -ForegroundColor Red
            Continue JobLoop
        }
    }

    # Determine if destination site exists
    Write-Host ("      Checking if Site Exists: {0}" -f $destURL)
    $destSiteExists = Test-CatPnPSiteExistsAppReg -fnUrl $destURL -fnAppID $destAppID -fnTenant $destTenantDomain -fnCertPath $destCertPath -fnCertPass $destCertPass

    # Run Migrations if destination site exists
    If ($destSiteExists){
        # Connecting to Destination Site
        $destSiteConnection = Connect-PnPOnline -ClientId $destAppID -Url $destURL -Tenant $destTenantDomain -CertificatePath $destCertPath -CertificatePassword $destCertPass -WarningAction SilentlyContinue #>
        $destSite = Get-PnPSite -Includes RootWeb    
        
        #region 5.  Adding ShareGate Permissions
        # Adding ShareGate service account to destination site
        Write-Host ("      Adding Service Account as Destination Site Owner: {0}" -f $destSGCreds.UserName)
        $destSiteConnection = Connect-PnPOnline -ClientId $destAppID -Url $destURL -Tenant $destTenantDomain -CertificatePath $destCertPath -CertificatePassword $destCertPass -WarningAction SilentlyContinue #>
        Add-PnPSiteCollectionAdmin -Owners $destSGCreds.UserName -Connection $destSiteConnection
        Disconnect-PnPOnline -Connection $destSiteConnection -WarningAction SilentlyContinue
        #endregion
                
        #region 6.  Create Destination Folder Structure
        $folder = $null
        $subFolders = $null
        if ($createDestFolders -eq "TRUE") {
            Write-Host ("      Creating folder structure in: {0}" -f $docLibName)
            $destSiteConnection = Connect-PnPOnline -ClientId $destAppID -Url $destURL -Tenant $destTenantDomain -CertificatePath $destCertPath -CertificatePassword $destCertPass -WarningAction SilentlyContinue #>            
            $i = 0
            

            foreach ($folder in $destSubFolders) {                
                if  ($i -eq 0) {
                    Write-Host ("      Creating folder: {0}" -f $folder) -ForegroundColor Cyan
                    try {
                        Add-PnPFolder -Name $folder -Folder $docLibName -ErrorAction Stop -WarningAction SilentlyContinue
                        Write-Host ("      Folder Created: {0}" -f $folder) -ForegroundColor Green
                    }
                    catch {
                        $ex = Get-PnPException
                        if ($ex.Message -match "already exists") {
                            Write-Host ("      Folder already exists: {0}" -f $folder)
                        }
                        else {
                            Write-Host $ex.Message -ForegroundColor Red
                        }                        
                    }
                }
                else {
                    $subFolders = $prevFolder + "\" + $folder
                    Write-Host ("      Creating sub folder: {0}" -f $subFolders) -ForegroundColor Cyan
                    #for some reason this line below reports an exception (file not found) but still creates the folders??
                    try {
                        Add-PnPFolder -Name $subFolders -Folder $docLibName -ErrorAction Stop -WarningAction SilentlyContinue
                        Write-Host ("      Sub Folder Created: {0}" -f $subFolders) -ForegroundColor Green
                    }
                    catch {
                        $ex = Get-PnPException
                        if ($ex.Message -match "already exists") {
                            Write-Host ("      Folder already exists: {0}" -f $subFolders)
                        }
                        else {
                            Write-Host $ex.Message -ForegroundColor Red
                        }                        
                    }                        
                } 
                if ($i -eq 0) {
                    $prevFolder += $folder
                }
                else {
                    $prevFolder += "\" + $folder
                }
                $i += 1
            }             
            $prevFolder = $null
            Disconnect-PnPOnline -Connection $destSiteConnection -WarningAction SilentlyContinue
        }
        #endregion
        
        #region 7.  Configuring ShareGate job
        # Loading site
        Write-Host ("      Loading Destination Site: {0}" -f $destURL)
        $destSite = Connect-Site -Url $destURL -Credential $destSGCreds
        if ($basicAuthEnabled) {
            $destSite = Connect-Site -Url $destURL -Credential $destSGCreds -ErrorAction Stop
        } else {
            $destSite = Connect-Site -Url $destURL -UseCredentialsFrom $destSGConnection -ErrorAction Stop
        }
        
        # Create ShareGate copy settings
        Write-Host "      Creating Copy Settings"
        Write-Host ("         Loading User Mappings file: {0}" -f $userFile) -ForegroundColor DarkGray
        $copyMappings = New-MappingSettings
        $copyMappings = Import-UserAndGroupMapping -MappingSettings $copyMappings -Path $userPath.Path

        # Determining Migration Type
        $migrationType = $job.migrationType.Trim()
        $migrationType = $migrationType.ToLower()
        switch ($migrationType) {
            full {
                Write-Host ("         Configuring a FULL migration") -ForegroundColor DarkGray
                $copySettings = New-CopySettings -OnSiteObjectExists Merge -OnContentItemExists Overwrite
            }
            incremental {
                Write-Host ("         Configuring a INCREMENTAL migration") -ForegroundColor DarkGray
                $copySettings = New-CopySettings -OnSiteObjectExists Merge -OnContentItemExists IncrementalUpdate
            }
            Default {
                Write-Host ("         Invalid Migration Type specified: {0} - Migration Canceled" -f $migrationType) -ForegroundColor Red
                Continue JobLoop
            }
        }       
       
        Write-Host ("         Getting Destination List:{0}" -f $job.documentLibrary) -ForegroundColor DarkGray
        $destList = Get-List -Site $destSite -name $job.documentLibrary
        #endregion

        #region 8.  Running ShareGate Migration
        if ($job.sourceType -eq "FileShare") {
            # Setting Job Name
            Write-Host ("         Configuring job name") -ForegroundColor DarkGray
            $taskName = ("File Share Migration - {0} - {1}" -f $srcFolderPath, $migrationType)
            Write-Host ("      Running Job: {0}" -f $taskName)
            $result = Import-Document -TaskName $taskName `
                                -DestinationList $destList `
                                -CopySettings $copySettings `
                                -MappingSettings $copyMappings `
                                -SourceFolder $srcFolderPath `
                                -DestinationFolder $job.subFolder
        }
        elseif ($job.sourceType -eq "Google") {
            #TODO this code still needs testing
            # Setting Job Name
            Write-Host ("         Configuring job name") -ForegroundColor DarkGray
            $taskName = ("Google Drive Migration - {0} - {1}" -f $job.googleDriveAccount, $migrationType)            

            Write-Host ("      Running Job: {0}" -f $taskName)
            $result = Import-GoogleDriveDocument -GoogleDrive $googleDrive `
                                -DestinationList $destList `
                                -SourceFolder $job.sourceFilePath `
                                -SourceView "MyDocuments" `
                                -CopySettings $copySettings `
                                -MappingSettings $copyMappings `
                                -DestinationFolder $job.subFolder
        }
        else {
            Write-Host ("         Invalid Source Type in CSV:{0}" -f $job.sourceType) -ForegroundColor Red
        }
        If ($result.Errors -gt 0) {
            Write-Host ("Migration Errors Found - Please review the ShareGate Task:{0}" -f $taskName) -ForegroundColor Red
        }
        
        # Clearing ShareGate Objects
        Clear-Variable destSite
        #endregion

        #region 9.  Removing ShareGate Permissions
        # Removing ShareGate service account to destination site
        Write-Host ("      Removing Service Account as Destination Site Owner: {0}" -f $destSGCreds.UserName)
        $destSiteConnection = Connect-PnPOnline -ClientId $destAppID -Url $destURL -Tenant $destTenantDomain -CertificatePath $destCertPath -CertificatePassword $destCertPass -WarningAction SilentlyContinue
        Remove-PnPSiteCollectionAdmin -Owners $destSGCreds.UserName -Connection $destSiteConnection
        Disconnect-PnPOnline -Connection $destSiteConnection -WarningAction SilentlyContinue
        #endregion
        
        # Log complete
        Write-Host ("      Completed Successfully: {0}" -f $destURL) -foregroundcolor "green"       
    }
    Else {
        Write-Host ("      Destination Site does not exist: {0}" -f $destURL) -ForegroundColor "red"
        Continue JobLoop
    }
}

# Script complete output
Write-Host
Write-Host ("File Migration is completed. Please refer to the ShareGate Job Logs for any warnings or errors that require remediation") -ForegroundColor Green
#endregion

#region Completed Script
# Output Script Info
Complete-CatScript $scriptStart[0]

# Stop logging
Stop-Transcript
#endregion