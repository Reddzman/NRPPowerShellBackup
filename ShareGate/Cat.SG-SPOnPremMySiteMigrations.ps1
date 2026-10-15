<#  Cat.SG-SPOnPremMySiteMigrations.ps1

This script automates ShareGate SharePoint on-premises to SharePoint Online migrations my sites document libraries.  If the destination OneDrive
does not exist they will be provisioned. It is currenlty leveraging an App Registration in the destination Tenant to perform configuration. ShareGate still requires a service account even if using the 
Azure ShareGate Desktop Application to improve performance.

Dependencies
    -   Destination Azure App Registrations (Catapult Powershell & ShareGate)
    -   Source & Destination Service Account for ShareGate
    -	ShareGate PowerShell Modules; Doesn't support PS Core 7 - Install-Module PnP.PowerShell -RequiredVersion 1.12 -Force
    -   PnP.PowerShell Module
    -	..\General\Cat.GeneralFunctions.psm1
    -   ..\Reporting\Cat.ReportingFunctions.psm1
    -	Input CSV File in Input Files folder
    -   User mapping file export from ShareGate

Call script with parameters
    cd C:\GitHub\CatPowerShell\ShareGate\
	.\Cat.SG-SPOnPremMySiteMigrations.ps1 -jobFile "Samples\ShareGate\Cat.SG-SPOnPremMySiteMigrations-Client.csv" -userFile "samples\ShareGate\Cat.SG-UserMapping-Client.sgum" -destTenant "nrpdev2" -destAppID "caf6b11a-cf03-416c-b0f0-67a09a8e30ee" -destCertPath "C:\GitHub\NRPDev\AppRegistration\PowerShell App Registration.pfx"

CSV Example
    
CSV Notes
    Settings columns are based on options for New-CopySettings & Copy-List ShareGate Documentation
        https://support-desktop.sharegate.com/hc/en-us/articles/115000597407-New-Copy-Settings
        https://support-desktop.sharegate.com/hc/en-us/articles/115000597727-Copy-List

Table of Contents
	1.	Snapin(s), Modules & Environment Connection(s)
	2.	Input & Output Files
	3.	Variables
    4.	Main Script
    5.  Connecting to Destination Admin Tenant
    6.  OneDrive Provisioning
    7.  Adding ShareGate Permissions
    8.  Configuring ShareGate job settings
    9.  Running ShareGate Migration
    10. Removing ShareGate Permissions
#>

#region Script Parameters
Param(
    [String] $jobFile = "Cat.SG-SPOnPremMySiteMigrations-Client.csv",
    [String] $userFile = "Cat.SG-UserMapping-CLIENT.sgum",
    [string] $destTenant = "{Destination Tenant Name}",
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
Start-Transcript -Path ("..\_Log Files\{0}.SG-SPOnPremMySiteMigrations-{1}.log" -f $destTenant, $scriptStart[1])
#endregion

#region 1.	Snapin(s), Modules & Environment Connection(s)
# Load MS / PnP PowerShell Modules
Get-Module -ListAvailable | Where-Object{($_.Name -like "*SharePoint*") -or ($_.Name -like "*PnP*") } | Remove-Module
Import-Module PnP.PowerShell

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
$destAdminUrl = ("https://{0}-admin.sharepoint.com/" -f $destTenant)
$destRootSiteUrl = ("https://{0}.sharepoint.com/" -f $destTenant)
$destTenantDomain = ("{0}.onmicrosoft.com" -f $destTenant)
$destCert = Get-PnPAzureCertificate -CertificatePath $destCertPath -CertificatePassword $destCertPass
$siteWaitMaxCount = 20
$siteWaitPause = 30 # Seconds

# User Inputs
Write-Host
Write-Host "Enter Connection Information" -ForegroundColor Cyan
Write-Host "   Enter ShareGate credentials for the SOURCE SharePoint environment:"
$srcSGCreds = Get-Credential -Message "Enter ShareGate credentials for the SOURCE environment:" 
$destCertPass = ConvertTo-SecureString -String (Read-Host -Prompt "   Enter DESTINATION certificate password for the app registration") -AsPlainText -Force
Write-Host "   Enter ShareGate credentials for the DESTINATION SharePoint environment:"
$destSGCreds = Get-Credential -Message "Enter ShareGate credentials for the DESTINATION environment:" 
if (-not $basicAuthEnabled) {
    Write-Host "   *** Using browser based authentication, although using modern auth, can potentially result in security token timeouts on long running scripts. ***" -ForegroundColor Yellow
    Write-Host "   Authenticating to M365"
    $destSGConnection = Connect-Site -Url $destRootSiteUrl -Browser -DisableSSO
}

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
Write-Host ("   Source Credentials: {0}" -f $srcSGCreds.UserName)
Write-Host ("   Destination Credentials: {0}" -f $destSGCreds.UserName)
Write-Host
Write-Host "Script Variables" -ForegroundColor Cyan
Write-Host ("   Destination Admin Url: {0}" -f $destAdminUrl)
Write-Host ("   Destination Root Site Url: {0}" -f $destRootSiteUrl)
Write-Host ("   Destination Tenant Url: {0}" -f $destTenantDomain)
Write-Host ("   Site Creation Wait Max Iterations: {0}" -f $siteWaitMaxCount)
Write-Host ("   Site Creation Wait Pause Duration: {0}" -f $siteWaitPause)
#endregion
    
#region 4.	Main Script
#region 5.  Connecting to Destination Admin Tenant
Write-Host
Write-Host ("Configuring & running migrations for job list: {0}" -f $jobFile) -ForegroundColor Cyan
Write-Host ("   {0} - Connecting to Destination Admin Tenant: {1}" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"), $destAdminUrl)
if ($basicAuthEnabled) {
            $dstAdminSite = Connect-Site -Url $destAdminUrl -Credential $destSGCreds -ErrorAction Stop
        } else {
            $dstAdminSite = Connect-Site -Url $destAdminUrl -UseCredentialsFrom $destSGConnection -ErrorAction Stop
        }
#endregion

# Looping through input file
:JobLoop ForEach ($job in $inFile){
    Write-Host
    Write-Host ("   {0} - Configuring & running migration job: {1}" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"), $job.destEmail) -ForegroundColor Magenta
    <# Output Job Configuration for Troubleshooting (if Needed remove "<" on this line)
    $job #>

    #region 6.  OneDrive Provisioning
    Write-Host ("      {0} - Confirming OneDrive Provisioning Complete: {1}" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"), $job.destEmail) -NoNewline
    try {
        $destOneDriveUrl = $null
        $destOneDriveUrl = Get-OneDriveUrl -Tenant $dstAdminSite -Email $job.destEmail -ProvisionIfRequired
    }
    catch {
        Write-Host "." -NoNewline
    }

    # Confirm OneDrive Provisioning
    $counter = 0
    $destSiteExists = $false
    while (($destSiteExists -eq $false) -and $counter -lt $siteWaitMaxCount) {
        $counter++
        $destSiteExists = Test-CatPnPSiteExistsAppReg -fnUrl $destOneDriveUrl -fnAppID $destAppID -fnTenant $destTenantDomain -fnCertPath $destCertPath -fnCertPass $destCertPass
        if (-not $destSiteExists) {
            Write-Host "." -NoNewline
            start-sleep -Seconds $siteWaitPause
        }
    }
    Write-Host "."
    #endregion #>

    # Run Migrations if destination site exists
    If ($destSiteExists){
        #region 7.  Adding ShareGate Permissions
        Write-Host ("      {0} - Adding Service Account as Destination Site Owner: {1}" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"), $destSGCreds.UserName)
        $destSiteConnection = Connect-PnPOnline -ClientId $destAppID -Url $destOneDriveUrl -Tenant $destTenantDomain -Thumbprint $destCert.Thumbprint
        Add-PnPSiteCollectionAdmin -Owners $destSGCreds.UserName -Connection $destSiteConnection
        #endregion #>

        #region 8.  Create Destination Folder
        if ($job.destFolder) {
            $destFolder = ("{0}/{1}" -f $job.destList, $job.destFolder)
            Write-Host ("      {0} - Checking if folder exists in destination: {1}" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"), $destFolder) -NoNewline
            try {
                $tempFolder = Get-PnPFolder -Url $destFolder -ErrorAction Stop
                Write-Host " - Folder Already Exists" -ForegroundColor Green
            }
            catch {
                Write-Host " - Folder Does not Exists" -ForegroundColor Yellow
            }
            if (-not $tempfolder) {
                Write-Host ("      {0} - Creating Folder in Destination: {0}" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"), $destFolder)
                $tempFolder = Add-PnPFolder -Name $job.destFolder -Folder $job.destList
            }
            Clear-Variable tempFolder
        }
        Disconnect-PnPOnline -Connection $destSiteConnection
        Clear-Variable destSiteConnection
        #endregion #>
                
        #region 9.  Configuring ShareGate job settings
        # Loading sites
        Write-Host ("      {0} - Loading Source My Site: {1}" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"), $job.srcUsername)
        $sgSourceSite = Connect-Site -Url $job.srcSiteURL -Credential $srcSGCreds
        $sgSourceList = Get-List -Site $sgSourceSite -Name $job.srcList
        Write-Host ("      {0} - Loading Destination OneDrive: {1}" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"), $job.destEmail)
        $sgDestSite = Connect-Site -Url $destOneDriveUrl -Credential $destSGCreds
        if ($basicAuthEnabled) {
            $sgDestSite = Connect-Site -Url $destOneDriveUrl -Credential $destSGCreds -ErrorAction Stop
        } else {
            $sgDestSite = Connect-Site -Url $destOneDriveUrl -UseCredentialsFrom $destSGConnection -ErrorAction Stop
        }
        $sgDestList = Get-List -Site $sgDestSite -Name $job.destList
        
        # Create ShareGate Mapping settings
        Write-Host ("      {0} - Creating Copy Settings" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"))
        Write-Host ("         Loading User Mappings file: {0}" -f $userFile) -ForegroundColor DarkGray
        $sgCopyMappings = New-MappingSettings
        $sgCopyMappings = Import-UserAndGroupMapping -MappingSettings $sgCopyMappings -Path $userPath.Path

        # Create ShareGate Property Template
        Write-Host "         Creating Property Template" -ForegroundColor DarkGray
        if ($job.userAlerts.ToLower() -eq "true") {$sgUserAlerts = $true} else {$sgUserAlerts = $false}
        if ($job.noCustomPermissions.ToLower() -eq "true") {$sgCustomPermissions = $false} else {$sgCustomPermissions = $true}
        $sgPropertyTemplate = New-PropertyTemplate `
            -CheckInAs "SameAsCurrent" `
            -ContentApproval "SameAsCurrent" `
            -AuthorsAndTimestamps:$true `
            -VersionHistory:$true `
            -Permissions:$sgCustomPermissions `
            -WebParts:$false `
            -NoLinkCorrection:$false `
            -FlattenFolders:$false `
            -VersionLimit $job.versionLimit `
            -UserAlerts:$sgUserAlerts

        # Setting ShareGate Job Settings
        Write-Host ("         Configuring job Settings") -ForegroundColor DarkGray
        if ($job.waitForImportCompletion.ToLower() -eq "true") {$sgWaitForCompletion = $true} else {$sgWaitForCompletion = $false}
        $sgTaskName = ("My Site Migration - {0} - {1}/{2}" -f $job.srcUsername, $job.destEmail, $job.destFolder)
        $sgCopySettings = New-CopySettings `
            -OnWarning $job.onWarning `
            -OnError $job.onError `
            -OnSiteObjectExists $job.onSiteObjectExists `
            -OnContentItemExists $job.onContentItemExists `
            -VersionOrModerationComment $job.versionComment
        #endregion #>

        #region 10.  Running ShareGate Migration
        Write-Host ("      {0} - Running Job: {1}" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"), $sgTaskName)
        $sgResult = Copy-Content `
            -SourceList $sgSourceList `
            -DestinationList $sgDestList `
            -DestinationFolder $job.destFolder `
            -WaitForImportCompletion:$sgWaitForCompletion `
            -Template $sgPropertyTemplate `
            -CopySettings $sgCopySettings `
            -MappingSettings $sgCopyMappings `
            -TaskName $sgTaskName

        # Clearing ShareGate Objects
        Clear-Variable sgWaitForCompletion
        Clear-Variable sgCustomPermissions
        Clear-Variable sgUserAlerts
        Clear-Variable sgSourceSite
        Clear-Variable sgSourceList
        Clear-Variable sgDestSite
        Clear-Variable sgDestList
        Clear-Variable sgPropertyTemplate
        Clear-Variable sgCopyMappings
        Clear-Variable sgTaskName
        Clear-Variable sgCopySettings
        Clear-Variable sgResult
        #endregion #>

        #region 11. Removing ShareGate Permissions
        Write-Host ("      {0} - Removing Service Account as Destination Site Owner: {1}" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"), $destSGCreds.UserName)
        $destSiteConnection = Connect-PnPOnline -ClientId $destAppID -Url $destOneDriveUrl -Tenant $destTenantDomain -Thumbprint $destCert.Thumbprint
        Get-PnPSiteCollectionAdmin  -Connection $destSiteConnection | Where-Object -Property Email -eq $destSGCreds.UserName | Remove-PnPSiteCollectionAdmin -Connection $destSiteConnection
        Disconnect-PnPOnline -Connection $destSiteConnection -WarningAction SilentlyContinue
        Clear-Variable destSiteConnection
        #endregion #>
        
        # Log complete
        Write-Host ("      Completed Successfully") -foregroundcolor "green"       
    }
    Else {
        Write-Host ("      OneDrive not Provisioned: {0}" -f $job.destEmail) -ForegroundColor "red"
        Continue JobLoop
    }
    Clear-Variable destSiteExists
    Clear-Variable destOneDriveUrl
}

# Script complete output
Write-Host
Write-Host ("{0} - My Site Migrations completed. " -f (Get-Date -Format "yy-MM-dd HH:mm:ss")) -ForegroundColor Green
Write-Host "Please refer to the ShareGate Job Logs for warnings/errors that require remediation" -ForegroundColor Yellow
#endregion

#region Completed Script
# Output Script Info
Complete-CatScript $scriptStart[0]

# Stop logging
Stop-Transcript
#endregion