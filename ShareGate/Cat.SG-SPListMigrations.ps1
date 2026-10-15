<#  Q.SG-SPListMigrations_MigrateOnly.ps1

This script automates ShareGate SharePoint on-premises or SharePoint Online to SharePoint Online migrations at a list level.  Destination sites 
must exist in advance. If the destination lists do not exist they will be created. It is currently leveraging an App Registration in the 
destination Tenant to perform configuration. ShareGate still requires a service account even if using the Azure ShareGate Desktop Application 
to improve performance.

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
    cd C:\Quisitive\PowerShell\ShareGate
	.\Cat.SG-SPListMigrations.ps1 -jobFile "Samples\ShareGate\Cat.SG-SPListMigrations-SPO.csv" -userFile "Samples\ShareGate\Cat.SG-UserMapping-Client.sgum" -destTenant "nrpdev2" -destAppID "caf6b11a-cf03-416c-b0f0-67a09a8e30ee" -destCertPath "C:\GitHub\NRPDev\AppRegistration\PowerShell App Registration.pfx" -destTeamsManagedPath "teams" -logSite "https://nrpdev2.sharepoint.com/sites/M365Admins" -logFolder "Shared Documents/Migration/Logs" -basicAuthEnabled $false

CSV Example
    
CSV Notes
    Settings columns are based on options for New-CopySettings & Copy-List ShareGate Documentation
        https://support-desktop.sharegate.com/hc/en-us/articles/115000597407-New-Copy-Settings
        https://support-desktop.sharegate.com/hc/en-us/articles/115000597727-Copy-List

Table of Contents
	1.	Snapin(s), Modules & Environment Connection(s)
	2.	Input & Output Files
	3.	Variables
    4.  Functions
    5.	Main Script
    6.  Configure Destination Site URL
    7.  Adding ShareGate Permission
    8. Configuring ShareGate job
    9. Running ShareGate Migration
#>

#region Script Parameters
Param(
    [String] $jobFile = "Q.SG-SPOnPremListMigrations-Client.csv",
    [String] $userFile = "Q.Client-UserMap.sgum",
    [string] $destTenant = "{Destination Tenant Name}",
    [String] $destAppID = "{Azure App Registration ID}",
    [String] $destCertPath = "C:\Quisitive\PowerShell\App Reg\ClientCertFileName.pfx",
    [String] $destTeamsManagedPath = "sites", # The managed path used by SharePoint site behind Microsoft Teams. IE "sites" or "teams"
    [String] $logSite = "{URL to site}", # SharePoint URL for site/team you want to upload log files to
    [String] $logFolder = "{Folder to upload files}", # Document Library/Channel/Folder
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
Start-Transcript -Path ("..\_Log Files\{0}.SG-SPListMigrations-{1}.log" -f $destTenant, $scriptStart[1])

# Disabling PNP Upgrade checks
$ProgressPreference = "SilentlyContinue"
$env:PNPPOWERSHELL_UPDATECHECK="false" # Disable Update Check
#endregion

#region 1.	Snapin(s), Modules & Environment Connection(s)
# Load modules
Import-Module Sharegate	

# Load Custom Modules
Import-Module "..\Reporting\Cat.ReportingFunctions.psm1"
#endregion

#region 2.	Input & Output Files
# File Paths
$inPath = ("..\_Input Files\{0}" -f $jobFile)
$userPath = ("..\_Input FIles\{0}" -f $userFile)
$errorPath = ("..\_Output Files\{0}.SG-SPListMigration-Errors-{1}.log" -f $destTenant, $scriptStart[1])

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
$siteWaitMaxCount = 20
$siteWaitPause = 30 # Seconds
$specialCharacters = "[~``!@#$%^&*():_]"

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
$destCert = Get-PnPAzureCertificate -Path $destCertPath -Password $destCertPass
 
# Variable Output
Write-Host
Write-Host "Script Parameters" -ForegroundColor Cyan
Write-Host ("   Job File: {0}" -f $jobFile)
Write-Host ("   User Mapping File: {0}" -f $userFile)
Write-Host ("   Destination Tenant: {0}" -f $destTenant)
Write-Host ("   Destination App ID: {0}" -f $destAppID)
Write-Host ("   Destination Cert Path: {0}" -f $destCertPath)
Write-Host ("   Destination Teams Managed Path: {0}" -f $destTeamsManagedPath)
Write-Host ("   Basic Auth Enabled: {0}" -f $basicAuthEnabled)
Write-Host
Write-Host "Script Variables" -ForegroundColor Cyan
Write-Host ("   Destination Admin Url: {0}" -f $destAdminUrl)
Write-Host ("   Destination Root Site Url: {0}" -f $destRootSiteUrl)
Write-Host ("   Destination Tenant Url: {0}" -f $destTenantDomain)
Write-Host ("   Site Creation Wait Max Iterations: {0}" -f $siteWaitMaxCount)
Write-Host ("   Site Creation Wait Pause Duration: {0}" -f $siteWaitPause)
Write-Host ("   Special Characters {0}" -f $specialCharacters)
Write-Host
Write-Host "User Inputs" -ForegroundColor Cyan
Write-Host ("   Source Credentials: {0}" -f $srcSGCreds.UserName)
Write-Host ("   Destination Credentials: {0}" -f $destSGCreds.UserName)
#endregion
    
#region 5.	Main Script
Write-Host
Write-Host ("Configuring & running migrations for job list: {0}" -f $jobFile) -ForegroundColor Cyan

# Looping through input file
:JobLoop ForEach ($job in $inFile){
    Write-Host
    Write-Host ("   {0} - Configuring & running migration job - Site: {1} List: {2}" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"), $job.destSiteTitle, $job.destListName) -ForegroundColor Magenta
    <# Output Job Configuration for Troubleshooting (if Needed remove "<" on this line)
    $job
    start-sleep -Seconds $siteWaitPause #>

    #region 6.  Configure Destination Site URL
    # Classic Vs. Modern Sites
    if ($job.destSiteURL -like "https://*") {$destURL = $job.destSiteURL.Trim()} else {$destURL = ("{0}/{1}" -f $job.classicSubSiteParentURL, $job.destSiteURL)}
    
    # Change $destURL if shared or private channel    
    switch ($job.destChannelType) {
        private {$destURL = ("{0}-{1}" -f $job.destSiteURL, ($job.destChannel -replace " ",""))}
        shared {$destURL = ("{0}{1}/{2}-{3}" -f $destRootSiteUrl, $destTeamsManagedPath, ($job.destSiteTitle -replace " ",""), ($job.destChannel -replace " ",""))}
    }
    #endregion #>

    #region 7. Run Migrations if destination site exists 
    Write-Host ("      Checking if Destination Site Exists: {0}" -f $destURL) -NoNewline
    $destSiteExists = Test-CatPnPSiteExistsAppReg -fnUrl $destURL -fnAppID $destAppID -fnTenant $destTenantDomain -fnCertPath $destCertPath -fnCertPass $destCertPass

    If ($destSiteExists){
        Write-Host (" - Destination Site Exists") -Foregroundcolor DarkGray
   
        #region 8. Configuring ShareGate job settings
        #region Loading Source site
        Write-Host ("      Loading Source Site: {0}" -f $job.srcTitle)
        try {
            $sgSourceSite = Connect-Site -Url $job.srcSiteURL -Credential $srcSGCreds -ErrorAction Stop
        }
        catch {
            # Output Error Information
            Write-CatErrorCSV -funcCSVItem $job -funcAction "Loading Source Site" -funcException $_.Exception -funcDetails $_.ErrorDetails -funcStackTrace $_.ScriptStackTrace -funcFilePath $errorPath
            Continue JobLoop
        }
        #endregion

        #region Loading Destination site
        Write-Host ("      Loading Destination Site: {0}" -f $destURL)
        try {
            if ($basicAuthEnabled) {
                $sgDestSite = Connect-Site -Url $destURL -Credential $destSGCreds -ErrorAction Stop
            } else {
                $sgDestSite = Connect-Site -Url $destURL -UseCredentialsFrom $destSGConnection -ErrorAction Stop
            }
        }
        catch {
            # Output Error Information
            Write-Host ("         {0} - Error Loading Destination Site: {1}" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"),$_.Exception) -ForegroundColor Red
            Log-MigrationError -funcJob $job -funcAction "Loading Destination Site" -funcException $_.Exception -funcDetails $_.ErrorDetails -funcStackTrace $_.ScriptStackTrace -funcLogPath $errorPath
            Continue JobLoop
        }
        #endregion
        
        #region Shared Settings
        # Create ShareGate Mapping settings
        Write-Host "      Creating Copy Settings"
        Write-Host ("         Loading User Mappings file: {0}" -f $userFile) -ForegroundColor DarkGray
        $sgCopyMappings= New-MappingSettings
        $sgCopyMappings= Import-UserAndGroupMapping -MappingSettings $sgCopyMappings -Path $userPath.Path

        # Create ShareGate Job Settings
        Write-Host ("         Creating Job Settings") -ForegroundColor DarkGray
        $sgTaskName = ("{0} - {1}" -f $job.srcTitle, $job.destSiteTitle)
        $sgCopySettings = New-CopySettings `
            -OnWarning $job.onWarning `
            -OnError $job.onError `
            -OnSiteObjectExists $job.onSiteObjectExists `
            -OnContentItemExists $job.onContentItemExists `
            -VersionOrModerationComment $job.versionComment

        # Shared Settings based on CSV 
        if ($job.noCustomPermissions.ToLower() -eq "true") {$sgNoCustomPermissions = $true} else {$sgNoCustomPermissions = $false}
        if ($job.userAlerts.ToLower() -eq "true") {$sgUserAlerts = $true} else {$sgUserAlerts = $false}
        if ($job.waitForCompletion.ToLower() -eq "true") {$sgWaitForCompletion = $true} else {$sgWaitForCompletion = $false}
        #endregion

        #region List Vs. Content Settings
        #If ($job.destFolder -or $job.destChannel) {
        If ((($job.destFolder.ToLower() -ne "none") -and ($job.destFolder -ne "")) -or (($job.destChannel.ToLower() -ne "none") -and ($job.destChannel -ne "") -and ($job.destChannel.ToLower() -ne "na"))) {
            #region Loading Lists
            Write-Host ("         Loading Source List: {0}" -f $job.srcListName) -ForegroundColor DarkGray
            try {
                $sgSourceList = Get-List -Site $sgSourceSite -Name $job.srcListName  
            }
            catch {
                # Output Error Information
                Log-MigrationError -funcJob $job -funcAction "Loading Source List" -funcException $_.Exception -funcDetails $_.ErrorDetails -funcStackTrace $_.ScriptStackTrace -funcLogPath $errorPath
                Continue JobLoop
            }   
            Write-Host ("         Loading Destination List: {0}" -f $job.destListName) -ForegroundColor DarkGray
            try {

                if(($job.destChannel.ToLower() -ne "none") -and ($job.destChannel -ne "") -and ($job.destChannel.ToLower() -ne "na")) {
                    # Get Destination List in Channel
                    $sgDestList = Get-List -Site $sgDestSite -Name "Documents"
                } else {
                    # Get Destination List
                    $sgDestList = Get-List -Site $sgDestSite -Name $job.destListName
                }
                
            }
            catch {
                # Output Error Information
                Write-Host ("         {0} - Error Loading Destination List: {1}" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"),$_.Exception) -ForegroundColor Red
                Log-MigrationError -funcJob $job -funcAction "Loading Destination List" -funcException $_.Exception -funcDetails $_.ErrorDetails -funcStackTrace $_.ScriptStackTrace -funcLogPath $errorPath
                Continue JobLoop
            }
            #endregion

            # Settings for Copy Content
            Write-Host "         Creating Property Template" -ForegroundColor DarkGray
            if ($job.noWebParts.ToLower() -eq "true") {$sgNoWebParts = $false} else {$sgNoWebParts = $true} # Opposit values due to how command works
            $sgPropertyTemplate = New-PropertyTemplate `
                -CheckInAs "SameAsCurrent" `
                -ContentApproval "SameAsCurrent" `
                -AuthorsAndTimestamps:$true `
                -VersionHistory:$true `
                -Permissions:$sgNoCustomPermissions `
                -WebParts:$sgNoWebParts `
                -NoLinkCorrection:$false `
                -FlattenFolders:$false `
                -VersionLimit $job.versionLimit `
                -UserAlerts:$sgUserAlerts
        } Else {
            # List Settings based on CSV
            if ($job.noWebParts.ToLower() -eq "true") {$sgNoWebParts = $true} else {$sgNoWebParts = $false}
            if ($job.noContent.ToLower() -eq "true") {$sgNoContent = $true} else {$sgNoContent = $false}
            if ($job.noWorkflow.ToLower() -eq "true") {$sgNoWorkflow = $true} else {$sgNoWorkflow = $false}
            if ($job.noSiteFeatures.ToLower() -eq "true") {$sgNoSiteFeatures = $true} else {$sgNoSiteFeatures = $false}
            if ($job.noCustomizedListForms.ToLower() -eq "true") {$sgNoCustomizedListForms = $true} else {$sgNoCustomizedListForms = $false}
            if ($job.noListViews.ToLower() -eq "true") {$sgNoListViews= $true} else {$sgNoListViews = $false}
            if ($job.noNintexWorkflowHistory.ToLower() -eq "true") {$sgNoNintexWorkflowHistory = $true} else {$sgNoNintexWorkflowHistory = $false}
            if ($job.normalMode.ToLower() -eq "true") {$sgNormalMode = $true} else {$sgNormalMode = $false}
            if ($job.forceNewListExperience.ToLower() -eq "true") {$sgForceNewListExperience = $false} else {$sgForceNewListExperience = $true}
        }
        #endregion        
        #endregion #>

        #region 9. Running ShareGate Migration        
        #If ($job.destFolder -or $job.destChannel) {
        If ((($job.destFolder.ToLower() -ne "none") -and ($job.destFolder -ne "")) -or (($job.destChannel.ToLower() -ne "none") -and ($job.destChannel -ne "") -and ($job.destChannel.ToLower() -ne "na"))) {
            #region Running Content Job
            switch ($job.destChannelType) {
                Standard {
                    if($job.destFolder) {
                        # Migrate to the channel's destination folder
                        #$sgDestFolder = ("{0}/{1}/{2}" -f $job.destListPath, $job.destChannel, $job.destFolder) 
                        $sgDestFolder = ("{0}/{1}" -f $job.destChannel, $job.destFolder) 
                    } else {
                        $sgDestFolder = $job.destChannel
                    }
                }
                Private {
                    if($job.destFolder) {
                        # Migrate to the channel's destination folder
                        #$sgDestFolder = ("{0}/{1}/{2}" -f $job.destListPath, $job.destChannel, $job.destFolder)
                        $sgDestFolder = ("{0}/{1}" -f $job.destChannel, $job.destFolder)
                    } else {
                        $sgDestFolder = $job.destChannel
                    }                    
                }
                Shared {
                    if($job.destFolder) {
                        # Migrate to the channel's destination folder
                        $sgDestFolder = $job.destFolder
                    } else {
                        $sgDestFolder = ""
                    }                     
                }
                Default {
                    $sgDestFolder = $job.destFolder
                }
            }   
            Write-Host ("      Running Content Job: {0}" -f $sgTaskName)
            try {
                $sgResult = Copy-Content `
                    -SourceList $sgSourceList `
                    -DestinationList $sgDestList `
                    -DestinationFolder $sgDestFolder `
                    -WaitForImportCompletion:$sgWaitForCompletion `
                    -Template $sgPropertyTemplate `
                    -CopySettings $sgCopySettings `
                    -MappingSettings $sgCopyMappings `
                    -TaskName $sgTaskName `
                    -ErrorAction Stop
            }
            catch {
                # Output Error Information
                Write-CatErrorCSV -funcCSVItem $job -funcAction "Running Content Job" -funcException $_.Exception -funcDetails $_.ErrorDetails -funcStackTrace $_.ScriptStackTrace -funcFilePath $errorPath
            }
            #endregion
        } Else {
            #region Running List Job
            Write-Host ("      Running List Job: {0}" -f $sgTaskName)
            try {
                $sgResult = Copy-List `
                    -Name $job.srcListName `
                    -SourceSite $sgSourceSite `
                    -DestinationSite $sgDestSite `
                    -ListTitle $job.destListName `
                    -ListTitleUrlSegment $job.destListPath `
                    -NoContent:$sgNoContent `
                    -VersionLimit $job.versionLimit `
                    -NoCustomPermissions:$sgNoCustomPermissions `
                    -NoWorkflows:$sgNoWorkflow `
                    -NoSiteFeatures:$sgNoSiteFeatures `
                    -NoWebParts:$sgNoWebParts `
                    -NoCustomizedListForms:$sgNoCustomizedListForms `
                    -NoListViews:$sgNoListViews `
                    -NoNintexWorkflowHistory:$sgNoNintexWorkflowHistory `
                    -UserAlerts:$sgUserAlerts `
                    -NormalMode:$sgNormalMode `
                    -WaitForImportCompletion:$sgWaitForCompletion `
                    -ForceNewListExperience:$sgForceNewListExperience `
                    -CopySettings $sgCopySettings `
                    -MappingSettings $sgCopyMappings `
                    -TaskName $sgTaskName `
                    -ErrorAction Stop
            }
            catch {
                # Output Error Information
                Write-CatErrorCSV -funcCSVItem $job -funcAction "Running List Job" -funcException $_.Exception -funcDetails $_.ErrorDetails -funcStackTrace $_.ScriptStackTrace -funcFilePath $errorPath
            }
            #endregion
        }

        # Clearing ShareGate Objects
        Clear-Variable sgSourceSite
        Clear-Variable sgDestSite
        Clear-Variable sgCopySettings
        Clear-Variable sgCopyMappings
        If ($job.destFolder) {
            Clear-Variable sgSourceList
            Clear-Variable sgDestList
            Clear-Variable sgPropertyTemplate
        }
        #endregion #>
        
        #region 10. Exporting ShareGate Logs
        try {
            # Exporting ShareGate Logs
            $logPath = (".\\Temp\\{0}-{1}.xlsx" -f ($sgTaskName -replace $specialCharacters),(Get-Date -Format "yyMMddHHmmss"))
            Write-Host ("         Exporting Job Log: {0}" -f $logPath) -ForegroundColor DarkGray
            $temp = Export-Report $sgResult -Path $logPath -ErrorAction Stop

            # Uploading Log File to SharePoint
            Write-Host ("         Uploading Log to SharePoint: {0}" -f $logPath) -ForegroundColor DarkGray
            $destConnection = Connect-PnPOnline -ClientId $destAppID -Url $logSite -Tenant $destTenantDomain -Thumbprint $destCert.Thumbprint
            $temp = Add-PnPFile -Path $logPath -Folder $logFolder -ErrorAction Stop
            Disconnect-PnPOnline

            # Deleting Local File
            Write-Host ("         Deleting Local File: {0}" -f $sgTaskName) -ForegroundColor DarkGray
            Remove-Item -Path $logPath -Force
        }
        catch {
            # Output Error Information
            Write-Host ("         {0} - Error Exporting Logs: {1}" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"),$_.Exception) -ForegroundColor Red
            Write-CatErrorCSV -funcJob $job -funcAction "Error Exporting Logs" -funcException $_.Exception -funcDetails $_.ErrorDetails -funcStackTrace $_.ScriptStackTrace -funcLogPath $errorPath
            Continue JobLoop
        }
        #endregion

        # Log complete
        Write-Host ("      Completed Successfully") -foregroundcolor "green"       
    }
    Else {        
        # Output Error Information
        Write-CatErrorCSV -funcCSVItem $job -funcAction "Connecting to Destination Site" -funcException $_.Exception -funcDetails $_.ErrorDetails -funcStackTrace $_.ScriptStackTrace -funcFilePath $errorPath
        Continue JobLoop
    }
    #endregion
}

# Script complete output
Write-Host
Write-Host "SharePoint Migration completed. " -ForegroundColor Green
Write-Host "Please refer to the ShareGate Job Logs for warnings/errors that require remediation" -ForegroundColor Yellow
#endregion

#region Completed Script
# Output Script Info
Complete-CatScript $scriptStart[0]

# Stop logging
Stop-Transcript
#endregion