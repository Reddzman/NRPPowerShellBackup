<#  Cat.SG-NetworkShareMigrations.ps1

This script automates ShareGate network share on-premises to Teams / SharePoint Online migrations at a folder level.  Destination teams/sites must exist in advance. Use the
"ShareGate\Cat.SG-NetworkShareMigrations.ps1" script to configure teams/sites in the tenant. It is currently leveraging an App Registration in the destination Tenant to 
perform configuration. ShareGate still requires a service account even if using the Azure ShareGate Desktop Application to improve performance. You can use the 
"_Input Files\Samples\ShareGate\Cat.SG-NetworkShareMigrations-Test.csv" file to validate the client setup if you update the srcUNCPath to a source that exists on their network.

Dependencies
    -   Destination Azure App Registrations (Catapult Powershell & ShareGate)
    -   Source & Destination Service Account for ShareGate
    -	ShareGate PowerShell Modules; Doesn't support PS Core 7
    -   PnP.PowerShell Module - Install-Module PnP.PowerShell -RequiredVersion 1.12 -Force
    -	..\General\Cat.GeneralFunctions.psm1
    -   ..\Reporting\Cat.ReportingFunctions.psm1
    -	Input CSV File in Input Files folder - _Input Files\Samples\ShareGate\Cat.SG-NetworkShareMigrations-Test.csv
    -   User mapping file export from ShareGate - _Input Files\Samples\ShareGate\Cat.SG-UserMapping-Client.sgum

Call script with parameters
    cd C:\Quisitive\PowerShell\ShareGate

    # Teams & SPO Command
	.\Cat.SG-NetworkShareMigrations.ps1 -jobFile "meritage.SG-NetworkShareMigrations-Test.csv" -userFile "Cat.SG-UserMapping-merTEST.sgum" -destTenant "mymeritage" -destAppID "aa73d0ee-1a51-41e7-91c4-d17a65c9531c" -destCertPath "..\AppRegistration\mermig.pfx" -destTeamsManagedPath "sites" -logSite "https://mymeritage.sharepoint.com/sites/2025FileShareMigration" -logFolder "migrationlogs" -containsODB $false -basicAuthEnabled $false
    
    # Contains ODB command
    .\Cat.SG-NetworkShareMigrations.ps1 -jobFile meritage.SG-NetworkShareMigrations-Test.csv" -userFile "Cat.SG-UserMapping-merTEST.sgum" -destTenant "mymeritage" -destAppID "aa73d0ee-1a51-41e7-91c4-d17a65c9531c" -destCertPath "..\AppRegistration\mermig.pfx" -destTeamsManagedPath "sites" -logSite "https://mymeritage.sharepoint.com/sites/2025FileShareMigration" -logFolder "migrationlogss" -containsODB $true -basicAuthEnabled $false

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
    5.  Configure Destination Site URL
    6.  ShareGate Settings & Migration
    7.  Configuring ShareGate job settings
    8.  Running ShareGate Migration
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
    [bool] $containsODB = $false,
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
Start-Transcript -Path ("..\_Log Files\{0}.NetworkShareMigrations-{1}.log" -f $destTenant, $scriptStart[1])

# Disabling PNP Upgrade checks
$env:PNPPOWERSHELL_UPDATECHECK="false" # Disable Update Check
#endregion

#region 1.	Snapin(s), Modules & Environment Connection(s)
# Load modules
#Import-Module PnP.PowerShell
Import-Module ShareGate

# Load Custom Modules
Import-Module "..\Reporting\Cat.ReportingFunctions.psm1"
#endregion

#region 2.	Input & Output Files
# File Paths
$inPath = ("..\_Input Files\{0}" -f $jobFile)
$userPath = ("..\_Input FIles\{0}" -f $userFile)
$errorPath = ("..\_Output Files\{0}.SG-ListMigrationErrors-{1}.log" -f $destTenant, $scriptStart[1])

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
$destMyUrl = ("https://{0}-my.sharepoint.com/" -f $destTenant)
$destRootSiteUrl = ("https://{0}.sharepoint.com/" -f $destTenant)
$destTenantDomain = ("{0}.onmicrosoft.com" -f $destTenant)
$siteWaitMaxCount = 20
$siteWaitPause = 30 # Seconds
$specialCharacters = "[~``!@#$%^&*():_]"

# User Inputs
Write-Host
Write-Host "Enter Connection Information" -ForegroundColor Cyan
$destCertPass = ConvertTo-SecureString -String (Read-Host -Prompt "   Enter DESTINATION certificate password for the app registration") -AsPlainText -Force
Write-Host "   Enter ShareGate credentials for the DESTINATION SharePoint environment:"
$destSGCreds = Get-Credential -Message "Enter ShareGate credentials for the DESTINATION environment:"   
if (-not $basicAuthEnabled) {
    Write-Host "   *** Using browser based authentication, although using modern auth, can potentially result in security token timeouts on long running scripts. ***" -ForegroundColor Yellow
    Write-Host "   Authenticating to M365"
    $rootSGConnection = Connect-Site -Url $destRootSiteUrl -Browser -DisableSSO
    if ($containsODB) {
        $mySGConnection = Connect-Site -Url $destMyUrl -Browser -DisableSSO
        $admSGConnection = Connect-Site -Url $destAdminUrl -Browser -DisableSSO
    }
}
$destCert = Get-PnPAzureCertificate -CertificatePath $destCertPath -CertificatePassword $destCertPass
 
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
Write-Host ("   Destination Credentials: {0}" -f $destSGCreds.UserName)
#endregion
    
#region 4.	Main Script
Write-Host
Write-Host ("Configuring & running migrations for job list: {0}" -f $jobFile) -ForegroundColor Cyan

# Looping through input file
:JobLoop ForEach ($job in $inFile){
    Write-Host
    Write-Host ("   {0} - Configuring & running migration job - Site: {1} List: {2}" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"), $job.destSiteTitle, $job.destListName) -ForegroundColor Magenta
    <# Output Job Configuration for Troubleshooting (if Needed remove "<" on this line)
    $job
    start-sleep -Seconds $siteWaitPause #>

    #region 5.  Configure Destination Site URL
    # Classic Vs. Modern Sites
    if ($job.destSiteURL -like "https://*") {$destURL = $job.destSiteURL.Trim()} else {$destURL = ("{0}/{1}" -f $job.classicSubSiteParentURL, $job.destSiteURL)}
    
    # Change $destURL if shared or private channel    
    switch ($job.destChannelType) {
        private {$destURL = ("{0}{1}/{2}-{3}" -f $destRootSiteUrl, $destTeamsManagedPath, ($job.destSiteTitle -replace " ",""), ($job.destChannel -replace " ",""))}
        #private {$destURL = ("{0}-{1}" -f $job.destSiteURL, ($job.destChannel -replace " ",""))}
        shared {$destURL = ("{0}{1}/{2}-{3}" -f $destRootSiteUrl, $destTeamsManagedPath, ($job.destSiteTitle -replace " ",""), ($job.destChannel -replace " ",""))}
    }

    # change $destURL if OneDrive
    if ($job.siteTemplate.ToLower() -eq "onedrive") {
        $sgAdmSite = Connect-Site -Url $destAdminUrl -UseCredentialsFrom $admSGConnection -ErrorAction Stop
        $destURL = Get-OneDriveUrl -Tenant $sgAdmSite -Email $job.Email
        Remove-Variable sgAdmSite
    }
    #endregion #>

    #region 6.  ShareGate Settings & Migration
    Write-Host ("      Checking if Destination Site Exists: {0}" -f $destURL) -NoNewline
    $destSiteExists = Test-CatPnPSiteExistsAppReg -fnUrl $destURL -fnAppID $destAppID -fnTenant $destTenantDomain -fnCertPath $destCertPath -fnCertPass $destCertPass
    $destSiteExists

    If ($destSiteExists){
        Write-Host (" - Destination Site Exists") -Foregroundcolor DarkGray
   
        #region 7.  Configuring ShareGate job settings
        #region Loading Destination site
        Write-Host ("      Loading Destination Site: {0}" -f $destURL)
        try {
            if ($job.siteTemplate.ToLower() -eq "onedrive") {
                $sgDestSite = Connect-Site -Url $destURL -UseCredentialsFrom $mySGConnection -ErrorAction Stop
            } else {
                $sgDestSite = Connect-Site -Url $destURL -UseCredentialsFrom $rootSGConnection -ErrorAction Stop
            }
        }
        catch {
            # Output Error Information
            Write-Host ("         {0} - Error Loading Destination Site: {1}" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"),$_.Exception) -ForegroundColor Red
            Write-CatErrorCSV -funcJob $job -funcAction "Loading Destination Site" -funcException $_.Exception -funcDetails $_.ErrorDetails -funcStackTrace $_.ScriptStackTrace -funcLogPath $errorPath
            Continue JobLoop
        }
        #endregion

        #region Loading Destination List
        Write-Host ("         Loading Destination List: {0}" -f $job.destListName) -ForegroundColor DarkGray
        try {
            $sgDestList = Get-List -Site $sgDestSite -Name $job.destListName
        }
        catch {
            # Output Error Information
            Write-Host ("         {0} - Error Loading Destination List: {1}" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"),$_.Exception) -ForegroundColor Red
            Write-CatErrorCSV -funcJob $job -funcAction "Loading Destination List" -funcException $_.Exception -funcDetails $_.ErrorDetails -funcStackTrace $_.ScriptStackTrace -funcLogPath $errorPath
            Continue JobLoop
        }
        #endregion

        #region Determining Destination Folder
        Write-Host "         Determining Destination Folder" -ForegroundColor DarkGray
        If ((($job.destFolder.ToLower() -ne "none") -and ($job.destFolder -ne "")) -and (($job.destChannel.ToLower() -ne "none") -and ($job.destChannel -ne "") -and ($job.destChannel.ToLower() -ne "na"))) { # Both Dest older & Channel
            switch ($job.destChannelType) {
                Standard {$sgDestFolder = ("{0}/{1}" -f $job.destChannel, $job.destFolder)}
                #private {$sgDestFolder = ("{0}/{1}" -f $job.destChannel, $job.destFolder)}
                private {$destURL = $sgDestFolder = $job.destFolder}
                shared {$destURL = $sgDestFolder = $job.destFolder}
            }
        } ElseIf (($job.destFolder.ToLower() -ne "none") -and ($job.destFolder -ne "") -and (($job.destChannel.ToLower() -eq "none") -or ($job.destChannel -eq "") -or ($job.destChannel.ToLower() -eq "na"))) { # Dest Folder only
            $sgDestFolder = $job.destFolder
        } elseif (($job.destFolder.ToLower() -eq "none") -or ($job.destFolder -eq "") -and (($job.destChannel.ToLower() -ne "none") -and ($job.destChannel -ne "") -and ($job.destChannel.ToLower() -ne "na"))) { # channel only
            switch ($job.destChannelType) {
                Standard {$sgDestFolder = ("{0}/{1}" -f $job.destChannel, $job.destFolder)}
                private {$sgDestFolder = ("{0}/{1}" -f $job.destChannel, $job.destFolder)}
                shared {$destURL = $sgDestFolder = $job.destFolder}
            }
        } Else { # niether, root of library
            $sgDestFolder = ""
        }
        #endregion
        
        #region ShareGate Settings
        try {
            # Create ShareGate Mapping settings
            Write-Host "      Creating Copy Mappings"
            Write-Host ("         Loading User Mappings file: {0}" -f $userFile) -ForegroundColor DarkGray
            $sgCopyMappings= New-MappingSettings
            $sgCopyMappings= Import-UserAndGroupMapping -MappingSettings $sgCopyMappings -Path $userPath.Path

            # Create ShareGate Job Settings
            Write-Host ("         Creating Copy Settings") -ForegroundColor DarkGray
            $sgTaskName = ("{0} - {1}" -f $job.srcTitle, $job.destSiteTitle)
            $sgCopySettings = New-CopySettings `
                -OnWarning $job.onWarning `
                -OnError $job.onError `
                -OnSiteObjectExists $job.onSiteObjectExists `
                -OnContentItemExists $job.onContentItemExists

            # Settings based on CSV 
            if ($job.noCustomPermissions.ToLower() -eq "true") {$sgNoCustomPermissions = $false} else {$sgNoCustomPermissions = $true}
            if ($job.waitForCompletion.ToLower() -eq "true") {$sgWaitForCompletion = $true} else {$sgWaitForCompletion = $false}
            if ($job.normalMode.ToLower() -eq "true") {$sgNormalMode = $true} else {$sgNormalMode = $false}
            if ($job.filterStartDate.ToLower() -ne "none") {
                $startDate = [DateTime]::ParseExact($job.filterStartDate, "MMddyyyy", $null) 
                
                # Create ShareGate Property Template with date filter
                Write-Host "      Creating Property Template"
                $sgPropertyTemplate = New-PropertyTemplate `
                        -AuthorsAndTimestamps:$true `
                        -Permissions:$sgNoCustomPermissions `
                        -NoLinkCorrection:$false `
                        -FlattenFolders:$false `
                        -From $startDate
            } else {
                # Create ShareGate Property Template without date filter
                $sgPropertyTemplate = New-PropertyTemplate `
                        -AuthorsAndTimestamps:$true `
                        -Permissions:$sgNoCustomPermissions `
                        -NoLinkCorrection:$false `
                        -FlattenFolders:$false `
            }
        }
        catch {
            # Output Error Information
            Write-Host ("         {0} - Error Configuring ShareGate Settings: {1}" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"),$_.Exception) -ForegroundColor Red
            Write-CatErrorCSV -funcJob $job -funcAction "Configuring ShareGate Settings" -funcException $_.Exception -funcDetails $_.ErrorDetails -funcStackTrace $_.ScriptStackTrace -funcLogPath $errorPath
            Continue JobLoop
        }
        #endregion
        #endregion

        #region 8.  Running ShareGate Migration
        try {
            Write-Host ("      Running Migration Job: {0}" -f $sgTaskName)
            $result = Import-Document -TaskName $sgTaskName `
                -SourceFolder $job.srcUNCPath `
                -DestinationList $sgDestList `
                -DestinationFolder $sgDestFolder `
                -NormalMode:$sgNormalMode `
                -WaitForImportCompletion $sgWaitForCompletion `
                -MappingSettings $sgCopyMappings `
                -CopySettings $sgCopySettings `
                -Template $sgPropertyTemplate
        }
        catch {
            # Output Error Information
            Write-Host ("         {0} - Error Running ShareGate Job: {1}" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"),$_.Exception) -ForegroundColor Red
            Write-CatErrorCSV -funcJob $job -funcAction "Running ShareGate Job" -funcException $_.Exception -funcDetails $_.ErrorDetails -funcStackTrace $_.ScriptStackTrace -funcLogPath $errorPath
            Continue JobLoop
        }

        try {
            # Exporting ShareGate Logs
            $logPath = (".\\Temp\\{0}-{1}.xlsx" -f ($sgTaskName -replace $specialCharacters),(Get-Date -Format "yyMMddHHmmss"))
            Write-Host ("         Exporting Job Log: {0}" -f $logPath) -ForegroundColor DarkGray
            $temp = Export-Report $result -Path $logPath -ErrorAction Stop

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
        #endregion #>

        #region Clearing Variables
        Clear-Variable sgDestSite
        Clear-Variable sgDestList
        if ($sgDestFolder) {Clear-Variable sgDestFolder}        
        Clear-Variable sgCopySettings
        Clear-Variable sgCopyMappings
        if ($startDate) {Clear-Variable startDate} 
        Clear-Variable sgPropertyTemplate
        #endregion

        # Log complete
        Write-Host ("      Completed Successfully") -foregroundcolor "green"       
    }
    Else {        
        # Output Error Information
        Write-CatErrorCSV -funcJob $job -funcAction "Connecting to Destination Site" -funcException "Initial Site Connection Failed" -funcDetails "Connecting to Destination Site" -funcStackTrace "Connecting to Destination Site" -funcLogPath $errorPath
        Continue JobLoop
    }
    #endregion
}

#endregion

#region Completed Script
# Script complete output
Write-Host
Write-Host "SharePoint Migration completed. " -ForegroundColor Green
Write-Host "Please refer to the ShareGate Job Logs for warnings/errors that require remediation" -ForegroundColor Yellow

# Output Script Info
Complete-CatScript $scriptStart[0]

# Stop logging
Stop-Transcript
#endregion