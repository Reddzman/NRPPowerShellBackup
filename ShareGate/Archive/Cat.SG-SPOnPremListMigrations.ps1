<#  Cat.SG-SPOnPremListMigrations.ps1

This script automates ShareGate SharePoint on-premises to SharePoint Online migrations at a list level.  Destination sites must exist in advance. If the destination lists
do not exist they will be created. It is currenlty leveraging an App Registration in the destination Tenant to perform configuration. ShareGate still requires a service account even if using the 
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
	.\Cat.SG-SPOnPremListMigrations.ps1 -jobFile "QMigTest-Delirium-01.csv" -userFile "QMigTest-UserMap.sgum" -destTenant "sanfordhealth" -destAppID "6bd4f453-6253-4f7a-be0c-717ecd464321" -destCertPath "C:\Quisitive\PowerShell\App Reg\SharePointMigration2024.pfx" -timeZone 11 -siteLocale 1033 -basicAuthEnabled $false

CSV Example
    
CSV Notes
    Settings columns are based on options for New-CopySettings & Copy-List ShareGate Documentation
        https://support-desktop.sharegate.com/hc/en-us/articles/115000597407-New-Copy-Settings
        https://support-desktop.sharegate.com/hc/en-us/articles/115000597727-Copy-List

Table of Contents
	1.	Input & Output Files
	2.	Variables
	3.	Snapin(s), Modules & Environment Connection(s)
    4.  Functions
    5.	Main Script
    6.  Site Creation
    7.  Configuring Hubs
    8.  Associating to Hub
    9.  Adding ShareGate Permissions
    10.  Create Destination Folder
    11. Configuring ShareGate job
    12. Running ShareGate Migration
    13. Removing ShareGate Permissions
#>

#region Script Parameters
Param(
    [String] $jobFile = "QMigTest-Delirium-13.csv",
    [String] $userFile = "QMigTest-UserMap.sgum",
    [string] $destTenant = "sanfordhealth",
    [String] $destAppID = "6bd4f453-6253-4f7a-be0c-717ecd464321",
    [String] $destCertPath = "C:\Quisitive\PowerShell\App Reg\SharePointMigration2024.pfx",
    [String] $timeZone = 12, # 13 = Pacific Time (US & Canada); 38 = Arizona; 12 = Mountain Time (US & Canada); 11 = Central Time (US & Canada); 10 = Eastern Time (US & Canada)
    [String] $siteLocale = 1033, # 1033 = English - United States
    [bool] $basicAuthEnabled = $false, # If $false, the script with leverage browser authentication with ShareGate. This uses modern authentication. The security tokens can potenitally time out.
    [bool] $createChannels = $true, # If $true, the script will create team channels based on the library names for team sites in the $jobFile based on the channelType column.
    [bool] $runForceCheckIn = $false, # Set to $false to not check in any files at the source site.
    [bool] $forceCheckInOnPrem = $true, # Set to $false to run force check in on a SharePoint online source. Requires $runForceCheckIn = $true
    [string] $pastDate = "6/17/2024",    # Date you want to ignore source library files for check in. Leave Blank to check in ALL files that are currently checked out.
    [bool] $useCustomSiteTemplates = $true, # If true create and store the job's custom xml site templates in the input files folder (i.e. Cat.SG-TeamSiteTemplate.xml, Cat.SG-CommunicationSiteTemplate.xml)
    [String] $teamSiteTemplateFile = "Cat.SG-TeamSiteTemplate.xml",
    [String] $commSiteTemplateFile = "Cat.SG-CommunicationSiteTemplate.xml",
    [String] $msteamTemplateFile = "Cat.SG-SiteTemplate-Test.xml"
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
Start-Transcript -Path ("..\_Log Files\{0}.SG-SPOnPremListMigrations-{1}.log" -f $destTenant, $scriptStart[1])
#endregion

#region 1.	Input & Output Files
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

#region 2.	Variables
# Script Variables
$destAdminUrl = ("https://{0}-admin.sharepoint.com/" -f $destTenant)
$destRootSiteUrl = ("https://{0}.sharepoint.com/" -f $destTenant)
$destTenantDomain = ("{0}.onmicrosoft.com" -f $destTenant)
$siteWaitMaxCount = 20
$siteWaitPause = 30 # Seconds
$specialCharacters = "[~``!@#$%^&*()_-{|?/\<>:'}]"

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
$destCert = Get-PnPAzureCertificate -CertificatePath $destCertPath -CertificatePassword $destCertPass
 
# Variable Output
Write-Host
Write-Host "Script Parameters" -ForegroundColor Cyan
Write-Host ("   Job File: {0}" -f $jobFile)
Write-Host ("   User Mapping File: {0}" -f $userFile)
Write-Host ("   Destination Tenant: {0}" -f $destTenant)
Write-Host ("   Destination App ID: {0}" -f $destAppID)
Write-Host ("   Destination Cert Path: {0}" -f $destCertPath)
Write-Host ("   Time Zone: {0}" -f $timeZone)
Write-Host ("   Site Locale: {0}" -f $siteLocale)
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

#region 3.	Snapin(s), Modules & Environment Connection(s)
# Load Sharegate modules
Import-Module Sharegate	

# Load Custom Modules
Import-Module "..\Reporting\Cat.ReportingFunctions.psm1"
#endregion

#region 4. Functions
<#  Log-MigrationError
This function creates a log file for errors that happen throughout the migration process.

Command: Log-MigrationError -funcJob $job -funcAction "Creating Site" -funcException $_.Exception -funcDetails $_.ErrorDetails -funcStackTrace $_.ScriptStackTrace -funcLogPath $errorPath

Parameters
    funcJob = Job information from Migraiton CSV file
    funcAction = Action that resulted in error
    funcException = Error Exception
    funcDetails = Error Details
    funcStackTrace = Error Script Stack Trace
    funcLogPath = Path to log file
#>
function Log-MigrationError
{ 
	param ( 
        $funcJob,
        $funcAction,
        $funcException,
        $funcDetails,
        $funcStackTrace,
        $funcLogPath
    )
    Write-Host
    Write-Host ("{0} - {1}: {2}" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"),$funcAction,$_.Exception) -ForegroundColor Red
    $err = $funcJob
    $err | Add-Member -MemberType NoteProperty -Name "errorAction" -Value $funcAction
    $err | Add-Member -MemberType NoteProperty -Name "errorException" -Value ((($funcException -replace $specialCharacters,'') -replace "`n",' ') -replace "`r",' ')
    $err | Add-Member -MemberType NoteProperty -Name "errorDetails" -Value ((($funcDetails -replace $specialCharacters,'') -replace "`n",' ') -replace "`r",' ')
    $err | Add-Member -MemberType NoteProperty -Name "errorStackTrace" -Value ((($funcStackTrace -replace $specialCharacters,'') -replace "`n",' ') -replace "`r",' ')
    $err | Export-Csv -Path $funcLogPath -NoTypeInformation -Append
    Clear-Variable err
}
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

    #region 6.  Site Creation
    # Determine Destination URL
    if ($job.destSiteURL -like "https://*") {$destURL = $job.destSiteURL.Trim()} else {$destURL = ("{0}/{1}" -f $job.classicSubSiteParentURL, $job.destSiteURL)}

    # Determin if destination site exists
    Write-Host ("      Checking if Site Exists: {0}" -f $job.srcSiteTitle) -NoNewline
    $destSiteExists = Test-CatPnPSiteExistsAppReg -fnUrl $destURL -fnAppID $destAppID -Tenant $destTenantDomain -fnCertPath $destCertPath -fnCertPass $destCertPass

    # Create Site if it doesn't exist
    If ($destSiteExists){ 
        Write-Host (" - Site Already Exists") -Foregroundcolor DarkGray
    }
    Else {
        Write-Host (" - Site Doesn't Exists") -Foregroundcolor DarkGray
        
        # Creating destination site collection
        Write-Host ("      Creating Destination Site: {0}" -f $destURL)
        
        # Get Site Template
        $siteTemplate = $job.siteTemplate.Trim()
        Write-Host ("         Site Template: {0}" -f $siteTemplate) -Foregroundcolor DarkGray
        
        # Creating Site based on Template
        #get public/private value from CSV
        if ($job.modernTeamSiteIsPublic.ToLower() -eq "true") {
            $modernTeamSitePublic = $true
            $msTeamsVisibility = "Public"
        } else {
            $modernTeamSitePublic = $false
            $msTeamsVisibility = "Private"
        }

        try {
            switch ($siteTemplate.ToLower()) {
                # Modern SharePoint Team Site, Modern SharePoint Communciation Site or New Classic Site Collection
                {((($_.ToLower() -eq "communicationsite") -or ($_.ToLower() -eq "teamsite")) -or ((($_.ToLower() -ne "communicationsite") -and ($_.ToLower() -ne "teamsite")) -and ($job.classicSubSiteParentURL -eq "")))} {
                    # Connecting to Root Site
                    Write-Host ("         Connecting to Root Site: {0}" -f $destRootSiteUrl) -Foregroundcolor DarkGray -NoNewline
                    $rootSiteConnection = Connect-PnPOnline -ClientId $destAppID -Url $destRootSiteUrl -Tenant $destTenantDomain -Thumbprint $destCert.Thumbprint -ErrorAction Stop
                    Write-Host " - Connected" -Foregroundcolor DarkGray
                }
                communicationsite {
                    # Creating Modern Communications Site
                    Write-Host ("         Creating Communication Site: {0}" -f $job.destSiteTitle) -Foregroundcolor DarkGray
                    $result = New-PnPSite -Type $siteTemplate -Title $job.destSiteTitle -Url $job.destSiteURL -Owner $job.siteOwner -Wait -SiteDesign $job.CommSiteDesign -ErrorAction Stop
                    Disconnect-PnPOnline -WarningAction SilentlyContinue
                }
                teamsite {
                    # creating Modern Team Site
                    Write-Host ("         Creating Modern Team Site: {0}" -f $job.destSiteTitle) -Foregroundcolor DarkGray
                    $result = New-PnPSite -Type $siteTemplate -Title $job.destSiteTitle -Alias $job.destGroupAlias -Owners $job.siteOwner -isPublic:$modernTeamSitePublic -Wait -ErrorAction Stop
                    Disconnect-PnPOnline -WarningAction SilentlyContinue
                }
                msteams {
                    # Creating Microsoft Teams Workspace
                    Write-Host ("         Creating M365 Teams Group: {0}" -f $job.destSiteTitle) -Foregroundcolor DarkGray
                    
                    #$rootSiteCollection = Connect-PnPOnline -ClientId $destAppID -Url $destRootSiteUrl -Tenant $destTenantDomain -Thumbprint $destCert.Thumbprint -ErrorAction Stop
                    #Write-Host " - Connected" -Foregroundcolor DarkGray

                    # Creating Teams Workspace
                    $result = New-PnPTeamsTeam -DisplayName $job.destSiteTitle -Owners $job.siteOwner -MailNickName $job.destSiteTitle -Visibility $msTeamsVisibility -ErrorAction SilentlyContinue

                    <#
                        try {
                            # Creating M365 Group...
                            $group = New-PnPMicrosoft365Group -DisplayName $job.destSiteTitle -MailnickName $job.destSiteTitle -Owners $job.siteOwner -Description $job.destSiteTitle -Force -ErrorAction SilentlyContinue
                        }
                        catch {
                            Log-MigrationError -funcJob $job -funcAction "Creating M365 Group: " -funcException $_.Exception -funcDetails $_.ErrorDetails -funcStackTrace $_.ScriptStackTrace -funcLogPath $errorPath
                        }

                        try {
                            $getGroup = Get-PnPMicrosoft365Group -Identity $job.destSiteTitle
                            # Creating Teams Workspace
                            $result = New-PnPTeamsTeam -GroupId $getGroup.GroupId
                        } 
                        catch {
                            Log-MigrationError -funcJob $job -funcAction "Creating Teams Workspace: " -funcException $_.Exception -funcDetails $_.ErrorDetails -funcStackTrace $_.ScriptStackTrace -funcLogPath $errorPath
                        }
                    #>
                    
                    #$result = New-PnPTeamsTeam -DisplayName $job.destSiteTitle -Owners $job.siteOwner -MailNickName $job.destSiteTitle.Trim()
                    #$result = New-PnPMicrosoft365Group -DisplayName $job.destSiteTitle -CreateTeam -MailnickName $job.destSiteTitle.Trim() -Owners $job.siteOwner -Description $job.destSiteTitle -Force
                    
                    Disconnect-PnPOnline -Connection $rootSiteConnection -WarningAction SilentlyContinue
                }
                # New Classic Site Collection
                {((($_.ToLower() -ne "communicationsite") -and ($_.ToLower() -ne "teamsite") -and ($_.ToLower() -ne "msteams")) -and ($job.classicSubSiteParentURL -eq ""))} {
                    Write-Host ("         Creating Classic Site: {0}" -f $job.destSiteTitle) -Foregroundcolor DarkGray
                    $result = New-PnPTenantSite -Title $job.destSiteTitle -Url $job.destSiteURL -Owner $job.siteOwner -Template $siteTemplate -TimeZone $timeZone -ErrorAction Stop
                    Disconnect-PnPOnline -WarningAction SilentlyContinue
                }
                # New Classic Web / Sub Site
                {((($_.ToLower() -ne "communicationsite") -and ($_.ToLower() -ne "teamsite") -and ($_.ToLower() -ne "msteams")) -and ($job.classicSubSiteParentURL -ne ""))} {
                    Write-Host ("         Connecting to Parent Site: {0}" -f $job.classicSubSiteParentURL) -Foregroundcolor DarkGray -NoNewline
                    $subsiteConnection = Connect-PnPOnline -ClientId $destAppID -Url $job.classicSubSiteParentURL.Trim() -Tenant $destTenantDomain -Thumbprint $destCert.Thumbprint -ErrorAction Stop
                    Write-Host " - Connected" -Foregroundcolor DarkGray
                    Write-Host ("         Creating SharePoint Classic Sub Site: {0}" -f $job.destSiteTitle) -Foregroundcolor DarkGray
                    $result = New-PnPWeb -Title $job.destSiteTitle -Url $job.destSiteURL.Trim() -Locale $siteLocale -Template $siteTemplate -ErrorAction Stop
                    Disconnect-PnPOnline -WarningAction SilentlyContinue
                }
                Default {
                    Write-Host ("         There are issues with site creation related to the site template: {0} - Migration Canceled" -f $job.destSiteTitle) -Foregroundcolor Red
                    Continue JobLoop
                }
            }
        }
        catch {
            # Output Error Information
            Log-MigrationError -funcJob $job -funcAction "Creating Site" -funcException $_.Exception -funcDetails $_.ErrorDetails -funcStackTrace $_.ScriptStackTrace -funcLogPath $errorPath
            Continue JobLoop
        } 

        #wait for destination site to be available
        $counter = 0
        Write-Host "         Waiting for Site Creation" -NoNewline -ForegroundColor DarkGray
        while (($destSiteExists -eq $false) -and $counter -lt $siteWaitMaxCount) {
            $counter++
            $destSiteExists = Test-CatPnPSiteExistsAppReg -fnUrl $destURL -fnAppID $destAppID -fnTenant $destTenantDomain -fnCertPath $destCertPath -fnCertPass $destCertPass
            if (-not $destSiteExists) {
                Write-Host "." -NoNewline -ForegroundColor DarkGray
                start-sleep -Seconds $siteWaitPause
            }
        }
        Write-Host "." -ForegroundColor DarkGray
        Write-Host ("         Created Destination Site: {0}" -f $job.srcSiteTitle) -ForegroundColor Green
    }
    #endregion #>

    # Run Migrations if destination site exists
    If ($destSiteExists){
        # Connecting to Destination Site
        $destSiteConnection = Connect-PnPOnline -ClientId $destAppID -Url $destURL -Tenant $destTenantDomain -Thumbprint $destCert.Thumbprint
        $destSite = Get-PnPSite -Includes RootWeb, HubSiteId #>

        #region 6.1 Apply Site Template
        If($useCustomSiteTemplates) {
            try {
                switch ($siteTemplate.ToLower()) {               
                    communicationsite {
                        # applying Modern Communications Site template
                        Write-Host ("         Applying Communication Site Template: {0}" -f $job.destinationURL) -Foregroundcolor DarkGray
                        #$applyTemplateResults = Invoke-PnPSiteTemplate -Path "..\_Input Files\$commSiteTemplateFile"
                        $applyTemplateResults = Invoke-PnPSiteTemplate -Path "..\_Input Files\$commSiteTemplateFile" -Parameters @{"SiteUrl"="/sites/$job.destSiteTitle.Trim()"}
                    }
                    teamsite {
                        # creating Modern Team Site template
                        Write-Host ("         Applying Team Site Template: {0}" -f $job.destinationURL) -Foregroundcolor DarkGray
                        #$applyTemplateResults = Invoke-PnPSiteTemplate -Path "..\_Input Files\teamSiteTemplateFile"
                        $applyTemplateResults = Invoke-PnPSiteTemplate -Path "..\_Input Files\$teamSiteTemplateFile" -Parameters @{"SiteUrl"="/sites/$job.destSiteTitle.Trim()"}

                    }
                    msteams {
                        # creating Modern Teams Workspace
                        Write-Host ("         Applying Modern Team Template: {0}" -f $job.destinationURL) -Foregroundcolor DarkGray
                        $applyTemplateResults = Invoke-PnPSiteTemplate -Path "..\_Input Files\$msteamTemplateFile"
                        #$applyTemplateResults = Invoke-PnPSiteTemplate -Path "..\_Input Files\$msteamTemplateFile" -Parameters @{"SiteUrl"="/sites/$job.destSiteTitle.Trim()"}

                    }
                    Default {
                        Write-Host ("         There are issues with applying template: {0} - Migration Canceled" -f $job.destinationURL) -Foregroundcolor Red
                    }
                }
            }
            catch {
                # Output Error Information
                Write-Host ("         {0} - Error Applying Template: {1}" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"),$_.Exception) -ForegroundColor Red
                $err = $job
                $err | Add-Member -MemberType NoteProperty -Name "errorAction" -Value "Site Template"
                $err | Add-Member -MemberType NoteProperty -Name "errorException" -Value $_.Exception
                $err | Add-Member -MemberType NoteProperty -Name "errorDetails" -Value $_.ErrorDetails
                $err | Add-Member -MemberType NoteProperty -Name "errorStackTrace" -Value $_.ScriptStackTrace
                $err | Export-Csv -Path $errorPath -NoTypeInformation -Append
            } 
        }
        #endregion #>

        #region 7.  Configuring Hubs
        # Register site as hub site if not already
        if ($job.isHubSite.ToLower() -eq "true") {
            Write-Host ("      Configuring Hub Site: {0}" -f $job.destSiteTitle) -NoNewline
            $hubSite = Get-PnPHubSite -Identity $job.destSiteURL
            if ($hubSite.SiteId) {
                Write-Host (" - Already Configured") -Foregroundcolor DarkGray
            }
            else {
                try {
                    Register-PnPHubSite -Site $destSite.Url -ErrorAction Stop
                    Write-Host (" - Configured") -Foregroundcolor DarkGray    
                }
                catch {
                    # Output Error Information
                    Log-MigrationError -funcJob $job -funcAction "Configuring Hub Site" -funcException $_.Exception -funcDetails $_.ErrorDetails -funcStackTrace $_.ScriptStackTrace -funcLogPath $errorPath
                }                         
            }
        }
        #endregion #>

        #region 8.  Associating to Hub
        # Determine if already associated to HUB
        if ($job.hubParentURL) {
            Write-Host ("      Configuring Hub Association: {0}" -f $job.destSiteTitle) -NoNewline
            $hubSite = Get-PnPHubSite -Identity $destSite.HubSiteId.GUID -ErrorAction SilentlyContinue
            if ($hubSite.SiteUrl -eq $job.hubParentURL) {
                Write-Host (" - Already Configured") -Foregroundcolor DarkGray
            }
            else {          
                try {
                    Add-PnPHubSiteAssociation -Site $job.destSiteURL -HubSite $job.hubParentURL -ErrorAction Stop
                    Write-Host (" - Configured") -Foregroundcolor DarkGray
                }
                catch {
                    # Output Error Information
                    Log-MigrationError -funcJob $job -funcAction "Associating Hub Site" -funcException $_.Exception -funcDetails $_.ErrorDetails -funcStackTrace $_.ScriptStackTrace -funcLogPath $errorPath
                }
            }
        }
        #endregion #>

        #region 9.  Adding ShareGate Permissions
        # Adding ShareGate service account to destination site
        Write-Host ("      Adding Service Account as Destination Site Owner: {0}" -f $destSGCreds.UserName)
        try {
            Add-PnPSiteCollectionAdmin -Owners $destSGCreds.UserName  -ErrorAction Stop
        }
        catch {
            # Output Error Information
            Log-MigrationError -funcJob $job -funcAction "Adding Service Account" -funcException $_.Exception -funcDetails $_.ErrorDetails -funcStackTrace $_.ScriptStackTrace -funcLogPath $errorPath
            Continue JobLoop
        }
        #endregion #>

        #region 10.  Create Destination Folder (if $createChannels is false) 
        if (-not $createChannels -and $job.destFolder) {
            Write-Host ("      Checking if folder exists in destination: {0}" -f $destFolder) -NoNewline
            $destFolder = ("{0}/{1}" -f $job.destListPath, $job.destFolder)
            try {
                $tempFolder = Get-PnPFolder -Url $destFolder -ErrorAction Stop
                Write-Host " - Folder Already Exists"
            }
            catch {
                Write-Host " - Folder Does not Exists"
            }
            if (-not $tempfolder) {
                Write-Host ("      Creating Folder in Destination: {0}" -f $destFolder)
                $folders = $job.destFolder -split "/"
                $parentFolder = $job.destListPath
                foreach ($folder in $folders) {
                    try {
                        try {
                            $folderExists = Get-PnPFolder -Url ($parentFolder + "/" + $folder) -ErrorAction Stop
                        }
                        catch {
                            $folderExists = $null
                        }
                        if (-not $folderExists) {
                            $tempFolder = Add-PnPFolder -Name $folder -Folder $parentFolder -ErrorAction Stop
                        }
                        $parentFolder = ($parentFolder + "/" + $folder)
                    }
                    catch {
                        # Output Error Information
                        Log-MigrationError -funcJob $job -funcAction "Creating Folder" -funcException $_.Exception -funcDetails $_.ErrorDetails -funcStackTrace $_.ScriptStackTrace -funcLogPath $errorPath
                        Continue JobLoop
                    }
                }
            }
            if (-not $tempfolder) {
                Clear-Variable folders
                Clear-Variable parentFolder
                Clear-Variable folderExists
            }
            Clear-Variable tempFolder
        }
        #endregion #>

        #region 10.1  Create Team Channel for Libraries (if $createChannels is true) 
        try {
            if($createChannels) {

                # Check if current item was marked for a channel in csv
                if ($job.channelType -eq "Standard" -or $job.channelType -eq "Shared" -or $job.channelType -eq "Private") {

                    # Destination Channel Name
                    $destChannel = ($job.srcListName -replace $specialCharacters,'')

                    Write-Host ("      Checking if team channel exists in destination: {0}" -f $destChannel) -NoNewline
                    try {
                        $channelExists = Get-PnPTeamsChannel -Team $job.destSiteTitle | Where-Object { $_.DisplayName -eq $destChannel }
                        Write-Host " - Channel Already Exists"
                    }
                    catch {
                        Write-Host " - Channel Does not Exists"
                    }
                    if (-not $channelExists) {
                        Write-Host ("      Creating Channel in Destination: {0}" -f $destChannel)                
                        try {
                            $newChannel = Add-PnPTeamsChannel -Team $job.destSiteTitle -DisplayName $destChannel -ChannelType $job.channelType -OwnerUPN $job.siteOwner -IsFavoriteByDefault $true
                        
                            try{
                                Write-Host " - Get the newly created channel folder metadata"
                                $getNewChannelFolders = Get-PnPTeamsChannelFilesFolder -Team $job.destSiteTitle -Channel $destChannel
                            
                                if(-not $getNewChannelFolders) {
                                    Write-Host " - Channel folder metadata retrieved"
                                }
                            }
                            catch{
                                Write-Host " - Channel folder metadata could NOT be retrieved"
                            }
                        
                        }
                        catch {
                            Write-Host " - Channel was not created"
                            
                            # Output Error Information
                            Log-MigrationError -funcJob $job -funcAction "Creating Channel" -funcException $_.Exception -funcDetails $_.ErrorDetails -funcStackTrace $_.ScriptStackTrace -funcLogPath $errorPath
                        }
                    }
        
                    if (-not $newChannel) {
                        Clear-Variable channelExists
                    }
                    Clear-Variable newChannel
                    Clear-Variable getNewChannelFolders
                }
            }
        }
        catch {
            # Output Error Information
            Log-MigrationError -funcJob $job -funcAction "Creating Team Channel" -funcException $_.Exception -funcDetails $_.ErrorDetails -funcStackTrace $_.ScriptStackTrace -funcLogPath $errorPath
            Continue JobLoop
        } 
        #endregion #>

        # Disconnecting from Destination Site
        Disconnect-PnPOnline #>
                
        #region 11. Configuring ShareGate job settings
        #region Loading sites
        Write-Host ("      Loading Source Site: {0}" -f $job.srcSiteTitle)
        try {
            $sgSourceSite = Connect-Site -Url $job.srcSiteURL -Credential $srcSGCreds -ErrorAction Stop
        }
        catch {
            # Output Error Information
            Log-MigrationError -funcJob $job -funcAction "Loading Source Site" -funcException $_.Exception -funcDetails $_.ErrorDetails -funcStackTrace $_.ScriptStackTrace -funcLogPath $errorPath
            Continue JobLoop
        }
        Write-Host ("      Loading Destination Site: {0}" -f $job.destSiteTitle)
        try {
            if ($basicAuthEnabled) {

                # Determine site to load based on whether it's a channel site or regular site
                if($createChannels) {
                    # Determine Channel Site URL
                    $destChannelSiteURL= $job.destSiteURL.Trim() + "-" + $job.destListName.Trim()

                    # Load destination channel site
                    $sgDestSite = Connect-Site -Url $destChannelSiteURL -Credential $destSGCreds -ErrorAction Stop

                } else {

                    # Load destination regular site
                    $sgDestSite = Connect-Site -Url $destURL -Credential $destSGCreds -ErrorAction Stop
                }
            } else {
                if($createChannels) {
                    # Determine Channel Site URL
                    $destChannelSiteURL= $job.destSiteURL.Trim() + "-" + $job.destListName.Trim()

                    # Load destination channel site
                    Write-Host ("         Channel Site URL: " -f $destChannelSiteURL) -ForegroundColor DarkGray
                    $sgDestSite = Connect-Site -Url $destChannelSiteURL -UseCredentialsFrom $destSGConnection -ErrorAction Stop

                } else {
                    Write-Host ("         Site URL: " -f $destChannelSiteURL) -ForegroundColor DarkGray
                    $sgDestSite = Connect-Site -Url $destURL -UseCredentialsFrom $destSGConnection -ErrorAction Stop
                }
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
        $sgTaskName = ("SP List - {0} - {1}" -f $job.srcSiteTitle, $job.srcListName)
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
        If ($job.destFolder) {
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
                $sgDestList = Get-List -Site $sgDestSite -Name $job.destListName
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

       #region 11.1 Force Check In Files        
       if ($runForceCheckIn -and $job.forceCheckIn.ToLower() -eq "true") { 
            
            # Check if files to check in are on-prem or online
            if($forceCheckInOnPrem) {

                #SP On Prem source
                try {
                    # Get site context
                    $site = Get-SPSite $job.srcSiteURL
                    $web = $site.RootWeb
                    $list = $web.Lists[$job.srcListName]

                    # Get checked out files
                    if ($list.BaseTemplate -eq "DocumentLibrary") {

                        # Get all checked out files
                        $checkedOutFiles = $list.CheckedOutFiles
                    
                        foreach ($file in $checkedOutFiles) {
                            # Check if file is greater than or equal to $pastDate
                            if ($file.TimeLastModified -ge [datetime]$pastDate) {
                                # Check in the file
                                $file.CheckIn("Checked in by System Account", [Microsoft.SharePoint.SPCheckinType]::MajorCheckIn)
                            }
                        }
                    }
                    #endregion
                } catch {
                    # Output Error Information
                    Log-MigrationError -funcJob $job -funcAction "Running Force Check In Files Job" -funcException $_.Exception -funcDetails $_.ErrorDetails -funcStackTrace $_.ScriptStackTrace -funcLogPath $errorPath              
                }
            } else {    
                # SPO source
                if ($runForceCheckIn -and $job.forceCheckIn.ToLower() -eq "true") {            
                    Write-Host
                    Write-Host "Checking In Source Library Files ..." -ForegroundColor Magenta
                    try {

                        # Get files from the current library omitting folders
                        If($pastDate -ne $Null -and $pastDate -ne '') {
                            # Get all checked out files that are greater than or equal to $pastDate
                            $files = Get-PnPListItem -List $job.srcListName -PageSize $pageSize | Where-Object { $_.FileSystemObjectType -eq 'File' -and $_.CheckoutUser -ne $Null -and $_.Modified -ge $pastDate }
                        } else {
                            # Get all checked out files 
                            $files = Get-PnPListItem -List $job.srcListName -PageSize $pageSize | Where-Object { $_.FileSystemObjectType -eq 'File' -and $_.CheckoutUser -ne $Null }
                        }

                        # region Loop through each file in library
                        :Files foreach ($item in $files) {
                            $fileUrl = $item.FieldValues.FileRef # Get the relative URL of the file

                            try {
                                # Check in the file
                                Set-PnPFileCheckedIn -Url $fileUrl -CheckinType Major -Comment "Checked in by System Account"                    
                                Write-Host "Checked in file: $fileUrl"
                            } catch {
                                Write-CatErrorCSV -funcCSVItem $item -funcAction "Error checking in file: $fileUrl" -funcException $_.Exception -funcDetails $_.ErrorDetails -funcStackTrace $_.ScriptStackTrace -funcFilePath $errorPath
                                Continue Files
                            }
                        }
                        #endregion
                    } catch {
                        # Output Error Information
                        Log-MigrationError -funcJob $job -funcAction "Running Force Check In Files Job" -funcException $_.Exception -funcDetails $_.ErrorDetails -funcStackTrace $_.ScriptStackTrace -funcLogPath $errorPath
                    }
                }
            }
        }
        #end region #>

        #region 12. Running ShareGate Migration
        If ($job.destFolder) {
            #region Running Content Job
            Write-Host ("      Running Content Job: {0}" -f $sgTaskName)
            try {
                $sgResult = Copy-Content `
                    -SourceList $sgSourceList `
                    -DestinationList $sgDestList `
                    -DestinationFolder $job.destFolder `
                    -WaitForImportCompletion:$sgWaitForCompletion `
                    -Template $sgPropertyTemplate `
                    -CopySettings $sgCopySettings `
                    -MappingSettings $sgCopyMappings `
                    -TaskName $sgTaskName `
                    -ErrorAction Stop
            }
            catch {
                # Output Error Information
                Log-MigrationError -funcJob $job -funcAction "Running Content Job" -funcException $_.Exception -funcDetails $_.ErrorDetails -funcStackTrace $_.ScriptStackTrace -funcLogPath $errorPath
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
                Log-MigrationError -funcJob $job -funcAction "Running List Job" -funcException $_.Exception -funcDetails $_.ErrorDetails -funcStackTrace $_.ScriptStackTrace -funcLogPath $errorPath
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

        #region 13. Removing ShareGate Permissions        
        # Removing ShareGate service account to destination site
        Write-Host ("      Removing Service Account as Destination Site Owner: {0}" -f $destSGCreds.UserName)
        try {
            $destSiteConnection = Connect-PnPOnline -ClientId $destAppID -Url $destURL -Tenant $destTenantDomain -Thumbprint $destCert.Thumbprint -ErrorAction Stop
            Get-PnPSiteCollectionAdmin | Where-Object -Property Email -eq $destSGCreds.UserName | Remove-PnPSiteCollectionAdmin
        }
        catch {
            # Output Error Information
            Log-MigrationError -funcJob $job -funcAction "Addding Service Account" -funcException $_.Exception -funcDetails $_.ErrorDetails -funcStackTrace $_.ScriptStackTrace -funcLogPath $errorPath
        }
        Disconnect-PnPOnline -WarningAction SilentlyContinue
        #endregion #>
        
        # Log complete
        Write-Host ("      Completed Successfully") -foregroundcolor "green"       
    }
    Else {
        # Output Error Information
        Log-MigrationError -funcJob $job -funcAction "Connecting to Destination Site" -funcException $null -funcDetails $null -funcStackTrace $null -funcLogPath $errorPath
        Continue JobLoop
    }
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