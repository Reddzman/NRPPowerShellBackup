<#  Cat.SG--MigrationPreFlight.ps1

This script automates SharePoint Online sites and workspaces configuration in preparation for ShareGate Migration.  It is currently leveraging an App 
Registration in the destination Tenant to perform configuration. 

This pre flight script has been tested with:
    -   ShareGate\Cat.SG-SPListMigrations.ps1
    -   ShareGate\Cat.SG-NetworkShareMigrations.ps1

** This should be executed in PowerShell 7+ with the latest PnP.PowerShell

Dependencies
    -   Destination Azure App Registrations (Catapult Powershell & ShareGate)
    -   PowerShell 7 (PowerShell Core)
    -   PnP.PowerShell Module
    -   MicrosoftTeams Module
    -	..\General\Cat.GeneralFunctions.psm1
    -   ..\Reporting\Cat.ReportingFunctions.psm1
    -	Input CSV File in Input Files folder

Call script with parameters
    import-module MicrosoftTeams # run do to strange module bug
    cd C:\GitHub\CatPowerShell\ShareGate\
	ShareGate\Cat.SG-SPOnPremListMigrations.ps1: .\Cat.SG--MigrationPreFlight.ps1 -jobFile "Samples\ShareGate\Cat.SG-SPOnPremListMigrations-Client.csv" -destTenant "nrpdev2" -destAppID "caf6b11a-cf03-416c-b0f0-67a09a8e30ee" -destCertPath "..\..\NRPDev\AppRegistration\PnPCert-260804.pfx" -destTeamsManagedPath "teams" -timeZone 12 -siteLocale 1033
    ShareGate\Cat.SG-NetworkShareMigrations.ps1: .\Cat.SG--MigrationPreFlight.ps1 -jobFile "Samples\ShareGate\Cat.SG-NetworkShareMigrations-Test.csv" -destTenant "nrpdev2" -destAppID "caf6b11a-cf03-416c-b0f0-67a09a8e30ee" -destCertPath "..\..\NRPDev\AppRegistration\PnPCert-260804.pfx" -destTeamsManagedPath "teams" -timeZone 12 -siteLocale 1033

Table of Contents
	1.	Snapin(s), Modules & Environment Connection(s)
	2.	Input & Output Files
	3.	Variables
    4.  Main Script
    5.  Site Creation
    6.  Configuring Hubs
    7.  Associating to Hub
    8.  Creating Document Library
    9.  Creating Teams Channels
    10. Applying SharePoint Site Template
    11. Adding ShareGate Permissions
    12. Adding ShareGate Permissions
    13. Create Destination Folder
#>

#region Script Parameters
Param(
    [String] $jobFile = "Q.SG-SPOnPremListMigrations-Client.csv",
    [string] $destTenant = "{Destination Tenant Name}",
    [String] $destAppID = "{Azure App Registration ID}",
    [String] $destCertPath = "C:\Quisitive\PowerShell\App Reg\ClientCertFileName.pfx",
    [String] $destTeamsManagedPath = "sites", # The managed path used by SharePoint site behind Microsoft Teams. IE "sites" or "teams"
    [String] $timeZone = 12, # 13 = Pacific Time (US & Canada); 38 = Arizona; 12 = Mountain Time (US & Canada); 11 = Central Time (US & Canada); 10 = Eastern Time (US & Canada)
    [String] $siteLocale = 1033 # 1033 = English - United States
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
Start-Transcript -Path ("..\_Log Files\{0}.SG-PreFlight-{1}.log" -f $destTenant, $scriptStart[1])

# PowerShell Settings
$ProgressPreference = "SilentlyContinue"
$env:PNPPOWERSHELL_UPDATECHECK="false"
#endregion

#region 1.	Snapin(s), Modules & Environment Connection(s)
# Load modules
Import-Module PnP.PowerShell
Import-Module MicrosoftTeams

# Load Custom Modules
Import-Module "..\Reporting\Cat.ReportingFunctions.psm1"
#endregion

#region 2.	Input & Output Files
# File Paths
$inPath = ("..\_Input Files\{0}" -f $jobFile)
$errorPath = ("..\_Output Files\{0}.SG-PreFlightErrors-{1}.log" -f $destTenant, $scriptStart[1])

# Import Input File(s)
# Load CSV file
If (Test-Path ($inPath)){
	$inFile = Import-Csv $inPath -Encoding Unicode
}
else {
	Write-Host
	throw ("The jobs file does not exist: {0}" -f $inPath)	
}
#endregion

#region 3.	Variables
# Script Variables
$destAdminUrl = ("https://{0}-admin.sharepoint.com/" -f $destTenant)
$destRootSiteUrl = ("https://{0}.sharepoint.com/" -f $destTenant)
$destTenantDomain = ("{0}.onmicrosoft.com" -f $destTenant)
$siteWaitMaxCount = 10
$siteWaitPause = 20 # Seconds
$channelWaitMaxCount = 10
$channelWaitPause = 20 # Seconds
$spTempWaitMaxCount = 10
$specialCharacters = "[~``!@#$%^&*()_-]"

# User Inputs
Write-Host
Write-Host "Enter Connection Information" -ForegroundColor Cyan
$destCertPass = ConvertTo-SecureString -String (Read-Host -Prompt "   Enter DESTINATION certificate password for the app registration") -AsPlainText -Force
$destCert = Get-PnPAzureCertificate -CertificatePath $destCertPath -CertificatePassword $destCertPass
$destSGCreds = Get-Credential -Message "Enter ShareGate credentials (username required) for the DESTINATION environment:"   
 
# Variable Output
Write-Host
Write-Host "Script Parameters" -ForegroundColor Cyan
Write-Host ("   Job File: {0}" -f $jobFile)
Write-Host ("   Destination Tenant: {0}" -f $destTenant)
Write-Host ("   Destination App ID: {0}" -f $destAppID)
Write-Host ("   Destination Cert Path: {0}" -f $destCertPath)
Write-Host ("   Time Zone: {0}" -f $timeZone)
Write-Host ("   Site Locale: {0}" -f $siteLocale)
Write-Host
Write-Host "Script Variables" -ForegroundColor Cyan
Write-Host ("   Destination Admin Url: {0}" -f $destAdminUrl)
Write-Host ("   Destination Root Site Url: {0}" -f $destRootSiteUrl)
Write-Host ("   Destination Tenant Url: {0}" -f $destTenantDomain)
Write-Host ("   Site Creation Wait Max Iterations: {0}" -f $siteWaitMaxCount)
Write-Host ("   Site Creation Wait Pause Duration: {0}" -f $siteWaitPause)
Write-Host ("   Channel Creation Wait Max Iterations: {0}" -f $channelWaitMaxCount)
Write-Host ("   Channel Creation Wait Pause Duration: {0}" -f $channelWaitPause)
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

    #region 6.  Site Creation
    # Determine Destination URL
    if ($job.destSiteURL -like "https://*") {$destURL = $job.destSiteURL.Trim()} else {$destURL = ("{0}/{1}" -f $job.classicSubSiteParentURL, $job.destSiteURL)}

    # Determine if destination site exists
    Write-Host ("      Checking if Site Exists: {0}" -f $job.srcTitle) -NoNewline
    $destSiteExists = Test-CatPnPSiteExistsAppReg -fnUrl $destURL -fnAppID $destAppID -fnTenant $destTenantDomain -fnCertPath $destCertPath -fnCertPass $destCertPass
    
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
        
        #region Creating Site based on Template
        #get public/private value from CSV
        if ($job.modernTeamSiteIsPublic.ToLower() -eq "true") {$modernTeamSitePublic = $true; $msTeamsVisibility = "Public"} else {$modernTeamSitePublic = $false; $msTeamsVisibility = "Private"}

        try {
            switch ($siteTemplate) {
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
                teamsitenogroup {
                    # creating Modern Team Site without a Micrsoft 365 Group
                    Write-Host ("         Creating Modern Team Site w/o M365 Group: {0}" -f $job.destSiteTitle) -Foregroundcolor DarkGray
                    $result = New-PnPSite -Type TeamSiteWithoutMicrosoft365Group -Title $job.destSiteTitle -Url $job.destSiteURL -Owner $job.siteOwner  -Wait -ErrorAction Stop
                    Disconnect-PnPOnline -WarningAction SilentlyContinue
                }
                msteams {
                    # Connecting to Teams
                    $rootSiteConnection = Connect-MicrosoftTeams -TenantId $destTenantDomain -CertificateThumbprint $destCert.Thumbprint -ApplicationId $destAppID
                    
                    # creating Microsoft Teams Workspace
                    Write-Host ("         Creating Microsoft Teams Workspace: {0}" -f $job.destSiteTitle) -Foregroundcolor DarkGray
                    $result = New-Team -DisplayName $job.destSiteTitle -MailNickName $job.destGroupAlias -Visibility $msTeamsVisibility -Owner $job.siteOwner -AllowCreatePrivateChannels:$true
                    Disconnect-MicrosoftTeams
                }
                # New Classic Site Collection
                {((($_.ToLower() -ne "communicationsite") -and ($_.ToLower() -ne "teamsite") -and ($_.ToLower() -ne "msteams") -and ($_.ToLower() -ne "teamsitenogroup")) -and ($job.classicSubSiteParentURL -eq ""))} {
                    Write-Host ("         Creating Classic Site: {0}" -f $job.destSiteTitle) -Foregroundcolor DarkGray
                    $result = New-PnPTenantSite -Title $job.destSiteTitle -Url $job.destSiteURL -Owner $job.siteOwner -Template $siteTemplate -TimeZone $timeZone -ErrorAction Stop
                    Disconnect-PnPOnline -WarningAction SilentlyContinue
                }
                # New Classic Web / Sub Site
                {((($_.ToLower() -ne "communicationsite") -and ($_.ToLower() -ne "teamsite") -and ($_.ToLower() -ne "msteams") -and ($_.ToLower() -ne "teamsitenogroup")) -and ($job.classicSubSiteParentURL -ne ""))} {
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
            Write-CatErrorCSV -funcJob $job -funcAction "Creating Site" -funcException $_.Exception -funcDetails $_.ErrorDetails -funcStackTrace $_.ScriptStackTrace -funcLogPath $errorPath
            Continue JobLoop
        }
        #endregion

        #region waiting for SharePoint site to be available
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
        #endregion

        #region waiting for Teams workspace to be available
        if ($siteTemplate.ToLower() -eq "msteams") {
            Write-Host "." -ForegroundColor DarkGray
            Write-Host "         Waiting for Teams Workspace Creation" -NoNewline -ForegroundColor DarkGray
            $destSiteExists = $false
            $counter = 0
            Write-Host "." -NoNewline -ForegroundColor DarkGray
            start-sleep -Seconds $siteWaitPause
            $teamsConnection = Connect-MicrosoftTeams -TenantId $destTenantDomain -CertificateThumbprint $destCert.Thumbprint -ApplicationId $destAppID
            while (($destSiteExists -eq $false) -and $counter -lt $siteWaitMaxCount) {
                $counter++
                $tempTeam = Get-Team -MailNickName $job.destGroupAlias
                If ($tempTeam) {$destSiteExists = $true} Else {$destSiteExists = $false}
                if (-not $destSiteExists) {
                    Write-Host "." -NoNewline -ForegroundColor DarkGray
                    start-sleep -Seconds $siteWaitPause
                }
            }
            Disconnect-MicrosoftTeams
        }
        #endregion

        Write-Host "." -ForegroundColor DarkGray
        Write-Host ("         Created Destination: {0}" -f $job.srcTitle) -ForegroundColor Green
    }
    #endregion #>

    # Run Migrations if destination site exists
    If ($destSiteExists){
        # Connecting to Destination Site
        $destConnection = Connect-PnPOnline -ClientId $destAppID -Url $destURL -Tenant $destTenantDomain -Thumbprint $destCert.Thumbprint
        $destSite = Get-PnPSite -Includes RootWeb, HubSiteId #>

        #region 6.  Configuring Hubs
        # Register site as hub site if not already
        if ($job.isHubSite.ToLower() -eq "true") {
            Write-Host ("      Configuring Hub Site: {0}" -f $job.destSiteTitle) -NoNewline
            $hubSite = Get-PnPHubSite -Identity $job.destSiteURL
            if ($hubSite.SiteId) {
                Write-Host (" - Already Configured") -Foregroundcolor DarkGray
            }
            else {
                try {
                    $result = Register-PnPHubSite -Site $destSite.Url -ErrorAction Stop
                    Write-Host (" - Configured") -Foregroundcolor DarkGray    
                }
                catch {
                    # Output Error Information
                    Write-CatErrorCSV -funcJob $job -funcAction "Configuring Hub Site" -funcException $_.Exception -funcDetails $_.ErrorDetails -funcStackTrace $_.ScriptStackTrace -funcLogPath $errorPath
                }                         
            }
        }
        #endregion #>

        #region 7.  Associating to Hub
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
                    Write-CatErrorCSV -funcJob $job -funcAction "Associating Hub Site" -funcException $_.Exception -funcDetails $_.ErrorDetails -funcStackTrace $_.ScriptStackTrace -funcLogPath $errorPath
                }
            }
        }
        #endregion #>

        #region 8.  Creating Document Library
        Write-Host ("      Creating Document Library: {0}" -f $job.destListName) -NoNewline
        $library = Get-PnPList -Identity $job.destListPath -ErrorAction SilentlyContinue
        if ($library) {
            Write-Host (" - Already Exists") -Foregroundcolor DarkGray
        }
        else {          
            try {
                $result = New-PnPList -Title $job.destListName -Template DocumentLibrary -Url $job.destListPath -OnQuickLaunch
                Write-Host (" - Created") -Foregroundcolor DarkGray
            }
            catch {
                # Output Error Information
                Write-CatErrorCSV -funcCSVItem $job -funcAction "Creating Library" -funcException $_.Exception -funcDetails $_.ErrorDetails -funcStackTrace $_.ScriptStackTrace -funcFilePath $errorPath
            }
        }
        Clear-Variable library
        #endregion

        #region 9.  Creating Teams Channels
        If (($job.destChannel.ToLower() -ne "none") -and ($job.destChannel -ne "")){
            try {
                Write-Host ("      Checking if team channel exists in destination: {0}" -f $job.destChannel) -NoNewline
                $destConnection = Connect-MicrosoftTeams -TenantId $destTenantDomain -CertificateThumbprint $destCert.Thumbprint -ApplicationId $destAppID
                $destGroupId = (Get-Team -MailNickName $job.destGroupAlias | where {$_.MailNickName -eq $job.destGroupAlias}).GroupID
                $destChannel = Get-TeamChannel -GroupId $destGroupId | Where {$_.DisplayName -eq $job.destChannel}

                if ($destChannel.DisplayName -ne $job.destChannel) {
                    Write-Host (" - Channel Doesn't exist") 
                    Write-Host ("      Creating Channel in Destination: {0}" -f $job.destChannel)
                    Write-Host ("         Channel Type: {0}" -f $job.destChannelType) -ForegroundColor DarkGray
                    If ($job.destChannelType -eq "Standard") {
                        $result = New-TeamChannel -GroupId $destGroupId -DisplayName $job.destChannel -MembershipType $job.destChannelType
                    } Else {
                        $result = New-TeamChannel -GroupId $destGroupId -DisplayName $job.destChannel -MembershipType $job.destChannelType -Owner $job.siteOwner
                    }
                    $newChannel = Get-TeamChannel -GroupId $destGroupId | Where {$_.DisplayName -eq $job.destChannel}                    
                    Write-Host ("         Created Teams Channel: {0}" -f $job.destChannel) -ForegroundColor Green
                    Disconnect-MicrosoftTeams

                    #region waiting for destination channel to be available
                    $destFolderExists = $false
                    $counter = 0
                    Write-Host "         Waiting for file folder creation" -NoNewline -ForegroundColor DarkGray
                    $result = Get-PnPTeamsChannelFilesFolder -Team $destGroupId -Channel "General" -ErrorAction Stop
                    while (($destFolderExists -eq $false) -and $counter -lt $channelWaitMaxCount) {
                        $counter++
                        try {
                            $result = Get-PnPTeamsChannelFilesFolder -Team $destGroupId -Channel $newChannel.Id -ErrorAction Stop
                            $destFolderExists = $true
                        }
                        catch {
                            Write-Host "." -NoNewline -ForegroundColor DarkGray
                            start-sleep -Seconds $channelWaitPause
                        }
                    }
                    Write-Host "." -ForegroundColor DarkGray                    

                    Clear-Variable newChannel
                    Clear-Variable result
                    #endregion
                } else {
                    if($job.destChannel.ToLower() -eq "general") {
                        $result = Get-PnPTeamsChannelFilesFolder -Team $destGroupId -Channel "General" -ErrorAction Stop
                    }
                    Write-Host (" - Channel Already Exists") -ForegroundColor DarkGray
                }
            }
            catch {
                # Output Error Information
                Write-CatErrorCSV -funcJob $job -funcAction "Creating Teams Channel" -funcException $_.Exception -funcDetails $_.ErrorDetails -funcStackTrace $_.ScriptStackTrace -funcLogPath $errorPath
                Continue JobLoop
            }

            # Variable Cleanup
            Clear-Variable destConnection
            Clear-Variable destChannel
            Clear-Variable destGroupId
        }
        #endregion #>

        #region 10. Applying SharePoint Site Template
        If(($job.customSPTemplate.ToLower() -ne "none") -and ($null -ne $job.customSPTemplate)) {
            Write-Host ("      Applying custom SharePoint site template: {0}" -f $job.customSPTemplate)
            try {
                $result = Add-PnPSiteDesignTask -SiteDesignId $job.customSPTemplate
            }
            catch {
                # Output Error Information
                Write-CatErrorCSV -funcJob $job -funcAction "Applying SharePoint Site Template" -funcException $_.Exception -funcDetails $_.ErrorDetails -funcStackTrace $_.ScriptStackTrace -funcLogPath $errorPath
                Continue JobLoop
            }

            #region waiting for destination channel to be available
            $tempApplied = $false
            $counter = 0
            Write-Host "         Waiting for template task to start" -NoNewline -ForegroundColor DarkGray
            while (($tempApplied -eq $false) -and $counter -lt $spTempWaitMaxCount) {
                $counter++
                $status = Get-PnPSiteDesignTask -Identity $result.Id
                if ($status.Id) {
                    Write-Host "." -NoNewline -ForegroundColor DarkGray
                    start-sleep -Seconds $channelWaitPause
                } else {
                    $tempApplied = $true
                }
            }
            Write-Host "." -ForegroundColor DarkGray
            Write-Host ("         Site Template Applied: {0}" -f $job.srcTitle) -ForegroundColor Green
            #endregion
        }
        #endregion

        #region 11. Adding ShareGate Parent Permissions
        # Adding ShareGate service account to destination site
        Write-Host ("      Adding Service Account as Destination Site Owner: {0}" -f $destSGCreds.UserName)
        try {
            Add-PnPSiteCollectionAdmin -Owners $destSGCreds.UserName  -ErrorAction Stop
        }
        catch {
            # Output Error Information
            Write-CatErrorCSV -funcJob $job -funcAction "Adding Service Account to Parent Site" -funcException $_.Exception -funcDetails $_.ErrorDetails -funcStackTrace $_.ScriptStackTrace -funcLogPath $errorPath
            Continue JobLoop
        }
        #endregion #>

        # Disconnecting from Destination Site
        Disconnect-PnPOnline #>

        #region 12.  Adding ShareGate Channel Permissions
        try {
            switch ($job.destChannelType) {
                private {
                    $channelType = "Private"
                    Write-Host ("      Adding Service Account as {0} Channel Site Admin: {1}" -f $channelType, $destSGCreds.UserName)
                    $channelURL = ("{0}{1}/{2}-{3}" -f $destRootSiteUrl, $destTeamsManagedPath, ($job.destSiteTitle -replace " ",""), ($job.destChannel -replace " ",""))
                    Write-Host ("         Site URL: {0}" -f $channelURL) -ForegroundColor DarkGray
                    $destConnection = Connect-PnPOnline -ClientId $destAppID -Url $channelURL -Tenant $destTenantDomain -Thumbprint $destCert.Thumbprint
                    Add-PnPSiteCollectionAdmin -Owners $destSGCreds.UserName -ErrorAction Stop
                    Disconnect-PnPOnline 
                }
                shared {
                    $channelType = "Shared"
                    Write-Host ("      Adding Service Account as {0} Channel Site Admin: {1}" -f $channelType, $destSGCreds.UserName)
                    $channelURL = ("{0}{1}/{2}-{3}" -f $destRootSiteUrl, $destTeamsManagedPath, ($job.destSiteTitle -replace " ",""), ($job.destChannel -replace " ",""))
                    Write-Host ("         Site URL: {0}" -f $channelURL) -ForegroundColor DarkGray
                    $destConnection = Connect-PnPOnline -ClientId $destAppID -Url $channelURL -Tenant $destTenantDomain -Thumbprint $destCert.Thumbprint
                    Add-PnPSiteCollectionAdmin -Owners $destSGCreds.UserName -ErrorAction Stop
                    Disconnect-PnPOnline 
                }
            }
        }
        catch {
            # Output Error Information
            Write-CatErrorCSV -funcJob $job -funcAction ("Adding Service Account to {0} Channel site" -f $job.destChannelType) -funcException $_.Exception -funcDetails $_.ErrorDetails -funcStackTrace $_.ScriptStackTrace -funcLogPath $errorPath
            Continue JobLoop
        }
        #endregion #>

        #region 13.  Create Destination Folder
        if (($job.destFolder.ToLower() -ne "none") -and ($job.destFolder -ne "")) {
            Write-Host ("      Checking if folder exists in destination: {0}" -f $job.destFolder) -NoNewline
            switch($job.destChannelType.ToLower()) {
                { $_ -eq "private" -or $_ -eq "shared" } {
                    # Private or Shared channel destination folder
                    $destFolder = ("{0}/{1}" -f $job.destListPath, $job.destFolder)
                    $parentFolder = $job.destListPath
                    $channelURL = ("{0}{1}/{2}-{3}" -f $destRootSiteUrl, $destTeamsManagedPath, ($job.destSiteTitle -replace " ",""), ($job.destChannel -replace " ",""))
                    $destConnection = Connect-PnPOnline -ClientId $destAppID -Url $channelURL -Tenant $destTenantDomain -Thumbprint $destCert.Thumbprint 
                }                
                standard {
                    # Standard channel destination folder
                    $destFolder = ("Shared Documents/{0}/{1}" -f $job.destChannel, $job.destFolder) 
                    $parentFolder = "Shared Documents/{0}" -f $job.destChannel
                    $channelURL = $destURL
                    $destConnection = Connect-PnPOnline -ClientId $destAppID -Url $channelURL -Tenant $destTenantDomain -Thumbprint $destCert.Thumbprint
                }
                Default {
                    # All other destinations
                    $destFolder = ("{0}/{1}" -f $job.destListPath, $job.destFolder)
                    $parentFolder = $job.destListPath
                    $destConnection = Connect-PnPOnline -ClientId $destAppID -Url $destURL -Tenant $destTenantDomain -Thumbprint $destCert.Thumbprint
                }
            }

            Write-Host
            Write-Host ("      Destination Folder {0}" -f $destFolder) -ForegroundColor DarkGray
            Write-Host ("      Parent Folder {0}" -f $parentFolder) -ForegroundColor DarkGray
            Write-Host ("      Channel Type {0}" -f $job.destChannelType ) -ForegroundColor DarkGray
            Write-Host ("      Channel {0}" -f $job.destChannel ) -ForegroundColor DarkGray
            Write-Host ("      Channel URL {0}" -f $channelURL) -ForegroundColor DarkGray

            #region - Check if library needs to be created first
            If($job.destChannel.ToLower() -eq "none" -or ($job.destChannel -eq "")) {
            #If( ($job.destChannelType.ToLower -eq "private" -or $job.destChannelType.ToLower -eq "shared") -and ($job.destChannel.ToLower() -eq "none" -or $job.destChannel -eq "")) {
                try {
                    try{
                        $tempLib = Get-PnPList $parentFolder -ErrorAction Stop
                        Write-Host " - Library Already Exists" -ForegroundColor DarkGray -NoNewline
                    }
                    catch {
                        Write-Host " - Library Does not Exists" -ForegroundColor DarkGray -NoNewline
                    }           
                }
                catch {
                    # Output Error Information
                    Write-CatErrorCSV -funcJob $job -funcAction "Creating Library" -funcException $_.Exception -funcDetails $_.ErrorDetails -funcStackTrace $_.ScriptStackTrace -funcLogPath $errorPath
                    Continue JobLoop
                }                
                if(-not $tempLib){
                    Write-Host ("      Creating Library in Destination: {0}" -f $parentFolder)
                    New-PnPList -Title $job.destListName -Url $job.destListPath -Template DocumentLibrary
                    #New-PnPList -Title $parentFolder -Url $job.destListPath -Template DocumentLibrary
                    
                }                
                if ($tempLib) {
                    Clear-Variable tempLib
                } 
            }
            #endregion

            #region - Check for folder and if not found create it
            try {
                $tempFolder = Get-PnPFolder -Url $destFolder -ErrorAction Stop
                Write-Host " - Folder Already Exists" -ForegroundColor DarkGray
            }
            catch {
                Write-Host " - Folder Does not Exists" -ForegroundColor DarkGray
            }
            if (-not $tempfolder) {
                Write-Host ("      Creating Folder in Destination: {0}" -f $destFolder)              

                $folders = $job.destFolder -split "/"                
                foreach ($folder in $folders) {
                    try {
                        try {
                            $folderExists = Get-PnPFolder -Url ($parentFolder + "/" + $folder) -ErrorAction Stop
                        }
                        catch {
                            $folderExists = $null
                        }
                        if (-not $folderExists) {
                            Write-Host ("         Channel URL: {0}" -f $channelURL) -ForegroundColor DarkGray
                            Write-Host ("         Parent Folder: {0}" -f $parentFolder) -ForegroundColor DarkGray
                            Write-Host ("         Destination Folder: {0}" -f $destFolder) -ForegroundColor DarkGray
                            Write-Host ("         Folder: {0}" -f $folder) -ForegroundColor DarkGray

                            $tempFolder = Add-PnPFolder -Name $folder -Folder $parentFolder -ErrorAction Stop
                        }
                        $parentFolder = ($parentFolder + "/" + $folder)
                    }
                    catch {
                        # Output Error Information
                        Write-CatErrorCSV -funcJob $job -funcAction "Creating Folder" -funcException $_.Exception -funcDetails $_.ErrorDetails -funcStackTrace $_.ScriptStackTrace -funcLogPath $errorPath
                        Continue JobLoop
                    }
                }
            }
            if (-not $tempfolder) {
                Clear-Variable folders
                Clear-Variable parentFolder
                Clear-Variable folderExists
            }
            if ($tempfolder) {
                Clear-Variable tempFolder
            }
            #endregion
        }
        #endregion #>
        
        # Log complete
        Write-Host ("      Destination Configuration Complete") -foregroundcolor "green"       
    }
    Else {
        # Output Error Information
        Write-CatErrorCSV -funcJob $job -funcAction "Connecting\Creating to Destination Site" -funcException $null -funcDetails $null -funcStackTrace $null -funcLogPath $errorPath
        Continue JobLoop
    }
}

#endregion

#region Completed Script
# Log complete
Write-Host
Write-Host ("Pre-migration Configuration Completed") -foregroundcolor "green"  

# Output Script Info
Complete-CatScript $scriptStart[0]

# Stop logging
Stop-Transcript
#endregion
