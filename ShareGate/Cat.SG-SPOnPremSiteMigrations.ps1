<#  Cat.SG-SPOnPremSiteMigrations.ps1

This script automates ShareGate SharePoint on-premises to SharePoint Online migrations and site modernization at a site level.  If the destination site collections
do not exist they will be created. This script will configure hub sites and membership. It can also flatten and modernize classic sites or migrate them as is. It 
is currenlty leveraging an App Registration in the destination Tenant to perform configuration. ShareGate still requires a service account even if using the 
Azure ShareGate Desktop Application to improve performance.

Dependencies
    -   Destination Azure App Registrations (Catapult Powershell & ShareGate)
    -   Source & Destination Service Account for ShareGate
    -	ShareGate PowerShell Modules; Doesn't support PS Core 7
    -   PnP.PowerShell Module
    -	..\General\Cat.GeneralFunctions.psm1
    -   ..\Reporting\Cat.ReportingFunctions.psm1
    -	Input CSV File in Input Files folder
    -   User mapping file export from ShareGate

Call script with parameters
	.\Cat.SG-SPOnPremMigrations.ps1 -jobFile "Cat.SG-SPOnPremSiteMigrations-Client.csv" -userFile "Cat.SG-UserMapping-CLIENT.sgum" -destTenant nrpdev2 -destAppID "57577043-3de0-40c0-9c68-7960e2cebe8b" -destCertPath "C:\GitHub\NRPDev\AppRegistration\PnP PowerShell Authentication.pfx"

CSV Example
    sourceURL,destinationURL,destinationAlias,siteTitle,siteOwner,siteTemplate,classicSubSiteParentURL,isHubSite,hubParentURL,versions,migrationType,migrateSubSites,migrateWorkflows,migrateUserAlerts,migrateNavigation,migrateCustomPermissions,migrateContent,migrateSiteFeatures,migrateWebParts,migrateCustomFormsandViews,migrateCustomListForms,migrateListViews,forceNewListExperiece,ModernTeamSiteIsPublic
    https://portal.sp13test.local/sites/migration,https://nrpdev2.sharepoint.com/sites/MigrationTestSP13AsIs-New02,,Migration Test SP13 As Is New02,nickp@nrpdev.com,STS#0,,FALSE,,25,FULL,TRUE,FALSE,FALSE,TRUE,TRUE,TRUE,TRUE,TRUE,TRUE,TRUE,TRUE,FALSE,FALSE
    https://portal.sp13test.local/sites/migration,https://nrpdev2.sharepoint.com/sites/MigrationTestSP13ClassicBrokenUp-New02,,Migration Test SP13 Classic Boken Up New02,nickp@nrpdev.com,STS#0,,FALSE,,25,FULL,FALSE,FALSE,FALSE,TRUE,TRUE,TRUE,TRUE,TRUE,TRUE,TRUE,TRUE,FALSE,FALSE
    https://portal.sp13test.local/sites/migration/sub1,sub1,,Migration Test SP13 As Is Sub1 New02,nickp@nrpdev.com,STS#0,https://nrpdev2.sharepoint.com/sites/MigrationTestSP13AsIsBrokenUp-New02,FALSE,,5,FULL,FALSE,FALSE,FALSE,TRUE,TRUE,TRUE,TRUE,TRUE,TRUE,TRUE,TRUE,FALSE,FALSE
    https://portal.sp13test.local/sites/migration/sub2,sub2,,Migration Test SP13 As Is Sub2 New02,nickp@nrpdev.com,STS#0,https://nrpdev2.sharepoint.com/sites/MigrationTestSP13AsIsBrokenUp-New02,FALSE,,5,FULL,FALSE,FALSE,FALSE,TRUE,TRUE,TRUE,TRUE,TRUE,TRUE,TRUE,TRUE,FALSE,FALSE
    https://portal.sp13test.local/sites/migration/sub1/subsub1,subsub1,,Migration Test SP13 As Is SubSub1 New02,nickp@nrpdev.com,STS#0,https://nrpdev2.sharepoint.com/sites/MigrationTestSP13AsIsBrokenUp-New02/sub1,FALSE,,5,FULL,FALSE,FALSE,FALSE,TRUE,TRUE,TRUE,TRUE,TRUE,TRUE,TRUE,TRUE,FALSE,FALSE
    https://portal.sp13test.local/sites/migration/sub2/subsub2,subsub2,,Migration Test SP13 As Is SubSub2 New02,nickp@nrpdev.com,STS#0,https://nrpdev2.sharepoint.com/sites/MigrationTestSP13AsIsBrokenUp-New02/sub2,FALSE,,5,FULL,FALSE,FALSE,FALSE,TRUE,TRUE,TRUE,TRUE,TRUE,TRUE,TRUE,TRUE,FALSE,FALSE
    https://portal.sp13test.local/sites/migration,https://nrpdev2.sharepoint.com/sites/MigrationTestSP13ModernTeamHub-New02,MigrationTestSP13ModernTeamHub-New02,Migration Test SP13 Modern Team Hub New02,nickp@nrpdev.com,TeamSite,,TRUE,,10,FULL,FALSE,FALSE,FALSE,TRUE,TRUE,TRUE,TRUE,TRUE,FALSE,FALSE,TRUE,FALSE,TRUE
    https://portal.sp13test.local/sites/migration/sub1,https://nrpdev2.sharepoint.com/sites/MigrationTestSP13ModernTeamHubSub1-New02,MigrationTestSP13ModernTeamHubSub1-New02,Migration Test SP13 Modern Team Hub Sub1 New02,nickp@nrpdev.com,TeamSite,,FALSE,https://nrpdev2.sharepoint.com/sites/MigrationTestSP13ModernTeamHub-New02,10,FULL,FALSE,FALSE,FALSE,TRUE,TRUE,TRUE,TRUE,TRUE,FALSE,FALSE,TRUE,FALSE,TRUE
    https://portal.sp13test.local/sites/migration/sub2,https://nrpdev2.sharepoint.com/sites/MigrationTestSP13ModernTeamHubSub2-New02,MigrationTestSP13ModernTeamHubSub2-New02,Migration Test SP13 Modern Team Hub Sub2 New02,nickp@nrpdev.com,TeamSite,,FALSE,https://nrpdev2.sharepoint.com/sites/MigrationTestSP13ModernTeamHub-New02,10,FULL,FALSE,FALSE,FALSE,TRUE,TRUE,TRUE,TRUE,TRUE,FALSE,FALSE,TRUE,FALSE,TRUE
    https://portal.sp13test.local/sites/migration/sub1/subsub1,https://nrpdev2.sharepoint.com/sites/MigrationTestSP13ModernTeamHubSubSub1-New02,MigrationTestSP13ModernTeamHubSubSub1-New02,Migration Test SP13 Modern Team Hub SubSub1 New02,nickp@nrpdev.com,TeamSite,,FALSE,https://nrpdev2.sharepoint.com/sites/MigrationTestSP13ModernTeamHub-New02,10,FULL,FALSE,FALSE,FALSE,TRUE,TRUE,TRUE,TRUE,TRUE,FALSE,FALSE,TRUE,FALSE,TRUE
    https://portal.sp13test.local/sites/migration/sub2/subsub2,https://nrpdev2.sharepoint.com/sites/MigrationTestSP13ModernTeamHubSubSub2-New02,MigrationTestSP13ModernTeamHubSubSub2-New02,Migration Test SP13 Modern TeamHub SubSub2 New02,nickp@nrpdev.com,TeamSite,,FALSE,https://nrpdev2.sharepoint.com/sites/MigrationTestSP13ModernTeamHub-New02,10,FULL,FALSE,FALSE,FALSE,TRUE,TRUE,TRUE,TRUE,TRUE,FALSE,FALSE,TRUE,FALSE,TRUE

CSV Notes
    This script will create Modern Communication and Teams sites along with "classic" subsites.
    The CSV Example above outlines how to define these three different examples.

Table of Contents
	1.	Snapin(s), Modules & Environment Connection(s)
	2.	Input & Output Files
	3.	Variables
    4.	Main Script
    5.  Site Creation
    6.  Configuring Hubs
    7.  Adding ShareGate Permissions
    8.  Configuring ShareGate job
    9.  Running ShareGate Migration
    10. Removing ShareGate Permissions
#>

#region Script Parameters
Param(
    [String] $jobFile = "Cat.SG-SPOnPremMigrations-Client.csv",
    [String] $userFile = "Cat.SG-UserMapping-CLIENT.sgum",
    [string] $destTenant = "{Destination Tenant Name}",
    [String] $destAppID = "{Azure App Registration ID}",
    [String] $destCertPath = "C:\certs\PnPAppRegPowerShell.pfx",
    [bool] $basicAuthEnabled = $false, # If $false, the script with leverage browser authentication with ShareGate. This uses modern authentication. The security tokens can potenitally time out.
    [bool] $useCustomSiteTemplates = $true, # If true create and store the job's custom xml site templates in the input files folder (i.e. Cat.SG-TeamSiteTemplate.xml, Cat.SG-CommunicationSiteTemplate.xml)
    [String] $teamSiteTemplateFile = "Cat.SG-TeamSiteTemplate.xml",
    [String] $commSiteTemplateFile = "Cat.SG-CommunicationSiteTemplate.xml",
    [String] $msteamTemplateFile = "Cat.SG-MSTeamsTemplate.xml"
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
Start-Transcript -Path ("..\_Log Files\{0}.SG-SPOnPremSiteMigrations-{1}.log" -f $destTenant, $scriptStart[1])
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
$errorPath = ("..\_Output Files\{0}.SG-SPOnPremSiteMigrations-{1}.log" -f $destTenant, $scriptStart[1])

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
$siteWaitPause = 15
$siteWaitMaxCount = 40
$timeZone = 11
$siteLocale = 1033
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
Write-Host ("   Wait Between Jobs (Seconds): {0}" -f $siteWaitPause)
Write-Host ("   Wait Max Attempts: {0}" -f $siteWaitMaxCount)
Write-Host ("   Time Zone: {0}" -f $timeZone)
Write-Host ("   Local: {0}" -f $siteLocale)
#endregion
    
#region 4.	Main Script
Write-Host
Write-Host ("Configuring & running migrations for job list: {0}" -f $jobFile) -ForegroundColor Cyan

# Looping through input file
:JobLoop ForEach ($job in $inFile){
    Write-Host
    Write-Host ("   {0} - Configuring & running migration job: {1}" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"), $job.sourceURL) -ForegroundColor Magenta
    <# Output Job Configuration for Troubleshooting (if Needed remove "<" on this line)
    $job #>

    #region 5.  Site Creation
    # Determine Destination URL
    if ($job.destinationURL -like "https://*") {$destURL = $job.destinationURL.Trim()} else {$destURL = ("{0}/{1}" -f $job.classicSubSiteParentURL, $job.destinationURL)}

    # Determin if destination site exists
    Write-Host ("      Checking if Site Exists: {0}" -f $destURL)
    $destSiteExists = Test-CatPnPSiteExistsAppReg -fnUrl $destURL -fnAppID $destAppID -fnTenant $destTenantDomain -fnCertPath $destCertPath -fnCertPass $destCertPass
    
    # Create Site if it doesn't exist
    If ($destSiteExists){ 
        Write-Host ("         Site Already Exists") -Foregroundcolor DarkGray
    }
    Else {
        Write-Host ("         Site Doesn't Exists") -Foregroundcolor DarkGray
        
        # Creating destination site collection
        Write-Host ("      Creating Destination Site: {0}" -f $job.destinationURL)

        # Creating Site based on Template
        #get public/private value from CSV
        if ($job.modernTeamSiteIsPublic.ToLower() -eq "true") {$modernTeamSitePublic = $true} else {$modernTeamSitePublic = $false}
        try {
            switch ($siteTemplate.ToLower()) {
                {((($_.ToLower() -eq "communicationsite") -or ($_.ToLower() -eq "teamsite")) -or ((($_.ToLower() -ne "communicationsite") -and ($_.ToLower() -ne "teamsite")) -and ($job.classicSubSiteParentURL -eq "")))} {
                    # Connecting to Root Site
                    Write-Host ("         Connecting to Root Site: {0}" -f $destRootSiteUrl) -Foregroundcolor DarkGray
                    $rootSiteConnection = Connect-PnPOnline -ClientId $destAppID -Url $destRootSiteUrl -Tenant $destTenantDomain -CertificatePath $destCertPath -CertificatePassword $destCertPass -WarningAction SilentlyContinue
                    Write-Host ("         Connected to Root Site: {0}" -f $destRootSiteUrl) -Foregroundcolor DarkGray
                }
                communicationsite {
                    # Creating Modern Communications Site
                    Write-Host ("         Creating Communication Site: {0}" -f $job.destinationURL) -Foregroundcolor DarkGray
                    $result = New-PnPSite -Type CommunicationSite -Title $job.siteTitle -Url $job.destinationURL -Owner $job.siteOwner -Wait -Connection $rootSiteConnection -SiteDesign Topic
                    Disconnect-PnPOnline -Connection $rootSiteConnection -WarningAction SilentlyContinue
                }
                teamsite {
                    # Creating Modern Team Site
                    Write-Host ("         Creating Modern Team Site: {0}" -f $job.destinationURL) -Foregroundcolor DarkGray
                    $result = New-PnPSite -Type "TeamSite" -Title $job.siteTitle -Alias $job.destinationAlias -Owners $job.siteOwner -Wait -Connection $rootSiteConnection -isPublic:$modernTeamSitePublic
                    Disconnect-PnPOnline -Connection $rootSiteConnection -WarningAction SilentlyContinue
                }
                msteams {
                    # Creating Modern Microsoft 365 Teams
                    Write-Host ("         Creating M365 Teams: {0}" -f $job.destinationURL) -Foregroundcolor DarkGray
                    $result = New-PnPTeamsTeam -DisplayName $job.siteTitle
                    Disconnect-PnPOnline -Connection $rootSiteConnection -WarningAction SilentlyContinue
                }
                {((($_.ToLower() -ne "communicationsite") -and ($_.ToLower() -ne "teamsite") -and ($_.ToLower() -ne "msteams")) -and ($job.classicSubSiteParentURL -eq ""))} {
                    Write-Host ("         Creating Classic Site: {0}" -f $job.destinationURL) -Foregroundcolor DarkGray
                    $result = New-PnPTenantSite -Title $job.siteTitle -Url $job.destinationURL -Owner $job.siteOwner -Template $siteTemplate.Trim() -TimeZone $timeZone -Connection $rootSiteConnection
                    Disconnect-PnPOnline -Connection $rootSiteConnection -WarningAction SilentlyContinue
                }
                {((($_.ToLower() -ne "communicationsite") -and ($_.ToLower() -ne "teamsite") -and ($_.ToLower() -ne "msteams")) -and ($job.classicSubSiteParentURL -ne ""))} {
                    Write-Host ("         Connecting to Parent Site: {0}" -f $job.classicSubSiteParentURL) -Foregroundcolor DarkGray
                    $subsiteConnection = Connect-PnPOnline -ClientId $destAppID -Url $job.classicSubSiteParentURL.Trim() -Tenant $destTenantDomain -CertificatePath $destCertPath -CertificatePassword $destCertPass -WarningAction SilentlyContinue
                    Write-Host ("         Creating SharePoint Classic Sub Site: {0}" -f $destURL) -Foregroundcolor DarkGray
                    $result = New-PnPWeb -Title $job.siteTitle -Url $job.destinationURL.Trim() -Locale $siteLocale -Template $siteTemplate.Trim() -Connection $subsiteConnection
                    Disconnect-PnPOnline -Connection $subsiteConnection -WarningAction SilentlyContinue
                }
                Default {
                    Write-Host ("         There are issues with site creation: {0} - Migration Canceled" -f $job.destinationURL) -Foregroundcolor Red
                    Continue JobLoop
                }
            }
        }
        catch {
            # Output Error Information
            Write-Host ("         {0} - Error Creating Site: {1}" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"),$_.Exception) -ForegroundColor Red
            $err = $job
            $err | Add-Member -MemberType NoteProperty -Name "errorAction" -Value "Site Creation"
            $err | Add-Member -MemberType NoteProperty -Name "errorException" -Value $_.Exception
            $err | Add-Member -MemberType NoteProperty -Name "errorDetails" -Value $_.ErrorDetails
            $err | Add-Member -MemberType NoteProperty -Name "errorStackTrace" -Value $_.ScriptStackTrace
            $err | Export-Csv -Path $errorPath -NoTypeInformation -Append
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
        Write-Host ("         Created Destination Site: {0}" -f $destURL) -ForegroundColor Green
    }
    #endregion #>

    # Run Migrations if destination site exists
    If ($destSiteExists){

        # Connecting to Destination Site
        $destSiteConnection = Connect-PnPOnline -ClientId $destAppID -Url $destURL -Tenant $destTenantDomain -CertificatePath $destCertPath -CertificatePassword $destCertPass -WarningAction SilentlyContinue #>
        $destSite = Get-PnPSite -Includes RootWeb    

        #region 5.1 Apply Site Template
        If($useCustomSiteTemplates) {
            try {
                switch ($siteTemplate.ToLower()) {               
                    communicationsite {
                        # applying Modern Communications Site template
                        Write-Host ("         Applying Communication Site Template: {0}" -f $job.destinationURL) -Foregroundcolor DarkGray
                        $applyTemplateResults = Invoke-PnPSiteTemplate -Path "..\_Input Files\{0}\" -f $commSiteTemplateFile
                    }
                    teamsite {
                        # creating Modern Team Site template
                        Write-Host ("         Applying Team Site Template: {0}" -f $job.destinationURL) -Foregroundcolor DarkGray
                        $applyTemplateResults = Invoke-PnPSiteTemplate -Path "..\_Input Files\{0}\" -f $teamSiteTemplateFile

                    }
                    msteams {
                        # creating Modern Team Site template
                        Write-Host ("         Applying Modern Team Template: {0}" -f $job.destinationURL) -Foregroundcolor DarkGray
                        $applyTemplateResults = Invoke-PnPSiteTemplate -Path "..\_Input Files\{0}\" -f $msteamTemplateFile

                    }
                    Default {
                        Write-Host ("         There are issues with applying template: {0} - Migration Canceled" -f $job.destinationURL) -Foregroundcolor Red
                        Continue JobLoop
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
        
        #region 6.  Configuring Hubs
        # Register site as hub site if not already
        if ($job.isHubSite.ToLower() -eq "true") {
            $hubSite = Get-PnPHubSite -Identity $job.destinationURL
            if ($hubSite.SiteId) {
                Write-Host ("      Site already configured as hub: {0}" -f $job.destinationURL)   
            }
            else {
                Write-Host ("      Configuring as Hub Site: {0}" -f $job.destinationURL)            
                try {
                    Register-PnPHubSite -Site $destSite.Url -Connection $destSiteConnection
                }
                catch {
                    # Output Error Information
                    Write-Host ("         {0} - Error Configuring Hub: {1}" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"),$_.Exception) -ForegroundColor Red
                    $err = $job
                    $err | Add-Member -MemberType NoteProperty -Name "errorAction" -Value "Hub Configuration"
                    $err | Add-Member -MemberType NoteProperty -Name "errorException" -Value $_.Exception
                    $err | Add-Member -MemberType NoteProperty -Name "errorDetails" -Value $_.ErrorDetails
                    $err | Add-Member -MemberType NoteProperty -Name "errorStackTrace" -Value $_.ScriptStackTrace
                    $err | Export-Csv -Path $errorPath -NoTypeInformation -Append
                }                        
            }
        }

        # Associate sites with the parent HUB (if specified)
        if ($job.hubParentURL) {
            Write-Host ("      Associating Site to Hub: {0} " -f $job.hubParentURL)
            try {
                Add-PnPHubSiteAssociation -Site $job.destinationURL -HubSite $job.hubParentURL -Connection $destSiteConnection
            }
            catch {
                # Output Error Information
                Write-Host ("         {0} - Error Associating Hub: {1}" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"),$_.Exception) -ForegroundColor Red
                $err = $job
                $err | Add-Member -MemberType NoteProperty -Name "errorAction" -Value "Hub Association"
                $err | Add-Member -MemberType NoteProperty -Name "errorException" -Value $_.Exception
                $err | Add-Member -MemberType NoteProperty -Name "errorDetails" -Value $_.ErrorDetails
                $err | Add-Member -MemberType NoteProperty -Name "errorStackTrace" -Value $_.ScriptStackTrace
                $err | Export-Csv -Path $errorPath -NoTypeInformation -Append
            }    
        }
        #endregion

        #region 7.  Adding ShareGate Permissions
        # Adding ShareGate service account to destination site
        Write-Host ("      Adding Service Account as Destination Site Owner: {0}" -f $destSGCreds.UserName)
        try {
            Add-PnPSiteCollectionAdmin -Owners $destSGCreds.UserName -Connection $destSiteConnection
        }
        catch {
            # Output Error Information
            Write-Host ("         {0} - Error Adding SG Account: {1}" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"),$_.Exception) -ForegroundColor Red
            $err = $job
            $err | Add-Member -MemberType NoteProperty -Name "errorAction" -Value "Adding SG Account"
            $err | Add-Member -MemberType NoteProperty -Name "errorException" -Value $_.Exception
            $err | Add-Member -MemberType NoteProperty -Name "errorDetails" -Value $_.ErrorDetails
            $err | Add-Member -MemberType NoteProperty -Name "errorStackTrace" -Value $_.ScriptStackTrace
            $err | Export-Csv -Path $errorPath -NoTypeInformation -Append
        }
        #endregion

        # Disconnecting from Destination Site
        Disconnect-PnPOnline -Connection $destSiteConnection -WarningAction SilentlyContinue
                
        #region 8.  Configuring ShareGate job
        # Loading sites
        Write-Host ("      Loading Source Site: {0}" -f $job.sourceURL)
        try {
            $srcSite = Connect-Site -Url $job.sourceURL -Credential $srcSGCreds
        }
        catch {
            # Output Error Information
            Write-Host ("         {0} - Error Loading Source: {1}" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"),$_.Exception) -ForegroundColor Red
            $err = $job
            $err | Add-Member -MemberType NoteProperty -Name "errorAction" -Value "Loading Source"
            $err | Add-Member -MemberType NoteProperty -Name "errorException" -Value $_.Exception
            $err | Add-Member -MemberType NoteProperty -Name "errorDetails" -Value $_.ErrorDetails
            $err | Add-Member -MemberType NoteProperty -Name "errorStackTrace" -Value $_.ScriptStackTrace
            $err | Export-Csv -Path $errorPath -NoTypeInformation -Append
        }  
        Write-Host ("      Loading Destination Site: {0}" -f $destURL)
        try {
            if ($basicAuthEnabled) {
                $dstSite = Connect-Site -Url $destURL -Credential $destSGCreds -ErrorAction Stop
            } else {
                $dstSite = Connect-Site -Url $destURL -UseCredentialsFrom $destSGConnection -ErrorAction Stop
            }
        }
        catch {
            # Output Error Information
            Write-Host ("         {0} - Error Loading Destination: {1}" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"),$_.Exception) -ForegroundColor Red
            $err = $job
            $err | Add-Member -MemberType NoteProperty -Name "errorAction" -Value "Loading Destination"
            $err | Add-Member -MemberType NoteProperty -Name "errorException" -Value $_.Exception
            $err | Add-Member -MemberType NoteProperty -Name "errorDetails" -Value $_.ErrorDetails
            $err | Add-Member -MemberType NoteProperty -Name "errorStackTrace" -Value $_.ScriptStackTrace
            $err | Export-Csv -Path $errorPath -NoTypeInformation -Append
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

        # Setting Job Name
        Write-Host ("         Configuring job name") -ForegroundColor DarkGray
        $taskName = ("SP On Prem Site Migration - {0} - {1}" -f $job.siteTitle, $migrationType)

        # Updating Job Settings based on CSV 
        $versionsToMigrate = $job.versions        
        if ($job.migrateSubSites.ToLower() -eq "true") {$migrateSubSites = $true} else {$migrateSubSites = $false}
        if ($job.migrateWorkflows.ToLower() -eq "true") {$noWorkflows = $false} else {$noWorkflows = $true}
        if ($job.migrateUserAlerts.ToLower() -eq "true") {$migrateUserAlerts = $true} else {$migrateUserAlerts = $false}                
        if ($job.migrateNavigation.ToLower() -eq "true") {$noNavigation = $false} else {$noNavigation = $true} 
        if ($job.migrateCustomPermissions.ToLower() -eq "true") {$noCustomPermissions = $false} else {$noCustomPermissions = $true}
        if ($job.migrateContent.ToLower() -eq "true") {$noContent = $false} else {$noContent = $true}
        if ($job.migrateSiteFeatures.ToLower() -eq "true") {$noSiteFeatures = $false} else {$noSiteFeatures = $true}
        if ($job.migrateWebParts.ToLower() -eq "true") {$noWebParts = $false} else {$noWebParts = $true}
        if ($job.migrateCustomListForms.ToLower() -eq "true") {$noCustListForms = $false} else {$noCustListForms = $true}
        if ($job.migrateListViews.ToLower() -eq "true") {$noListViews = $false} else {$noListViews = $true}
        if ($job.forceNewListExperiece.ToLower() -eq "true") {$forceNewListExperience = $true} else {$forceNewListExperience = $false}
        #endregion

        #region 9.  Running ShareGate Migration
        Write-Host ("      {0} - Running Job: {1}" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"), $destURL)
        try {
            $result = Copy-Site -Site $srcSite `
                                -DestinationSite $dstSite `
                                -merge -TaskName $taskName `
                                -CopySettings $copySettings `
                                -MappingSettings $copyMappings `
                                -VersionLimit $versionsToMigrate `
                                -NoWorkflows:$noWorkflows `
                                -NoContent:$noContent `
                                -Subsites:$migrateSubSites `
                                -UserAlerts:$migrateUserAlerts `
                                -NoNavigation:$noNavigation `
                                -NoCustomPermissions:$noCustomPermissions `
                                -NoSiteFeatures:$noSiteFeatures `
                                -NoWebParts:$noWebParts `
                                -NoCustomizedListForms:$noCustListForms `
                                -NoListViews:$noListViews `
                                -ForceNewListExperience:$forceNewListExperience
        }
        catch {
            # Output Error Information
            Write-Host ("         {0} - Error Running Migration: {1}" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"),$_.Exception) -ForegroundColor Red
            $err = $job
            $err | Add-Member -MemberType NoteProperty -Name "errorAction" -Value "Running Migration"
            $err | Add-Member -MemberType NoteProperty -Name "errorException" -Value $_.Exception
            $err | Add-Member -MemberType NoteProperty -Name "errorDetails" -Value $_.ErrorDetails
            $err | Add-Member -MemberType NoteProperty -Name "errorStackTrace" -Value $_.ScriptStackTrace
            $err | Export-Csv -Path $errorPath -NoTypeInformation -Append
        }

        # Clearing ShareGate Objects
        Clear-Variable srcSite
        Clear-Variable dstSite
        #endregion

        #region 10. Removing ShareGate Permissions        
        # Removing ShareGate service account to destination site
        Write-Host ("      Removing Service Account as Destination Site Owner: {0}" -f $destSGCreds.UserName)
        $destSiteConnection = Connect-PnPOnline -ClientId $destAppID -Url $destURL -Tenant $destTenantDomain -CertificatePath $destCertPath -CertificatePassword $destCertPass -WarningAction SilentlyContinue
        try {
            Remove-PnPSiteCollectionAdmin -Owners $destSGCreds.UserName -Connection $destSiteConnection
        }
        catch {
            # Output Error Information
            Write-Host ("         {0} - Error Removing SG Account: {1}" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"),$_.Exception) -ForegroundColor Red
            $err = $job
            $err | Add-Member -MemberType NoteProperty -Name "errorAction" -Value "Removing SG Account"
            $err | Add-Member -MemberType NoteProperty -Name "errorException" -Value $_.Exception
            $err | Add-Member -MemberType NoteProperty -Name "errorDetails" -Value $_.ErrorDetails
            $err | Add-Member -MemberType NoteProperty -Name "errorStackTrace" -Value $_.ScriptStackTrace
            $err | Export-Csv -Path $errorPath -NoTypeInformation -Append
        }
        Disconnect-PnPOnline -Connection $destSiteConnection -WarningAction SilentlyContinue
        #endregion
        
        # Log complete
        Write-Host ("      Completed Successfully: {0}" -f $destURL) -foregroundcolor "green"       
    }
    Else {
        Write-Host ("      Site not created: {0}" -f $destURL) -ForegroundColor "red"
        Continue JobLoop
    }
}

# Script complete output
Write-Host
Write-Host ("SharePoint Migration completed. Please refer to the ShareGate Job Logs for any warnings or errors that require remediation") -ForegroundColor Green
#endregion

#region Completed Script
# Output Script Info
Complete-CatScript $scriptStart[0]

# Stop logging
Stop-Transcript
#endregion