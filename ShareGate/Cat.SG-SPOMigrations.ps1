<#  Cat.SG-SPModernizeMigrations.ps1

This script automates ShareGate tenant to tenant SharePoint migrations and site modernization at a site collection level.  If the destination site collections
do not exist they will be created. This script Will configure hub sites and membership. It can also flatten and modernize classic sites or migrate them as is. It 
is currenlty leveraging an App Registration in the destination Tenant to perform configuration. ShareGate still requires a service account even if using the 
Azure ShareGate Desktop Application.

Dependencies
    -   Source & Destination Azure App Registrations
    -   Source & Destination Service Account for ShareGate
    -	ShareGate PowerShell Modules; Doesn't support PS Core 7 - Install-Module PnP.PowerShell -RequiredVersion 1.12 -Force
    -   SharePointPnPPowerShellOnline Module
    -	..\General\Cat.GeneralFunctions.psm1
    -   ..\Reporting\Cat.ReportingFunctions.psm1
    -	Input CSV File in Input Files folder
    -   User mapping file export from ShareGate

Call script with parameters
	.\Cat.SG-SPModernizeMigrations.ps1 -destTenant "tenantname" -destAppID = "7f0fd640-e811-44cc-a19f-e9bcd0e28d95" -destCertPath = "C:\Temp\appReg\PnPAppRegPowerShell.pfx" -inputFile Cat.SG-ModernSiteList-Test.csv -userFile Cat.SG-UserMapping-Test.sgum

CSV Example
    sourceURL,destinationURL,destinationAlias,siteTitle,siteOwner,secondOwner,siteTemplate,classicSubSiteParentURL,isHubSite,hubParentURL,versions,migrationType
    https://tenantname.sharepoint.com/sites/ContosoCommunications,https://tenantname.sharepoint.com/sites/ContosoHub,,Constoso Hub,user@tenantname.onmicrosoft.com,,CommunicationSite,,TRUE,,25,FULL
    https://tenantname.sharepoint.com/teams/ContosoDocuments,https://tenantname.sharepoint.com/sites/ContosoDocs,ContosoDocs,Contoso Documents,user@tenantname.onmicrosoft.com,,TeamSite,,FALSE,https://tenantname.sharepoint.com/sites/ContosoHub,25,FULL
    https://tenantname.sharepoint.com/sites/ContosoCommunications/ClassicSubsite2,https://tenantname.sharepoint.com/sites/ContosoDocs/ClassicSubsite,ClassicSubsite,TORO Classic Subsite,michelle@tenantname.onmicrosoft.com,,STS#0,https://tenantname.sharepoint.com/sites/ContosoDocs,FALSE,,25,FULL

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
    [String] $jobFile = "Cat.SG-ModernSiteList-CLIENT.csv",
    [String] $userFile = "Cat.SG-UserMapping-CLIENT.sgum",
    [string] $srcTenant = "{Source Tenant Name}",
    [String] $srcAppID = "{Azure App Registration ID}",
    [String] $srcCertPath = "C:\certs\PnP PowerShell Authentication.pfx",
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
Start-Transcript -Path ("..\_Log Files\{0}.SG-SPOModernizeMigrations-{1}.log" -f $destTenant, $scriptStart[1])
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
$timeZone = 11
$siteLocale = 1033

# User Inputs
Write-Host
Write-Host "Enter Connection Information" -ForegroundColor Cyan
$srcCertPass = ConvertTo-SecureString -String (Read-Host -Prompt "   Enter SOURCE certificate password for the app registration") -AsPlainText -Force
Write-Host "   Enter ShareGate credentials for the SOURCE SharePoint environment:"
$srcSGCreds = Get-Credential -Message "Enter ShareGate credentials for the SOURCE environment:" 
if (-not $basicAuthEnabled) {
    Write-Host "   *** Using browser based authentication, although using modern auth, can potentially result in security token timeouts on long running scripts. ***" -ForegroundColor Yellow
    Write-Host "   Authenticating to M365"
    $srcSGConnection = Connect-Site -Url $srcRootSiteUrl -Browser -DisableSSO
}
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
Write-Host ("   Source Tenant: {0}" -f $srcTenant)
Write-Host ("   Source App ID: {0}" -f $srcAppID)
Write-Host ("   Source Cert Path: {0}" -f $srcCertPath)
Write-Host ("   Destination Tenant: {0}" -f $destTenant)
Write-Host ("   Destination App ID: {0}" -f $destAppID)
Write-Host ("   Destination Cert Path: {0}" -f $destCertPath)
Write-Host
Write-Host "User Inputs" -ForegroundColor Cyan
Write-Host ("   Source Credentials: {0}" -f $srcSGCreds.UserName)
Write-Host ("   Destination Credentials: {0}" -f $destSGCreds.UserName)
Write-Host
Write-Host "Script Variables" -ForegroundColor Cyan
Write-Host ("   Source Root Site URL: {0}" -f $srcRootSiteUrl)
Write-Host ("   Source SP Admin URL: {0}" -f $srcAdminUrl)
Write-Host ("   Source Tenant Domain: {0}" -f $srcTenantDomain)
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
    Write-Host ("   Configuring & running migration job: {0}" -f $job.sourceURL) -ForegroundColor Magenta
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
        
        # Get Site Template
        $siteTemplate = $job.siteTemplate.Trim()
        Write-Host ("         Site Template: {0}" -f $siteTemplate) -Foregroundcolor DarkGray
        
        # Creating Site based on Template
        #get public/private value from CSV
        if ($job.modernTeamSiteIsPublic.ToLower() -eq "true") {$modernTeamSitePublic = $true} else {$modernTeamSitePublic = $false}

        switch ($siteTemplate) {
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
                # creating Modern Team Site
                Write-Host ("         Creating Modern Team Site: {0}" -f $job.destinationURL) -Foregroundcolor DarkGray
                $result = New-PnPSite -Type "TeamSite" -Title $job.siteTitle -Alias $job.destinationAlias -Owners $job.siteOwner -Wait -Connection $rootSiteConnection -isPublic:$modernTeamSitePublic
                Disconnect-PnPOnline -Connection $rootSiteConnection -WarningAction SilentlyContinue
            }
            {((($_.ToLower() -ne "communicationsite") -and ($_.ToLower() -ne "teamsite")) -and ($job.classicSubSiteParentURL -eq ""))} {
                Write-Host ("         Creating Classic Site: {0}" -f $job.destinationURL) -Foregroundcolor DarkGray
                $result = New-PnPTenantSite -Title $job.siteTitle -Url $job.destinationURL -Owner $job.siteOwner -Template $siteTemplate.Trim() -TimeZone $timeZone -Connection $rootSiteConnection
                Disconnect-PnPOnline -Connection $rootSiteConnection -WarningAction SilentlyContinue
            }
            {((($_.ToLower() -ne "communicationsite") -and ($_.ToLower() -ne "teamsite")) -and ($job.classicSubSiteParentURL -ne ""))} {
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
        
        #region 6.  Configuring Hubs
        # Register site as hub site if not already
        if ($job.isHubSite.ToLower() -eq "true") {
            $hubSite = Get-PnPHubSite -Identity $job.destinationURL
            if ($hubSite.SiteId) {
                Write-Host ("      Site already configured as hub: {0}" -f $job.destinationURL)   
            }
            else {
                Write-Host ("      Configuring as Hub Site: {0}" -f $job.destinationURL)            
                Register-PnPHubSite -Site $destSite.Url -Connection $destSiteConnection                         
            }
        }

        # Associate sites with the parent HUB (if specified)
        if ($job.hubParentURL) {
            Write-Host ("      Associating Site to Hub: {0} " -f $job.hubParentURL)
            Add-PnPHubSiteAssociation -Site $job.destinationURL -HubSite $job.hubParentURL -Connection $destSiteConnection      
        }
        Disconnect-PnPOnline -Connection $destSiteConnection
        #endregion

        #region 7.  Adding ShareGate Permissions
        # Adding ShareGate service account to source site
        Write-Host ("      Adding Service Account as Source Site Owner: {0}" -f $srcSGCreds.UserName)
        $srcSiteConnection = Connect-PnPOnline -ClientId $srcAppID -Url $job.sourceURL -Tenant $srcTenantDomain -CertificatePath $srcCertPath -CertificatePassword $srcCertPass -WarningAction SilentlyContinue
        Add-PnPSiteCollectionAdmin -Owners $srcSGCreds.UserName -Connection $srcSiteConnection
        Disconnect-PnPOnline -Connection $srcSiteConnection -WarningAction SilentlyContinue

        # Adding ShareGate service account to destination site
        Write-Host ("      Adding Service Account as Destination Site Owner: {0}" -f $destSGCreds.UserName)
        $destSiteConnection = Connect-PnPOnline -ClientId $destAppID -Url $destURL -Tenant $destTenantDomain -CertificatePath $destCertPath -CertificatePassword $destCertPass -WarningAction SilentlyContinue #>
        Add-PnPSiteCollectionAdmin -Owners $destSGCreds.UserName -Connection $destSiteConnection
        Disconnect-PnPOnline -Connection $destSiteConnection -WarningAction SilentlyContinue
        #endregion
                
        #region 8.  Configuring ShareGate job
        # Loading sites
        Write-Host ("      Loading Source Site: {0}" -f $job.sourceURL)
        $srcSite = Connect-Site -Url $job.sourceURL -Credential $srcSGCreds
        if ($basicAuthEnabled) {
            $srcSite = Connect-Site -Url $job.sourceURL -Credential $srcSGCreds -ErrorAction Stop
        } else {
            $srcSite = Connect-Site -Url $job.sourceURL -UseCredentialsFrom $srcSGConnection -ErrorAction Stop
        }     
        Write-Host ("      Loading Destination Site: {0}" -f $destURL)
        $dstSite = Connect-Site -Url $destURL -Credential $destSGCreds
        if ($basicAuthEnabled) {
            $dstSite = Connect-Site -Url $destURL -Credential $destSGCreds -ErrorAction Stop
        } else {
            $dstSite = Connect-Site -Url $destURL -UseCredentialsFrom $destSGConnection -ErrorAction Stop
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
        $taskName = ("SPO Migration - {0} - {1}" -f $job.siteTitle, $migrationType)

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
        Write-Host ("      Running Job: {0}" -f $destURL)
        $result = Copy-Site -Site $srcSite `
                            -DestinationSite $dstSite `
                            -merge -TaskName $taskName `
                            -CopySettings $copySettings `
                            -MappingSettings $copyMappings `
                            -VersionLimit $versionsToMigrate `
                            -NoWorkflows `
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
        
        # Clearing ShareGate Objects
        Clear-Variable srcSite
        Clear-Variable dstSite
        #endregion

        #region 10. Removing ShareGate Permissions
        # Removing ShareGate service account to source site
        Write-Host ("      Removing Service Account as Source Site Owner: {0}" -f $destSGCreds.UserName)
        $srcSiteConnection = Connect-PnPOnline -ClientId $srcAppID -Url $job.sourceURL -Tenant $srcTenantDomain -CertificatePath $srcCertPath -CertificatePassword $srcCertPass -WarningAction SilentlyContinue
        Remove-PnPSiteCollectionAdmin -Owners $srcSGCreds.UserName -Connection $srcSiteConnection
        Disconnect-PnPOnline -Connection $srcSiteConnection -WarningAction SilentlyContinue
        
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