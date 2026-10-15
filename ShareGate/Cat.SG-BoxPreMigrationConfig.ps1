<#  Cat.SG-BoxPreMigrationConfig.ps1

This script automates the configuration of OneDrive and SharePoint Online to prepare for migration box folders.

Dependencies
    -   Destination Azure App Registrations (Catapult Powershell & ShareGate)
    -   Destination Service Account for ShareGate
    -	ShareGate PowerShell Modules (Doesn't support PS Core 7)
    -   PnP.PowerShell Module
        -   Install-Module -Name "PnP.PowerShell" -RequiredVersion 1.12.0 - Can't use Version 2 because ShareGate doesn't support PowerShell 7
        -   $env:PNPPOWERSHELL_UPDATECHECK="false" - To supress warnings for updates
    -	..\General\Cat.GeneralFunctions.psm1
    -   ..\Reporting\Cat.ReportingFunctions.psm1
    -	Input CSV File in Input Files folder (Same as migration CSV)

Call script with parameters
    cd C:\GitHub\CatPowerShell\ShareGate\
	.\Cat.SG-BoxPreMigrationConfig.ps1 -jobFile "Samples\ShareGate\Cat.SG-BoxFolderMigrations-Test1.csv" -dstTenant "nrpdev2" -dstAppID "caf6b11a-cf03-416c-b0f0-67a09a8e30ee" -dstCertPath "C:\GitHub\NRPDev\AppRegistration\PowerShell App Registration.pfx" -timeZone 12 -siteLocale 1033

Table of Contents
	1.	Snapin(s), Modules & Environment Connection(s)
    2.	Input & Output Files
	3.	Variables
    4.	Main Script
    5.  Provision Site / OneDrive
    6.  Creating SharePoint Site
    7.  Provisioning OneDrive for User
    8.  Confirming Site Exists
    9.  Configuring Hubs
    10.  Associating to Hub
    11.  Adding ShareGate Permissions
    12. Create/Confirm Destination Library
    13.  Create/Confirm Destination Folder
#>

#region Script Parameters
Param(
    [String] $jobFile = "Cat.SG-SPOnPremListMigrations-Client.csv",
    [string] $dstTenant = "{Destination Tenant Name}",
    [String] $dstAppID = "{Azure App Registration ID}",
    [String] $dstCertPath = "C:\certs\PnPAppRegPowerShell.pfx",
    [String] $timeZone = 12, # 13 = Pacific Time (US & Canada); 38 = Arizona; 12 = Mountain Time (US & Canada); 11 = Central Time (US & Canada); 10 = Eastern Time (US & Canada)
    [String] $siteLocale = 1033 # 1033 = English - United States
)
#endregion

#region Start Date / Time
# Remove Existing Catapult Modules to ensure PS session has most up to date functions
$env:PNPPOWERSHELL_UPDATECHECK="false"
Get-Module Cat.* | Remove-Module

# Load General Functions
Import-Module "..\General\Cat.GeneralFunctions.psm1"

# Get Start Date / Time
$scriptStart = Start-CatScript

# Start Transcript
Start-Transcript -Path ("..\_Log Files\{0}.SG-BoxPreMigrationConfig-{1}.log" -f $dstTenant, $scriptStart[1])
#endregion

#region 1.	Snapin(s), Modules & Environment Connection(s)
# Load Sharegate modules
Import-Module Sharegate	

# Load PnP Modules
Get-Module -ListAvailable | Where-Object{($_.Name -like "*SharePoint*") -or ($_.Name -like "*PnP*") } | Remove-Module
Import-Module PnP.PowerShell -WarningAction SilentlyContinue

# Load Custom Modules
Import-Module "..\Reporting\Cat.ReportingFunctions.psm1"
#endregion

#region 2.	Input & Output Files
# File Paths
$inPath = ("..\_Input Files\{0}" -f $jobFile)
$errorPath = ("..\_Output Files\{0}.SG-BoxPreMigrationConfigErrors-{1}.log" -f $dstTenant, $scriptStart[1])

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
$dstAdminUrl = ("https://{0}-admin.sharepoint.com/" -f $dstTenant)
$dstRootSiteUrl = ("https://{0}.sharepoint.com/" -f $dstTenant)
$dstTenantDomain = ("{0}.onmicrosoft.com" -f $dstTenant)
$siteWaitMaxCount = 20
$siteWaitPause = 30 # Seconds
$specialCharacters = "[~``!@#$%^&*()_-]"

# User Inputs
Write-Host
Write-Host "User Input Variables" -ForegroundColor Cyan
Write-Host "Enter ShareGate credentials for the DESTINATION environment:"
$dstSGCreds = Get-Credential -Message "Enter ShareGate credentials for the DESTINATION environment:"
$dstCertPass = ConvertTo-SecureString -String (Read-Host -Prompt "   Enter certificate password for the app registration") -AsPlainText -Force
$dstCert = Get-PnPAzureCertificate -Path $dstCertPath -Password $dstCertPass
 
#region Variable Output
Write-Host
Write-Host "Script Parameters" -ForegroundColor Cyan
Write-Host ("   Job File: {0}" -f $jobFile)
Write-Host ("   Destination Tenant: {0}" -f $dstTenant)
Write-Host ("   Destination App ID: {0}" -f $dstAppID)
Write-Host ("   Time Zone: {0}" -f $timeZone)
Write-Host ("   Site Locale: {0}" -f $siteLocale)
Write-Host
Write-Host "Script Variables" -ForegroundColor Cyan
Write-Host ("   Destination Admin Url: {0}" -f $dstAdminUrl)
Write-Host ("   Destination Root Site Url: {0}" -f $dstRootSiteUrl)
Write-Host ("   Destination Tenant Url: {0}" -f $dstTenantDomain)
Write-Host ("   Site Creation Wait Max Iterations: {0}" -f $siteWaitMaxCount)
Write-Host ("   Site Creation Wait Pause Duration: {0}" -f $siteWaitPause)
Write-Host ("   Special Characters {0}" -f $specialCharacters)
<#Write-Host
Write-Host "App Reg Details   ** Testing Only **" -ForegroundColor Cyan
Write-Host ("   App Reg ID: {0}" -f $dstAppID)
Write-Host ("   App Reg Thumbprint: {0}" -f $dstCert.Thumbprint) #>
#endregion
#endregion
    
#region 4.	Main Script
Write-Host
Write-Host ("Configuring SharePoint and OneDrive for box migrations: {0}" -f $jobFile) -ForegroundColor Cyan

#region Connecting to SharePoint Online Administration
Write-Host ("   {0} - Connecting to SharePoint Online Administration - {1}" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"), $dstAdminUrl)
try {
	$dstAuth = Connect-Site -Url $dstAdminUrl -Browser -DisableSSO -ErrorAction Stop
}
catch {
	Write-CatErrorCSV -funcCSVItem $job -funcAction "Connecting to SharePoint Online Administration" -funcException $_.Exception -funcDetails $_.ErrorDetails -funcStackTrace $_.ScriptStackTrace -funcFilePath $errorPath
	throw ("Connecting to SharePoint Online Administration: {0}" -f $dstAdminUrl)	
}
#endregion

# Looping through input file
:JobLoop ForEach ($job in $inFile){
    # Determine Destination URL
    Write-Host
    if ($job.dstURL -like "https://*") {$destURL = $job.dstURL.Trim()} else {$destURL = ("{0}/{1}" -f $job.classicSubSiteParentURL, $job.dstURL)}
    Write-Host ("   {0} - Configuring Job - URL: {1}" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"), $destURL) -ForegroundColor Magenta
    #$job | fl # Use this to output the job information for troubleshooting

    #region 5.  Provision Site / OneDrive
    # Determine if destination site exists
    Write-Host "      Checking if Site Exists" -NoNewline
    $dstSiteExists = Test-CatPnPSiteExistsAppReg -fnUrl $destURL -fnAppID $dstAppID -fnTenant $dstTenantDomain -fnCertPath $dstCertPath -fnCertPass $dstCertPass
    
    # Create Site if it doesn't exist
    If ($dstSiteExists){ 
        Write-Host (" - Site Already Exists") -Foregroundcolor Green
    }
    Else {
        Write-Host (" - Site Doesn't Exists") -Foregroundcolor DarkGray
        
        # SharePoint Vs. OneDrive Destination
        If ($job.dstType -eq "SPO"){
            #region 6.  Creating SharePoint Site
            # Creating destination site collection
            Write-Host "      Creating Destination Site"
            
            # Get Site Template
            $siteTemplate = $job.siteTemplate.Trim()
            Write-Host ("         Site Template: {0}" -f $siteTemplate) -Foregroundcolor DarkGray
            
            # Creating Site based on Template
            if ($job.modernTeamSiteIsPublic.ToLower() -eq "true") {$modernTeamSitePublic = $true} else {$modernTeamSitePublic = $false}

            try {
                switch ($siteTemplate) {
                    # Modern SharePoint Team Site, Modern SharePoint Communciation Site or New Classic Site Collection
                    {((($_.ToLower() -eq "communicationsite") -or ($_.ToLower() -eq "teamsite")) -or ((($_.ToLower() -ne "communicationsite") -and ($_.ToLower() -ne "teamsite")) -and ($job.classicSubSiteParentURL -eq "")))} {
                        # Connecting to Root Site
                        Write-Host ("         Connecting to Root Site: {0}" -f $dstRootSiteUrl) -Foregroundcolor DarkGray -NoNewline
                        $rootSiteConnection = Connect-PnPOnline -ClientId $dstAppID -Url $dstRootSiteUrl -Tenant $dstTenantDomain -Thumbprint $dstCert.Thumbprint -ErrorAction Stop
                        Write-Host " - Connected" -Foregroundcolor DarkGray
                    }
                    communicationsite {
                        # Creating Modern Communications Site
                        Write-Host ("         Creating Communication Site: {0}" -f $job.dstSiteTitle) -Foregroundcolor DarkGray
                        $result = New-PnPSite -Type $siteTemplate -Title $job.dstSiteTitle -Url $destURL -Owner $job.siteOwner -Wait -SiteDesign $job.CommSiteDesign -ErrorAction Stop
                        Disconnect-PnPOnline -WarningAction SilentlyContinue
                    }
                    teamsite {
                        # creating Modern Team Site
                        Write-Host ("         Creating Modern Team Site: {0}" -f $job.dstSiteTitle) -Foregroundcolor DarkGray
                        $result = New-PnPSite -Type $siteTemplate -Title $job.dstSiteTitle -Alias $job.dstGroupAlias -Owners $job.siteOwner -isPublic:$modernTeamSitePublic -Wait -ErrorAction Stop
                        Disconnect-PnPOnline -WarningAction SilentlyContinue
                        Clear-Variable modernTeamSitePublic
                    }
                    # New Classic Site Collection
                    {((($_.ToLower() -ne "communicationsite") -and ($_.ToLower() -ne "teamsite")) -and ($job.classicSubSiteParentURL -eq ""))} {
                        Write-Host ("         Creating Classic Site: {0}" -f $job.dstSiteTitle) -Foregroundcolor DarkGray
                        $result = New-PnPTenantSite -Title $job.dstSiteTitle -Url $destURL -Owner $job.siteOwner -Template $siteTemplate -TimeZone $timeZone -ErrorAction Stop
                        Disconnect-PnPOnline -WarningAction SilentlyContinue
                    }
                    # New Classic Web / Sub Site
                    {((($_.ToLower() -ne "communicationsite") -and ($_.ToLower() -ne "teamsite")) -and ($job.classicSubSiteParentURL -ne ""))} {
                        Write-Host ("         Connecting to Parent Site: {0}" -f $job.classicSubSiteParentURL) -Foregroundcolor DarkGray -NoNewline
                        $subsiteConnection = Connect-PnPOnline -ClientId $dstAppID -Url $job.classicSubSiteParentURL -Tenant $dstTenantDomain -Thumbprint $dstCert.Thumbprint -ErrorAction Stop
                        Write-Host " - Connected" -Foregroundcolor DarkGray
                        Write-Host ("         Creating SharePoint Classic Sub Site: {0}" -f $job.dstSiteTitle) -Foregroundcolor DarkGray
                        $result = New-PnPWeb -Title $job.dstSiteTitle -Url $destURL -Locale $siteLocale -Template $siteTemplate -ErrorAction Stop
                        Disconnect-PnPOnline -WarningAction SilentlyContinue
                        Clear-Variable subsiteConnection
                    }
                    Default {
                        Write-Host ("         There are issues with site creation related to the site template: {0} - Migration Canceled" -f $job.dstSiteTitle) -Foregroundcolor Red
                        Continue JobLoop
                    }
                }
            }
            catch {
                # Output Error Information
                Write-CatErrorCSV -funcCSVItem $job -funcAction "Creating Site" -funcException $_.Exception -funcDetails $_.ErrorDetails -funcStackTrace $_.ScriptStackTrace -funcFilePath $errorPath
                Continue JobLoop
            } 
            Clear-Variable siteTemplate
            Clear-Variable rootSiteConnection
            #endregion
        } else {
            #region 7.  Provisioning OneDrive for User
            Write-Host ("      Provisioning User OneDrive: {0}" -f $job.dstEmail)
            $result = Get-OneDriveUrl -Tenant $dstAuth -Email $job.dstEmail -ProvisionIfRequired
            #endregion
        }
        
        #region 8.  Confirming Site Exists
        $counter = 0
        Write-Host "         Confirming Site Exists" -NoNewline -ForegroundColor DarkGray
        while (($dstSiteExists -eq $false) -and $counter -lt $siteWaitMaxCount) {
            $counter++
            $dstSiteExists = Test-CatPnPSiteExistsAppReg -fnUrl $destURL -fnAppID $dstAppID -fnTenant $dstTenantDomain -fnCertPath $dstCertPath -fnCertPass $dstCertPass
            if (-not $dstSiteExists) {
                Write-Host "." -NoNewline -ForegroundColor DarkGray
                start-sleep -Seconds $siteWaitPause
            }
        }
        Write-Host "." -ForegroundColor DarkGray
        #endregion
    }
    #endregion #>

    # Additional Configuration if destination site exists
    If ($dstSiteExists){
        # Connecting to Destination Site
        $dstSiteConnection = Connect-PnPOnline -ClientId $dstAppID -Url $destURL -Tenant $dstTenantDomain -Thumbprint $dstCert.Thumbprint
        $dstSite = Get-PnPSite -Includes RootWeb, HubSiteId #>

        # If SharePoint Destination
        If ($job.dstType -eq "SPO"){
            #region 9.  Configuring Hubs
            # Register site as hub site if not already
            if ($job.isHubSite.ToLower() -eq "true") {
                Write-Host ("      Configuring Hub Site: {0}" -f $destURL) -NoNewline
                $hubSite = Get-PnPHubSite -Identity $destURL
                if ($hubSite.SiteId) {
                    Write-Host (" - Already Configured") -Foregroundcolor DarkGray
                }
                else {
                    try {
                        $result = Register-PnPHubSite -Site $destURL -ErrorAction Stop
                        Write-Host (" - Configured") -Foregroundcolor DarkGray    
                    }
                    catch {
                        # Output Error Information
                        Write-CatErrorCSV -funcCSVItem $job -funcAction "Configuring Hub Site" -funcException $_.Exception -funcDetails $_.ErrorDetails -funcStackTrace $_.ScriptStackTrace -funcFilePath $errorPath
                    }                         
                }
            }
            If ($hubsite) {Clear-Variable hubSite}
            #endregion #>

            #region 10.  Associating to Hub
            # Determine if already associated to HUB
            if ($job.hubParentURL) {
                Write-Host ("      Configuring Hub Association: {0}" -f $job.dstSiteTitle) -NoNewline
                $hubSite = Get-PnPHubSite -Identity $destSite.HubSiteId.GUID -ErrorAction SilentlyContinue
                if ($hubSite.SiteUrl -eq $job.hubParentURL) {
                    Write-Host (" - Already Configured") -Foregroundcolor DarkGray
                }
                else {          
                    try {
                        Add-PnPHubSiteAssociation -Site $destURL -HubSite $job.hubParentURL -ErrorAction Stop
                        Write-Host (" - Configured") -Foregroundcolor DarkGray
                    }
                    catch {
                        # Output Error Information
                        Write-CatErrorCSV -funcCSVItem $job -funcAction "Associating Hub Site" -funcException $_.Exception -funcDetails $_.ErrorDetails -funcStackTrace $_.ScriptStackTrace -funcFilePath $errorPath
                    }
                }
            }
            If ($hubsite) {Clear-Variable hubSite}
            #endregion #>
        }

        #region 11.  Adding ShareGate Permissions
        # Adding ShareGate service account to destination site
        Write-Host ("      Adding Service Account as Destination Site Owner: {0}" -f $dstSGCreds.UserName)
        try {
            Add-PnPSiteCollectionAdmin -Owners $dstSGCreds.UserName  -ErrorAction Stop
        }
        catch {
            # Output Error Information
            Write-CatErrorCSV -funcCSVItem $job -funcAction "Adding Service Account" -funcException $_.Exception -funcDetails $_.ErrorDetails -funcStackTrace $_.ScriptStackTrace -funcFilePath $errorPath
            Continue JobLoop
        }
        #endregion #>

        #region 12. Create/Confirm Destination Library
        Write-Host ("      Confirming Library Exists: {0}" -f $job.dstLibrary) -NoNewline
        $library = Get-PnPList -Identity $job.dstLibraryPath -ErrorAction SilentlyContinue
        if ($library) {
            Write-Host (" - Already Exists") -Foregroundcolor DarkGray
        }
        else {          
            try {
                $result = New-PnPList -Title $job.dstLibrary -Template DocumentLibrary -Url $job.dstLibraryPath -OnQuickLaunch
                Write-Host (" - Created") -Foregroundcolor DarkGray
            }
            catch {
                # Output Error Information
                Write-CatErrorCSV -funcCSVItem $job -funcAction "Creating Library" -funcException $_.Exception -funcDetails $_.ErrorDetails -funcStackTrace $_.ScriptStackTrace -funcFilePath $errorPath
            }
        }
        Clear-Variable library
        #endregion

        #region 13.  Create/Confirm Destination Folder
        if ($job.dstFolder) {
            Write-Host ("      Checking if folder exists in destination: {0}" -f $dstFolder) -NoNewline
            $dstFolder = ("{0}/{1}" -f $job.dstLibraryPath, $job.dstFolder)
            try {
                $tempFolder = Get-PnPFolder -Url $dstFolder -ErrorAction Stop
                Write-Host " - Folder Already Exists" -ForegroundColor DarkGray
            }
            catch {
                Write-Host " - Folder Does not Exists" -ForegroundColor DarkGray
            }
            if (-not $tempfolder) {
                Write-Host ("      Creating Folder in Destination: {0}" -f $dstFolder)
                $folders = $job.dstFolder -split "/"
                $parentFolder = $job.dstLibraryPath
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
                        Write-CatErrorCSV -funcCSVItem $job -funcAction "Creating Folder" -funcException $_.Exception -funcDetails $_.ErrorDetails -funcStackTrace $_.ScriptStackTrace -funcFilePath $errorPath
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
            Clear-Variable dstFolder
        }
        #endregion #>

        # Disconnecting from Destination Site
        Disconnect-PnPOnline
        Clear-Variable dstSiteConnection
        Clear-Variable dstSite
        
        # Log complete
        Write-Host ("      Configuration Complete") -foregroundcolor "green"       
    }
    Else {
        # Output Error Information
        Write-CatErrorCSV -funcCSVItem $job -funcAction "Connecting to Destination Site" -funcException $_.Exception -funcDetails $_.ErrorDetails -funcStackTrace $_.ScriptStackTrace -funcFilePath $errorPath
        Continue JobLoop
    }

    #region Cleaning Up Loop Variables
    #Clear-Variable job
    Clear-Variable dstAuth
    Clear-Variable destURL
    Clear-Variable dstSiteExists
    If ($result) {Clear-Variable result}
    #endregion
}
#endregion

#region Completed Script
# Variable Clean Up
Clear-Variable jobFile
Clear-Variable dstTenant
Clear-Variable dstAppID
Clear-Variable dstCertPath
Clear-Variable timeZone
Clear-Variable siteLocale
Clear-Variable inPath
Clear-Variable errorPath
Clear-Variable inFile
Clear-Variable dstAdminUrl
Clear-Variable dstRootSiteUrl
Clear-Variable dstTenantDomain
Clear-Variable siteWaitMaxCount
Clear-Variable siteWaitPause
Clear-Variable specialCharacters
Clear-Variable dstSGCreds
Clear-Variable dstCertPass
Clear-Variable dstCert

# Output Script Info
Complete-CatScript $scriptStart[0]
Clear-Variable scriptStart

# Stop logging
Stop-Transcript
#endregion