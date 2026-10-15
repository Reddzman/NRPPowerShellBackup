#region Documentation
<#  Cat.PnP-ConfigLockState.ps1

This script script will configure the lock state on SPO / Teams / ODB related sites.  Most common use is for configuring a site read only
post migration.

Dependencies
    -	SharePointPnPPowerShellOnline Modules
    -   MicrosoftTeams Modules
	-	General Functions - ..\General\Cat.PnP-GeneralFunctionsFunctions.psm1
	-	Input CSV File - Example: ..\_Input FIles\Samples\Configuration\Cat.PnP-ConfigLockState-Test.csv

Call script with parameters
    import-module MicrosoftTeams
    cd C:\GitHub\CatPowerShell\Configuration\
	.\Cat.PnP-ConfigLockState.ps1 -jobFile "Samples\Configuration\Cat.PnP-ConfigLockState-Test-1.csv" -appID "57577043-3de0-40c0-9c68-7960e2cebe8b" -certPath "C:\GitHub\NRPDev\AppRegistration\PnP PowerShell Authentication.pfx" -clientTenant "nrpdev2" -clientTenantID "2811aa92-b6b5-458a-ab91-076e9ce5e8b3"

Table of Contents
	1.	Input & Output Files
	2.	Variables
	3.	Snapin(s), Modules & Environment Connection(s)
    4.	Main Script
    5.  Determine Settings to configure
    6.  Configuring SharePoint Lock State
    7.  Configuring Teams Archived Setting
#>
#endregion

#region Script Parameters
Param(	[string] $jobFile = "Cat.PnP-ConfigLockState-Test.csv", # Must exist in _Input Files folder
		[string] $appID = "{Azure App Registration ID}", # no {}
		[string] $certPath = "C:\AppRegistration\AzureAppRegistrationCert.pfx", # full local path
        [string] $clientTenant = "{client}", # IE. {client}.sharepoint.com  No {}
        [string] $clientTenantID = "{Client Tenant ID}" # From Azure AD  No {}
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
Start-Transcript -Path ("..\_Log FIles\{0}.PnP-ConfigLockState-{1}.log" -f $client, $scriptStart[1])
#endregion

#region 1.	Input & Output Files
# File Paths
$inPath = ("..\_Input Files\{0}" -f $jobFile)

# Import Input File(s)
# Load CSV file
If (Test-Path ($inPath)){
	$inFile = Import-Csv $inPath -Encoding Unicode
}
else {
	Write-Host
	throw ("The input file does not exist: {0}" -f $inPath)	
}
#endregion

#region 2.	Variables
# User Inputs
Write-Host
Write-Host "User Input Variables" -ForegroundColor Cyan
$certPass = ConvertTo-SecureString -String (Read-Host -Prompt "   Enter certificate password for the app registration") -AsPlainText -Force

# Script Parameters
$spoAdminURL = ("https://{0}-admin.sharepoint.com" -f $clientTenant)
$tenant = ("{0}.onmicrosoft.com" -f $clientTenant)

# Output Parameters & Variables
Write-Host
Write-Host "Script Parameters" -ForegroundColor Cyan
Write-Host ("   Input File: {0}" -f $jobFile)
Write-Host ("   App ID: {0}" -f $appID)
Write-Host ("   Certificate Path: {0}" -f $certPath)
Write-Host ("   Client Tenant Domain: {0}" -f $clientTenant)
Write-Host ("   Client Tenant ID: {0}" -f $clientTenantID)
Write-Host
Write-Host "Script Variables" -ForegroundColor Cyan
Write-Host ("   SP Admin URL: {0}" -f $spoAdminURL)
Write-Host ("   Tenant: {0}" -f $tenant)
#endregion

#region 3.	Snapin(s), Modules & Environment Connection(s)
# Load MS / PnP PowerShell Modules
Import-Module PnP.PowerShell
Import-Module MicrosoftTeams
#endregion
    
#region 4.	Main Script
# Connecting To SP Online
Write-Host
Write-Host ("Configuring Lock State / Archive: {0}" -f $jobFile) -ForegroundColor Cyan

# Loading Certificate
$cert = Get-PnPAzureCertificate -CertificatePath $certPath -CertificatePassword $certPass

:siteLoop ForEach ($site in $inFile){
    Write-Host
    Write-Host ("   Configuring {0}" -f $site.site) -ForegroundColor Magenta

    #region 5.  Determine Settings to configure
    Write-Host "      Determining Settings"
    if ($site.teamsArchived.ToLower() -eq "true") {$teamsArchived = $true} else {$teamsArchived = $false}
    switch ($site.spLockState) {
        unlock {$spLockState = "Unlock"}
        noaccess {$spLockState = "NoAccess"}
        readonly {$spLockState = "ReadOnly"}
        Default {
            Write-Host "         Incorrect Lock State.  Cancelling Job" -ForegroundColor Red
            Continue siteLoop
        }
    }
    #endregion

    #region 6.  Configuring SharePoint Lock State
    Write-Host ("      Connecting to SharePoint Online Administration: {0}" -f $spoAdminURL)
    $connectionSPAdmin = Connect-PnPOnline $spoAdminURL -tenant $tenant -ClientId $appID -Thumbprint $cert.Thumbprint -WarningAction SilentlyContinue
    #$connectionSPAdmin = Connect-PnPOnline -ClientId $appID -Url $spoAdminURL -Tenant $tenant -CertificatePath $certPath -CertificatePassword $certPass
    # Determin Site URL & Determining Teams
    Write-Host "      Retrieving URL & Determining if Teams"
    $configTeams = $false
    switch ($site.sitetype) {
        onedrive {
            # Get OneDrive URL
            $profile = Get-PnPUserProfileProperty -Account $site.site
            $siteURL = $profile.PersonalUrl.TrimEnd("/")
         }
         teams {
             $configTeams = $true
             $siteURL = $site.spURL.TrimEnd("/")
         }
         sharepoint {
            # Get SharePoint or Teams URL
            $siteURL = $site.spURL.TrimEnd("/")
        }
        Default {
            Write-Host "         Incorrect Site Type.  Cancelling Job" -ForegroundColor Red

            # Disconnecting from SPO
            Write-Host "      Disconnecting from SharePoint Online Administration"
            Disconnect-PnPOnline
            Continue siteLoop
        }
    }        

    # Retrieve Team
    Write-Host "      Retrieving SharePoint Site Collection"
    $spSite = Get-PnPTenantSite -Url $siteURL

    # Update Lock State
    if ($spSite.LockState -ne $spLockState) {
        Write-Host ("      Configuring lock state: {0}" -f $spLockState)
        Set-PnPTenantSite -Url $siteURL -LockState $spLockState -WarningAction SilentlyContinue
    }
    else {
        Write-Host ("      Lock State already configured: {0}" -f $site.site) -ForegroundColor Yellow
    }

    # Disconnecting from SPO
    Write-Host "      Disconnecting from SharePoint Online Administration"
    Disconnect-PnPOnline
    #endregion
    
    #region 7.  Configuring Teams Archived Setting
    if ($configTeams) {
        # Connecting to Teams
        Write-Host ("      Connecting to Microsoft Teams: {0}" -f $clientTenantID)
        $connectionTeams = Connect-MicrosoftTeams -ApplicationId $appID -CertificateThumbprint $cert.Thumbprint -TenantId $clientTenantID

        # Retrieve Team
        Write-Host "      Retrieving Team"
        $team = Get-Team -DisplayName $site.site

        # Set Archived State
        if ($team.Archived -ne $teamsArchived) {
            Write-Host ("      Configuring Archived State: {0}" -f $teamsArchived)
            $result = Set-TeamArchivedState -GroupId $team.GroupId -Archived $teamsArchived
        }
        else {
            Write-Host ("      Team already Archived: {0}" -f $site.site) -ForegroundColor Yellow
        }

        # Disconnecting from SPO
        Write-Host "      Disconnecting from Microsoft Teams"
        Disconnect-MicrosoftTeams
    }
    #endregion

    # Site Complete Output
    Write-Host "      Site Configuration Complete" -ForegroundColor Green
}

# Script Complete Output
Write-Host "Script Complete" -ForegroundColor Green
#endregion

#region Completed Script
# Output Script Info
Complete-CatScript $scriptStart[0]

# Stop logging
Stop-Transcript
#endregion