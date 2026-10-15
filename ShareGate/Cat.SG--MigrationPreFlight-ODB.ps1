<#  Cat.SG--MigrationPreFlight-ODB.ps1

This script automates Provisioning of user OneDrive sites and configures the destination folder in preparation for ShareGate Migration.  It is currently leveraging an App 
Registration in the destination Tenant to perform configuration.

This pre flight script has been tested with:
    -   ShareGate\Cat.SG-NetworkShareMigrations.ps1

** This should be executed in Windows PowerShell with the latest ShareGate PowerShell Modules

Dependencies
    -   Destination Azure App Registrations (Catapult Powershell & ShareGate)
    -   Windows PowerShell 
    -   ShareGate Module
    -	..\General\Cat.GeneralFunctions.psm1
    -   ..\Reporting\Cat.ReportingFunctions.psm1
    -	Input CSV File in Input Files folder

Call script with parameters
    cd C:\GitHub\CatPowerShell\ShareGate\
	.\Cat.SG--MigrationPreFlight-ODB.ps1 -jobFile "Samples\ShareGate\Cat.SG-NetworkShareMigrations-ODB-Test.csv" -destTenant "nrpdev2" -destAppID "caf6b11a-cf03-416c-b0f0-67a09a8e30ee" -destCertPath "..\..\NRPDev\AppRegistration\PowerShell App Registration.pfx"

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
Param(	[string] $jobFile = "{Teams input file path relative to _Input Files}", # Location from _Input Files folder - IE. "sample\filename.csv" or "filename.csv"
		[string] $destAppID = $null, # Leave Blank for interactive login
		[string] $destCertPath = "{Path to client certificate file must be installed on local user}",# Leave Blank for interactive login
		[string] $destTenant = "{Client tenant SharePoint prefix}" # IE. destTenant.sharepoint.com
)
#endregion Script Parameters

#region Start Date / Time
# Version Update Check
$env:PNPPOWERSHELL_UPDATECHECK="false"

# Remove Existing Catapult Modules to ensure PS session has most up to date functions
Get-Module Cat.* | Remove-Module

# Load General Functions
Import-Module "..\General\Cat.GeneralFunctions.psm1"

# Get Start Date / Time
$scriptStart = Start-CatScript

# Start Transcript
Start-Transcript -Path ("..\_Log Files\{0}.SG--MigrationPreFlight-ODB-{1}.log" -f $destTenant, $scriptStart[1])
#endregion Start Date / Time

#region 1.	Snapin(s), Modules & Environment Connection(s)
# Load MS / PnP PowerShell Modules
Import-Module ShareGate

# Load Custom Modules
Import-Module "..\Reporting\Cat.ReportingFunctions.psm1"
#endregion Snapin(s), Modules & Environment Connection(s)

#region 2.	Input & Output Files
# File Paths
$inPath = ("..\_Input Files\{0}" -f $jobFile)
$errorPath = ("..\_Output Files\{0}.SG--MigrationPreFlight-ODB-Errors-{1}.log" -f $destTenant, $scriptStart[1])

# Import Input File(s)
# Load CSV file
If (Test-Path ($inPath)){
	$inFile = Import-Csv $inPath -Encoding Unicode
}
else {
	Write-Host
	throw ("The input file does not exist: {0}" -f $inPath)	
}
#endregion Input & Output Files

#region 3.	Variables
# User Inputs
Write-Host
Write-Host "User Input Variables" -ForegroundColor Cyan
$certPass = ConvertTo-SecureString -String (Read-Host -Prompt "   Enter certificate password for the app registration") -AsPlainText -Force
$destCert = Get-PnPAzureCertificate -Path $destCertPath -Password $certPass

# Script Variables
$destAdminUrl = ("https://{0}-admin.sharepoint.com" -f $destTenant)
$destTenantDomain = ("{0}.onmicrosoft.com" -f $destTenant)
$specialCharacters = "[~``!@#$%^&*()_-]" #(($pnpGroup.Description -replace $specialCharacters,'') -replace "`n",' ') -replace "`r",' '
$siteWaitMaxCount = 20
$siteWaitPause = 20

#region Output Parameters & Variables
Write-Host
Write-Host "Script Parameters" -ForegroundColor Cyan
Write-Host ("   Input File: {0}" -f $jobFile)
Write-Host ("   App ID: {0}" -f $destAppID)
Write-Host ("   Client Tenant Prefix: {0}" -f $destTenant)
Write-Host
Write-Host "Output Variables" -ForegroundColor Cyan
Write-Host ("   Error Path: {0}" -f $errorPath)
Write-Host
Write-Host "User Inputs" -ForegroundColor Cyan
Write-Host ("   Certificate Thumbprint: {0}" -f "Only output for troubleshooting", $destCert.Thumbprint)
Write-Host
Write-Host "Script Variables" -ForegroundColor Cyan
Write-Host ("   SPO admin URL: {0}" -f $destAdminUrl)
#endregion Output Parameters & Variables
#endregion
    
#region 4.	Main Script
Write-Host
Write-Host ("Configuring & running migrations for job list: {0}" -f $jobFile) -ForegroundColor Cyan

#region 5.  Connecting to Destination Admin Tenant
Write-Host ("   {0} - Connecting to Destination Admin Tenant: {1}" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"), $destAdminUrl)
try {
    $dstAdminSite = Connect-Site -Url $destAdminUrl -Browser -DisableSSO -ErrorAction Stop
}
catch {
	Write-CatErrorCSV -funcCSVItem $job -funcAction "Connecting to SharePoint Admin (SG)" -funcException $_.Exception -funcDetails $_.ErrorDetails -funcStackTrace $_.ScriptStackTrace -funcFilePath $errorPath
	throw ("Unable to connect to SharePoint Admin")
}

#endregion

# Looping through input file
:JobLoop ForEach ($job in $inFile){
    Write-Host
    Write-Host ("   {0} - Configuring & running migration job: {1}" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"), $job.Email) -ForegroundColor Magenta
    <# Output Job Configuration for Troubleshooting (if Needed remove "<" on this line)
    $job #>

    #region 6.  OneDrive Provisioning
    Write-Host ("      {0} - Retrieving OneDrive URL: {1}" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"), $job.Email) -NoNewline
    try {
        $destOneDriveUrl = $null
        $destOneDriveUrl = Get-OneDriveUrl -Tenant $dstAdminSite -Email $job.Email -ProvisionIfRequired
		Write-Host (" - Success: {0}" -f $destOneDriveUrl) -ForegroundColor Green
    }
    catch {
		Write-Host " - Failed" -ForegroundColor Red
        Write-CatErrorCSV -funcCSVItem $job -funcAction "Connecting to SharePoint" -funcException $_.Exception -funcDetails $_.ErrorDetails -funcStackTrace $_.ScriptStackTrace -funcFilePath $errorPath
		Continue JobLoop
    }

    # Confirm OneDrive Provisioning
	Write-Host ("      {0} - Confirming OneDrive Provisioning Complete: {1}" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"), $job.Email) -NoNewline
    $counter = 0
    $destSiteExists = $false
    while (($destSiteExists -eq $false) -and $counter -lt $siteWaitMaxCount) {
        $counter++
        $destSiteExists = Test-CatPnPSiteExistsAppReg -fnUrl $destOneDriveUrl -fnAppID $destAppID -fnTenant $destTenantDomain -fnCertPath $destCertPath -fnCertPass $certPass
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
        Write-Host ("      {0} - Adding Service Account as Destination Site Owner: {1}" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"), $job.siteOwner)
		try {
			$destSiteConnection = Connect-PnPOnline -ClientId $destAppID -Url $destOneDriveUrl -Tenant $destTenantDomain -Thumbprint $destCert.Thumbprint
			Add-PnPSiteCollectionAdmin -Owners $job.siteOwner
		}
		catch {
			Write-CatErrorCSV -funcCSVItem $job -funcAction "Adding Service Account Permission" -funcException $_.Exception -funcDetails $_.ErrorDetails -funcStackTrace $_.ScriptStackTrace -funcFilePath $errorPath
			Continue JobLoop
		}
        #endregion #>

        #region 8.  Create Destination Folder
        if ($job.destFolder) {
            $destFolder = ("{0}/{1}" -f $job.destListPath, $job.destFolder)
            Write-Host ("      {0} - Checking if folder exists in destination: {1}" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"), $destFolder) -NoNewline
            try {
                $tempFolder = Get-PnPFolder -Url $destFolder -ErrorAction Stop
                Write-Host " - Folder Already Exists" -ForegroundColor Green
            }
            catch {
                Write-Host " - Folder Does not Exists" -ForegroundColor Yellow
            }
			try {
                if (-not $tempfolder) {
					Write-Host ("      {0} - Creating Folder in Destination: {1}" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"), $destFolder) -NoNewline
					$tempFolder = Add-PnPFolder -Name $job.destFolder -Folder $job.destListPath
					Write-Host " - Success" -ForegroundColor Green
				}
            }
            catch {
                Write-Host " - Failed" -ForegroundColor Red
				Write-CatErrorCSV -funcCSVItem $job -funcAction "Creating Folder" -funcException $_.Exception -funcDetails $_.ErrorDetails -funcStackTrace $_.ScriptStackTrace -funcFilePath $errorPath
				Continue JobLoop
            }
            Clear-Variable tempFolder
        }
        Disconnect-PnPOnline
        Clear-Variable destSiteConnection
        #endregion #>
	}
}

#endregion Main Script

#region Completed Script
# Output Script Info
Complete-CatScript $scriptStart[0]

# Stop logging
Stop-Transcript
#endregion Completed Script