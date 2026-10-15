#region Documentation
<#  Cat.PnP-RemoveCustomListPermissions.ps1

This script was created as part of a migration project.  The client accidentally migrated custom permissions for each list and library. To fix 
the issue this script iterates through the Cat.SG-SPOnPremListMigrations.ps1 (ShareGate Migration CSV used for client) scripts CSV File 
(Cat.SG-SPOnPremListMigrations-Client.csv) and removes custom permissions on the list. It also iterates through all list items to ensure there 
are no custom permissions.  This is a very long running script.

Dependencies
    -   Client Azure App Registration
    -   PnP.PowerShell Module
    -	..\General\Cat.GeneralFunctions.psm1
    -   ..\Reporting\Cat.ReportingFunctions.psm1
    -	Input CSV File in Input Files folder

Call script with parameters
    cd C:\GitHub\CatPowerShell\Configuration\
	.\Cat.PnP-RemoveCustomListPermissions.ps1 -jobFile "Samples\Configuration\Cat.PnP-RemoveCustomListPermissions-Client.csv" -clientTenant nrpdev2 -clientAppID "caf6b11a-cf03-416c-b0f0-67a09a8e30ee" -clientCertPath "C:\GitHub\NRPDev\AppRegistration\PowerShell App Registration.pfx"

CSV Notes
    Settings columns are based on options for New-CopySettings & Copy-List ShareGate Documentation
        https://support-desktop.sharegate.com/hc/en-us/articles/115000597407-New-Copy-Settings
        https://support-desktop.sharegate.com/hc/en-us/articles/115000597727-Copy-List

Table of Contents
	1.	Input & Output Files
	2.	Variables
	3.	Snapin(s), Modules & Environment Connection(s)
    4.	Main Script
    5.  Confirmation to Proceed
    6.  Connecting to Site & Retrieving List Information
    7.  Reseting List Permissions
    8.  Resetting Item Permissions
#>
#endregion

#region Script Parameters
Param(
    [String] $jobFile = "Cat.SG-SPOnPremListMigrations-Client.csv",
    [string] $clientTenant = "{Client Tenant Name}",
    [String] $clientAppID = "{Azure App Registration ID}",
    [String] $clientCertPath = "C:\certs\PnPAppRegPowerShell.pfx"
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
Start-Transcript -Path ("..\_Log Files\{0}.PnP-RemoveCustomListPermissions-{1}.log" -f $clientTenant, $scriptStart[1])
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
	throw ("The jobs file does not exist: {0}" -f $inPath)	
}
#endregion

#region 2.	Variables
# User Inputs
Write-Host
Write-Host "User Inputs" -ForegroundColor Cyan
$clientCertPass = ConvertTo-SecureString -String (Read-Host -Prompt "   Enter Client certificate password for the app registration") -AsPlainText -Force

# Output Parameters & Variables
# Script Variables
$destAdminUrl = ("https://{0}-admin.sharepoint.com/" -f $clientTenant) #TODO confirm this is needed
$destRootSiteUrl = ("https://{0}.sharepoint.com/" -f $clientTenant)
$clientTenantDomain = ("{0}.onmicrosoft.com" -f $clientTenant)
$clientCert = Get-PnPAzureCertificate -CertificatePath $clientCertPath -CertificatePassword $clientCertPass
$itemPageSize = 5000

# Variable Output
Write-Host
Write-Host "Script Parameters" -ForegroundColor Cyan
Write-Host ("   Job File: {0}" -f $jobFile)
Write-Host ("   Client Tenant: {0}" -f $clientTenant)
Write-Host ("   Client App ID: {0}" -f $clientAppID)
Write-Host ("   Client Cert Path: {0}" -f $clientCertPath)
Write-Host
Write-Host "Script Variables" -ForegroundColor Cyan
Write-Host ("   Client Admin Url: {0}" -f $destAdminUrl)
Write-Host ("   Client Root Site Url: {0}" -f $destRootSiteUrl)
Write-Host ("   Client Tenant Url: {0}" -f $clientTenantDomain)
Write-Host ("   Item Page Size: {0}" -f $itemPageSize)
#endregion

#region 3.	Snapin(s), Modules & Environment Connection(s)
# Load MS / PnP PowerShell Modules
Get-Module -ListAvailable | Where-Object{($_.Name -like "*SharePoint*") -or ($_.Name -like "*PnP*") } | Remove-Module
Import-Module PnP.PowerShell

# Load Custom Modules
Import-Module "..\Reporting\Cat.ReportingFunctions.psm1"
#endregion
    
#region 4.	Main Script
#region 5.  Confirmation to Proceed
Write-Host
Write-Host "*** Deleting existing custom permissions will permanantly remove permissions and inherit the partent's permissions for all lists in this CSV. ***" -ForegroundColor Yellow
$userProceed = Read-Host "   Are you sure you want to proceed with deleting custom pemrissions (Yes or No)" 

# if user chose to proceed
if ($userProceed.ToLower() -eq "yes") {
    $deleteCustomPermissions = $true
}
else {
    $deleteCustomPermissions = $false
}
#endregion

if ($deleteCustomPermissions) {
    Write-Host
    Write-Host ("Removing Custom List Permissions: {0}" -f $jobFile) -ForegroundColor Magenta

    # Looping through input file
    :JobLoop ForEach ($job in $inFile){
        Write-Host
        Write-Host ("   {0} - Updating Permissions - Site: {1} List: {2}" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"), $job.destSiteTitle, $job.destListName)

        #region 6.  Connecting to Site & Retrieving List Information
        Write-host "      Connecting to SharePoint Site" -ForegroundColor DarkGray
        $spSiteConnection = Connect-PnPOnline -ClientId $clientAppID -Url $job.destSiteURL -Tenant $clientTenantDomain -Thumbprint $clientCert.Thumbprint

        # Retrieving List object
        Write-host "      Retrieving List" -ForegroundColor DarkGray
        $spList = Get-PnPList -Identity $job.destListName  -Includes HasUniqueRoleAssignments
        #endregion
        
        If ($spList) {
            #region 7.  Reseting List Permissions
            # Setting removing custom permissions from the list object
            Write-host "      Processing List" -ForegroundColor DarkGray -NoNewline
            If ($spList.HasUniqueRoleAssignments) {
                Set-PnPList -Identity $spList.Id -ResetRoleInheritance:$true -Connection $spSiteConnection
                Write-host " - Custom Permissions Removed" -ForegroundColor Green
            } else {
                Write-host " - No Custom Permissions" -ForegroundColor DarkGray
            }
            #endregion

            #region 8.  Resetting Item Permissions
            # Retrieving all list items
            Write-host "      Retrieving & Processing List items" -ForegroundColor DarkGray
            $spListItems = Get-PnPListItem -List $spList.Id -PageSize $itemPageSize

            # looping through all list items
            foreach ($item in $spListItems) {
                # Determine if Item has custom permissions
                Write-host ("         {0} - Processing Item: {1}" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"), $item.Id) -ForegroundColor Blue -NoNewline
                [bool] $itemHasUniqueRoleAssignments = Get-PnPProperty -ClientObject $item -Property HasUniqueRoleAssignments

                # If permissions is broken, remove custom permisisons
                If ($itemHasUniqueRoleAssignments) {
                    # removing custom list permissions
                    Set-PnPListItemPermission -List $spList.Id -Identity $item.Id -InheritPermissions:$true
                    Write-Host " - Custom Permissions Removed" -ForegroundColor Green
                } else {
                    Write-Host " - No Custom Permissions" -ForegroundColor Blue
                }
                Remove-Variable itemHasUniqueRoleAssignments
            }
            #endregion
        } else {
            Write-host "      No List Found" -ForegroundColor Yellow
        }

        #region Disconnecting and removing variables
        Disconnect-PnPOnline
        Remove-Variable spSiteConnection
        If ($spList) {
            Remove-Variable spList
            Remove-Variable spListItems
        }
        #endregion
    }
} else {
    Write-Host
    Write-Host "No Custom Permissions have been removed:" -ForegroundColor Yellow -NoNewline
    Write-Host (" {0}" -f $jobFile)
}
#endregion

#region Completed Script
# Output Script Info
Complete-CatScript $scriptStart[0]

# Stop logging
Stop-Transcript
#endregion