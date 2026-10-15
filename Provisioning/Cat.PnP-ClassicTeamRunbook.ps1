<# Cat.PnP-ClassicTeamRunbook.ps1

This script is leveraged as an Azure runbook to provision a classic team site.

Dependencies
    -	PnP PowerShell https://docs.microsoft.com/en-us/powershell/sharepoint/sharepoint-pnp/sharepoint-pnp-cmdlets?view=sharepoint-ps
    -   Azure App Permissions https://github.com/SharePoint/PnP-PowerShell/tree/master/Samples/SharePoint.ConnectUsingAppPermissions
    -   Self Signed Cert
    -   Azure Storage / file share

Call Script with Input File
    .\Cat.PnP-ClassicTeamRunbook.ps1 -adminURL "https://spnickp-admin.sharepoint.com" -tenantURL "https://spnickp.sharepoint.com" -etc.

Table of Contents
	1.	Variables
	2.	Main Script
#>

<#region ***  Local Only Script Parameters ***
Param(  [String] $adminURL = "https://spnickp-admin.sharepoint.com",
        [String] $tenantURL = "https://spnickp.sharepoint.com", 
        [String] $tenantDomain = "spnickp.onmicrosoft.com", 
        [String] $certFile = "siteprovisioning-teamsite.pfx",
        [String] $certStgAcct = "siteprovisioningacct",
        [String] $certStgKey = "fc1OOzMKMQDyRMpYXYfL26A8Q5TJd2K4/uOico9DjYpPbBwbByn56DSPru3Ev9DlYcNqZWcCLVCOuK+iMgq5XA==",
        [String] $certShare = "siteprovisioning",
        [String] $managedPath = "/sites/",
        [String] $templateLibrary = "https://spnickp.sharepoint.com/sites/SPOAdmin/Sitetemplates",
        [String] $templateName = "ClassicTeamSite.pnp",
        [String] $siteName = "Test Site 01",
        [String] $sitePath = "siteprovisioning-test01",
        [String] $siteOwner = "nickp@nrpsolutions.com", 
        [String] $scAdmins = "nickp@nrpsolutions.com",
        [String] $siteOwners = "nicksman@nrpsolutions.com",
        [String] $siteMembers = "nickp@nrpsolutions.com",
        [String] $siteVisitors = "nickp@nrpsolutions.com;nicksman@nrpsolutions.com",
        [String] $requestSite = "https://spnickp.sharepoint.com/sites/SPOAdmin",
        [String] $requestList = "Site Collections",
        [String] $requestID= "1",
        [String] $storageQuota = "1024",
        [String] $resourceQuota = "100",
        [String] $locale = "1033",
        [String] $timezoneNum = "4"
)
#endregion *** End of Local Only Script Parameters #>

#region ***  Runbook Only Script Parameters ***
Param([object] $WebhookData)
#endregion *** End of Runbook Only Script Parameters #>

#region Start Date / Time
<#  ***  Use this section if running locally ***

# Set PS locPath
Set-Location $PSScriptRoot

# Get Script Parent locPath
$locSplit = $PSScriptRoot.Split("\")
$locReplace = $locSplit[($locSplit.Count - 1)]
$locPath = $PSScriptRoot.Replace($locReplace,"")

# Remove Existing Catapult Modules to ensure PS session has most up to date functions
Get-Module Cat.* | Remove-Module

# Load General Functions
Import-Module ("{0}General\Cat.GeneralFunctions.psm1" -f $locPath)

***  End of section for local running *** #>

# Get Start Date / Time
$startTime = Start-CatScript
#endregion

#region 1.	Variables
<# ***  Local Only Variables ***
# User Inputs
<#Write-Output
Write-Output "User Input Variables" -ForegroundColor Cyan
Write-Output "   Enter credentials for SharePoint:"
$spCred = Get-Credential -Message "Enter your credentials for SharePoint:"

# Certificate Variables
#$certPath = ("..\_Input Files\Certs\{0}" -f $certName)
# ***  End of Local Variables *** #>

# *** Runbook Only Variables
# Webhook Data Variables
$variables = (ConvertFrom-Json -InputObject $WebhookData.RequestBody)
$adminURL = $Variables.ADMINURL
$tenantURL = $Variables.TENANTURL
$tenantDomain = $Variables.TENANTDOMAIN
$certFile = $Variables.CERTFILE
$certStgAcct = $Variables.CERTSTGACCT
$certStgKey = $Variables.CERTSTGKEY
$certShare = $Variables.CERTSHARE
$managedPath = $Variables.MANAGEDPATH
$templateLibrary = $Variables.TEMPLATELIBRARY
$templateName = $Variables.TEMPLATENAME
$siteName = $Variables.SITENAME
$sitePath = $Variables.SITEPATH
$siteOwner = $variables.SITEOWNER
$scAdmins = $Variables.SCADMINS
$siteOwners = $Variables.SITEOWNERS
$siteMembers = $Variables.SITEMEMBERS
$siteVisitors = $Variables.SITEVISITORS
$requestSite = $Variables.REQUESTSITE
$requestList = $Variables.REQUESTLIST
$requestID= $Variables.REQUESTID
$storageQuota = $Variables.STORAGEQUOTA
$resourceQuota = $Variables.RESOURCEQUOTA
$locale = $Variables.LOCALE
$timezoneNum = $Variables.TIMEZONENUM

# Certificate Variables
$certPath = ("{0}\{1}" -f $env:TEMP,$certFile)
# *** End of Runbook Only Variables *** #>

# Script Variables
$compStatus = "Completed"
$appID = "c4c1a7b2-757f-456e-a8a2-6690dd1bb1cb"
$siteURL = ("{0}{1}{2}" -f $tenantURL,$managedPath,$sitePath)   

# Permissions Groups
$ownersGroup = $siteName + ' Owners'
$membersGroup = $siteName + ' Members'
$visitorsGroup = $siteName + ' Visitors'

# Standard Owners
$defaultOwners = @('nickp@nrpsolutions.com')

# Folder variables (Term Store)
$folderLibrary = "Documents"
$folderTermGroup = "290dcdc2-7ddf-4839-b477-c3e382182dae" # GUID For Term Group in SPO Managed Metadata
$folderTermSet = "e02cf659-a44f-4a3a-8294-e2f4e9e854ba" # GUID For Term Set in SPO Managed Metadata

# Navigation Clean Up
$removeTopNav = @("Template - Classic Team Site")
$removeQuickLaunch = @("Home";"Notebook";"Calendar";"Contacts";"Documents";"Tasks";"Site Contents")
#endregion

#region Output Parameters & Variables
Write-Output "Script Parameters"
Write-Output ("   Admin URL: {0}" -f $adminURL)
Write-Output ("   Tenant URL: {0}" -f $tenantURL)
Write-Output ("   Tenant Domain: {0}" -f $tenantDomain)
Write-Output ("   Certificate Name: {0}" -f $certFile)
Write-Output ("   Storage Account: {0}" -f $certStgAcct)
Write-Output ("   Storage Key: {0}" -f $certStgKey)
Write-Output ("   Share Name: {0}" -f $certShare)
Write-Output ("   Managed Path: {0}" -f $managedPath)
Write-Output ("   Template Library: {0}" -f $templateLibrary)
Write-Output ("   Template Name: {0}" -f $templateName)
Write-Output ("   Site Name: {0}" -f $siteName)
Write-Output ("   Site Path: {0}" -f $sitePath)
Write-Output ("   Template Name: {0}" -f $templateName)
Write-Output ("   Site URL: {0}" -f $siteURL)
Write-Output ("   Site Collection Admins: {0}" -f $scAdmins)     
Write-Output ("   Site Owners: {0}" -f $siteOwners)
Write-Output ("   Site Members: {0}" -f $siteMembers)
Write-Output ("   Site Visitors: {0}" -f $siteVisitors)
Write-Output ("   Request Site: {0}" -f $requestSite)
Write-Output ("   Request List: {0}" -f $sitePath)
Write-Output ("   Request ID: {0}" -f $requestID)
Write-Output ("   Storage Quota: {0}" -f $storageQuota)
Write-Output ("   Resource Quota: {0}" -f $resourceQuota)
Write-Output ("   Locale: {0}" -f $locale)
Write-Output ("   Timezone Number: {0}" -f $timezoneNum)
Write-Output "Script Variables"
Write-Output ("   Application ID: {0}" -f $appID)
Write-Output ("   Completion Status: {0}" -f $compStatus)
Write-Output ("   Owners Group: {0}" -f $ownersGroup)
Write-Output ("   Members Group: {0}" -f $membersGroup)
Write-Output ("   Visitors Group: {0}" -f $visitorsGroup)
Write-Output ("   Default Owners: {0}" -f $defaultOwners)
Write-Output ("   Folders Library: {0}" -f $folderLibrary)
Write-Output ("   Folders Term Group: {0}" -f $folderTermGroup)
Write-Output ("   Folders Term Set: {0}" -f $folderTermSet)
#endregion #>

#region 2.	Main Script
try{
    # *** Runbook Only Logic ***
    # Getting Certificate for connection
    Write-Output ("Retrieving Certificate from Azure: {0}" -f $certFile)
    $ctx = New-AzureStorageContext $certStgAcct $certStgKey
    $share = Get-AzureStorageShare $certShare –Context $ctx
    Get-AzureStorageFileContent –Share $share –Path $certFile -Destination $env:TEMP -Force
    # *** End of Runbook Only Logic *** #>

    # Get Certificate information
    Write-Output ("Loading PnP Certificate: {0}" -f $certPath)
    $appCert = Get-PnPAzureCertificate -CertificatePath $certPath

    # Connecting to SP Admin
    Write-Output ("Connecting to SP Admin Site: {0}" -f $adminURL)
    Connect-PnPOnline -Tenant $tenantDomain -Url $adminURL -PEMCertificate $appCert.Certificate -PEMPrivateKey $appCert.PrivateKey -ClientId $appID

    # Create site collection
    Write-Output ("   Creating site collection: {0} " -f $siteName)
    New-PnPTenantSite -Title $siteName -Url $siteURL -Owner $siteOwner -TimeZone $timezoneNum -ResourceQuota $resourceQuota -StorageQuota $storageQuota -Lcid $locale -Wait -Force -ErrorAction Stop

    # Waiting on Site Creation
    Write-Output ("Connecting to new site: {0}" -f $siteURL)
    Connect-PnPOnline -Tenant $tenantDomain -Url $siteURL -PEMCertificate $appCert.Certificate -PEMPrivateKey $appCert.PrivateKey -ClientId $appID

    # Apply site template
    Write-Output ("   Applying Template: {0}" -f $templateName)
    Apply-PnPProvisioningtemplate -Path "$templateLibrary/$templateName" -ErrorAction Stop

    # Add Site collection admins
    Write-Output ("   Adding Site Collection Admins: {0}" -f $scAdmins)
    If (-not ([string]::IsNullOrEmpty($scAdmins))) {
        Add-PnPSiteCollectionAdmin -Owners $scAdmins
    }

    # Add default owners
    Write-Output ("   Adding Default Site Owners: {0}" -f $defaultOwners)
    If (-not ([string]::IsNullOrEmpty($siteOwners))) {
        $arrDefaultOwners = $defaultOwners -split ";"
        foreach ($user in $arrDefaultOwners) {
            Add-PnPUserToGroup -LoginName $user -Identity $ownersGroup
        }
    }

    # Add TSP site specific owners
    Write-Output ("   Adding Specific Site Owners: {0}" -f $siteOwners)    
    If (-not ([string]::IsNullOrEmpty($siteOwners))) {
        $arrSiteOwners = $siteOwners -split ";"
        foreach ($user in $arrSiteOwners) {
            Add-PnPUserToGroup -LoginName $user -Identity $ownersGroup
        }
    }

    # Add TSP site specific members
    Write-Output ("   Adding Specific Site Members: {0}" -f $siteMembers)
    If (-not ([string]::IsNullOrEmpty($siteMembers))) {
        $arrSiteMembers = $siteMembers -split ";"
        foreach ($user in $arrSiteMembers) {
            Add-PnPUserToGroup -LoginName $user -Identity $membersGroup
        }
    }

    # Add TSP site specific visitors
    Write-Output ("   Adding Specific Site Visitors: {0}" -f $siteVisitors)
    If (-not ([string]::IsNullOrEmpty($siteVisitors))) {
        $arrSiteVisitors = $siteVisitors -split ";"
        foreach ($user in $arrSiteVisitors) {
            Add-PnPUserToGroup -LoginName $user -Identity $VisitorsGroup
        }
    }

    # Get Document Library for Folders
    $library = Get-PnPList -Identity $folderLibrary

    # Get child managed metadata terms for folders
    $parentTerms = Get-PnPTerm -TermGroup $folderTermGroup -TermSet $folderTermSet -IncludeChildTerms -Recursive

    # Iterate through Parent Terms and creating folders
    Write-Output "   Creating Folders"
    ForEach ($pTerm in $parentTerms){
        $folderPath = ("{0}/{1}" -f $library.RootFolder.Name,$pTerm.Name)
        $tempFolder = Resolve-PnPFolder -SiteRelativePath $folderPath
        Write-Output ("      Folder: {0}" -f $folderPath)

        # Iterate through Child Terms and creating folders
        if ($pTerm.TermsCount -gt 0){
            ForEach ($cTerm in $pTerm.Terms){
                $folderPath = ("{0}/{1}/{2}" -f $library.RootFolder.Name,$pTerm.Name,$cTerm.Name)
                $tempFolder = Resolve-PnPFolder -SiteRelativePath $folderPath
                Write-Output ("      Folder: {0}" -f $folderPath)
            }
        }
    }

    # Clean Up Global Navigation
    Write-Output "   Configuring Global Navigation"
    ForEach ($Node in $removeTopNav){
        Get-PnPNavigationNode -Location TopNavigationBar | Where{$_.Title -eq $Node} | Remove-PnPNavigationNode -Force
    }

    # Clean Up Quick Launch Navigation
    Write-Output "   Configuring Quick Launch"
    $arrNodes = @()
    ForEach ($Node in $removeQuickLaunch){
        $arrNodes = Get-PnPNavigationNode -Location QuickLaunch | Where{$_.Title -eq $Node}
        Remove-PnPNavigationNode -Identity $arrNodes[1] -Force
    }
    Get-PnPNavigationNode| Where{($_.Title -eq "Recent") -or ($_.Title -eq "Pages")} | Remove-PnPNavigationNode -Force
}
catch{
	Write-Output "!! EXCEPTION OCCURRED !!"
	Write-Output "!! Exception Type: $($_.Exception.GetType().FullName)"
    Write-Output "!! Exception Message: $($_.Exception.Message)"
    $compStatus = "!! Exception Occurred !!"
}

# Updating Site Provisioning List Item
try {
    Write-Output ("Connecting to Request site: {0}" -f $requestSite)
    Connect-PnPOnline -Tenant $tenantDomain -Url $requestSite -PEMCertificate $appCert.Certificate -PEMPrivateKey $appCert.PrivateKey -ClientId $appID

    # Updating List Item
    Write-Output ("   Updating Request: {0}" -f $requestID)
    $updateValues = @{"Status"=$compStatus;"scurl"=$siteURL}
    $tempItem = Set-PnPListItem -List $requestList -Identity $requestID -Values $updateValues
}
catch {
	Write-Output "!! EXCEPTION OCCURRED !!"
	Write-Output "!! Exception Type: $($_.Exception.GetType().FullName)"
    Write-Output "!! Exception Message: $($_.Exception.Message)"
}

# Output Completion
Write-Output "TSP Site Provisioning Completed"
#endregion

#region Completed Screen Output
Disconnect-PnPOnline
Complete-CatScript $startTime[0]
#endregion