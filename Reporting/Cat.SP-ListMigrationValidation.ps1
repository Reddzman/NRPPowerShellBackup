<#  Cat.SG-SPOnPremListMigrations.ps1

This script iterates the Job file for list level migration jobs (..\ShareGate\Cat.SG-SPOnPremListMigrations.ps1) to compare the item counts for each list object
in the source and destination enironments. The output from this script can then be processed with Power BI to identify discrepencies in the migration.

*** This script has been validated with SharePoint 2013, 2016 and needs to be validated on 2010, 2019

Dependencies
    -   Destination Azure App Registrations (Catapult Powershell)
    -   Source Service Account w/ Site Collection Admin on all sites
    -	PnP.PowerShell Module
    -	..\General\Cat.GeneralFunctions.psm1
    -	Input CSV File in Input Files folder
    -   ..\_Output Files folder

Call script with parameters 
    cd C:\Quisitive\PowerShell\Reporting
    .\Cat.SP-ListMigrationValidation.ps1 -jobFile "Cat.SG-SPOnPremListMigrations-Client" -destTenant "clientdomain" -destAppID "app guid"  -destCertPath "C:\Quisitive\PowerShell\App Reg\ClientCertFileName.pfx" -destTeamsManagedPath "sites" -mode "SPO"
    .\Cat.SP-ListMigrationValidation.ps1 -jobFile "Cat.SG-SPOnPremListMigrations-Client" -destTenant "clientdomain" -destAppID "app guid"  -destCertPath "C:\Quisitive\PowerShell\App Reg\ClientCertFileName.pfx" -destTeamsManagedPath "sites" -mode "SP on-premises"
    .\Cat.SP-ListMigrationValidation.ps1 -jobFile "Cat.SG-SPOnPremListMigrations-Client" -destTenant "clientdomain" -destAppID "app guid"  -destCertPath "C:\Quisitive\PowerShell\App Reg\ClientCertFileName.pfx" -destTeamsManagedPath "sites" -mode "both"

Table of Contents
	1.	Snapin(s), Modules & Environment Connection(s)
	2.	Input & Output Files
	3.	Variables
    4.  Functions
    5.  Write-CatListInfo
    6.  Write-CatErrorInfo
    7.	Main Script
    8.  Retreiving source SP on-premises list information
    9.  Retreiving destination SP Online list information
#>

#region Script Parameters
Param(
    [string] $jobFile = "Cat.SG-SPOnPremListMigrations-Client.csv",
    [string] $destTenant = "{Destination Tenant Name}",
    [string] $destAppID = "{Azure App Registration ID}",
    [string] $destCertPath = "C:\certs\PnPAppRegPowerShell.pfx",
    [String] $destTeamsManagedPath = "sites", # The managed path used by SharePoint site behind Microsoft Teams. IE "sites" or "teams"
    [string] $mode = "both" # SP on-premises, SPO, or Both
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
Start-Transcript -Path ("..\_Log Files\{0}.SP-ListMigrationValidation-{1}.log" -f $destTenant, $scriptStart[1])
#endregion

#region 1.	Snapin(s), Modules & Environment Connection(s)
# Load MS / PnP PowerShell Modules
Get-Module -ListAvailable | Where-Object{($_.Name -like "*SharePoint*") -or ($_.Name -like "*PnP*") } | Remove-Module
switch ($mode) {
    {(($_.ToLower() -eq "sp on-premises") -or ($_.ToLower() -eq "both"))} {Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue}
    {(($_.ToLower() -eq "spo") -or ($_.ToLower() -eq "both"))} {Import-Module PnP.PowerShell}
    Default { 
        Write-Host
	    throw ("Mode is not recognized as an option: {0}" -f $mode)	
    }
}

# Load Custom Modules
Import-Module "..\Reporting\Cat.ReportingFunctions.psm1"
#endregion

#region 2.	Input & Output Files
# Input Files
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

# Output Files
$outputBasePath = ("..\_Output Files\{0}\{1}" -f $destTenant, $scriptStart[1])
$outputErrors = ("{0}\catListMigrationValidationErrors.csv" -f $outputBasePath)
switch ($mode) {
    {(($_.ToLower() -eq "sp on-premises") -or ($_.ToLower() -eq "both"))} {$outputSPOnPrem = ("{0}\catListMigrationValidationOnPrem.csv" -f $outputBasePath)}
    {(($_.ToLower() -eq "spo") -or ($_.ToLower() -eq "both"))} {$outputSPO = ("{0}\catListMigrationValidationSPO.csv" -f $outputBasePath)}
}
#endregion

#region 3.	Variables
if (($mode.ToLower() -eq "spo") -or ($mode.ToLower() -eq "both")) {
    # User Inputs
    Write-Host
    Write-Host "Enter Connection Information" -ForegroundColor Cyan
    $destCertPass = ConvertTo-SecureString -String (Read-Host -Prompt "   Enter DESTINATION certificate password for the app registration") -AsPlainText -Force

    # Output Parameters & Variables
    # Script Variables
    $destTenantDomain = ("{0}.onmicrosoft.com" -f $destTenant)
    $destCert = Get-PnPAzureCertificate -CertificatePath $destCertPath -CertificatePassword $destCertPass
}

# Variable Output
Write-Host
Write-Host "Script Parameters" -ForegroundColor Cyan
Write-Host ("   Job File: {0}" -f $jobFile)
if (($mode.ToLower() -eq "spo") -or ($mode.ToLower() -eq "both")) {
    Write-Host ("   Destination Tenant: {0}" -f $destTenant)
    Write-Host ("   Destination App ID: {0}" -f $destAppID)
    Write-Host ("   Destination Cert Path: {0}" -f $destCertPath)
}
Write-Host ("   Destination Teams Managed Path: {0}" -f $destTeamsManagedPath)
Write-Host ("   Script Mode: {0}" -f $mode)
Write-Host
Write-Host "Script Variables" -ForegroundColor Cyan
switch ($mode) {
    {(($_.ToLower() -eq "sp on-premises") -or ($_.ToLower() -eq "both"))} {Write-Host ("   Output SP On-premises Path: {0}" -f $outputSPOnPrem)}
    {(($_.ToLower() -eq "spo") -or ($_.ToLower() -eq "both"))} {
        Write-Host ("   Output SPO Path: {0}" -f $outputSPO)
        Write-Host ("   Destination Tenant Domain: {0}" -f $destTenantDomain)
    }
}
#endregion

#region 4.  Functions

#region 5.  Write-CatListInfo

<# This function outputs the information collected from a list to a CSV file for validation of list based Migrations.

Sample Call:
    Write-CatListInfo -funJobInfo $job -funListItemCount $spOnPremList.ItemCount -funOutputPath $outputSPOnPrem #>

function Write-CatListInfo { 
	param(
        [psobject] $funJob,
        [string] $funListCount,
		[string] $funListItemCount,
        [string] $funOutputPath,
        [bool] $funErrorState,
        [string] $funErrorEnv
    )

    #$funJob # Uncomment for job information troubleshooting

    if (($funListCount -eq 1) -and (!$funErrorState)) {
        # Ouput List Information
        $csvRecord = New-Object -TypeName PSObject -Property (@{
            srcSiteTitle=$funJob.srcSiteTitle;
            srcSiteURL=$funJob.srcSiteURL;
            srcListName=$funJob.srcListName;
            srcListPath=$funJob.srcListPath;
            destSiteTitle=$funJob.destSiteTitle;
            destSiteURL=$funJob.destSiteURL;
            destGroupAlias=$funJob.destGroupAlias;
            siteTemplate=$funJob.siteTemplate;
            isHubSite=$funJob.isHubSite;
            destListName=$funJob.destListName;
            destListPath=$funJob.destListPath;
            destFolder=$funJob.destFolder;
            destChannel=$funjob.destChannel;
            destChannelType=$funjob.destChannelType;
            srcItemCount=$funListItemCount})
        $csvRecord | Export-Csv -Path $funOutputPath -NoTypeInformation -Append
        Remove-Variable csvRecord
        Write-Host " - Success" -ForegroundColor Green
    } elseif (!$errorState) {
        # Ouput if list is missing or there are duplicate titles
        Write-CatErrorInfo -funError "Missing or Duplicate List" -funJobInfo $funJob -funErrorOn $funErrorEnv -funOutputPath $outputErrors
    }
}
#endregion

#region 6.  Write-CatErrorInfo

<# This function outputs the errors that happen while collecting list information for validation of list based Migrations.

**** Need to update to leverage shared function (..\General\Functions\Cat.PS-WriteError.psm1)

Sample Call:
    Write-CatErrorInfo -funError $_ funJobInfo $job -funErrorOn "SP On-premises" -funOutputPath $outputErrors #>

function Write-CatErrorInfo { 
    param(
        [psobject] $funError,
        [psobject] $funJobInfo,
        [string] $funErrorOn,
        [string] $funOutputPath
    )

    Write-Host (" - {0}" -f $funError) -ForegroundColor Red
    Write-Host
    Write-Host ("Exception: {0}" -f $funError.Exception) -ForegroundColor Red
    Write-Host ("Error Details: {0}" -f $funError.ErrorDetails) -ForegroundColor Red
    Write-Host ("Stack Trace: {0}" -f $funError.ScriptStackTrace) -ForegroundColor Red
    $errorRecord = New-Object -TypeName PSObject -Property (@{
        srcSiteTitle=$funJobInfo.srcSiteTitle;
        srcSiteURL=$funJobInfo.srcSiteURL;
        srcListName=$funJobInfo.srcListName;
        srcListPath=$funJobInfo.srcListPath;
        destSiteTitle=$funJobInfo.destSiteTitle;
        destSiteURL=$funJobInfo.destSiteURL;
        destListName=$funJobInfo.destListName;
        destListPath=$funJobInfo.destListPath;
        errorOn=$funErrorOn;
        errorMessage=$funError})
    $errorRecord | Export-Csv -Path $funOutputPath -NoTypeInformation -Append
    Remove-Variable errorRecord
}
#endregion

#endregion
    
#region 7.	Main Script
# Created directory for additional reports
Write-Host
Write-Host ("{0} - Creating Report Directory" -f (Get-Date -Format "yy-MM-dd HH:mm:ss")) -ForegroundColor Cyan
New-Item -Path $outputBasePath -ItemType Directory

# Begin Collection List Information
Write-Host
Write-Host ("Collecting Migrated List Information: {0}" -f $jobFile) -ForegroundColor Cyan

# Looping through input file
:JobLoop ForEach ($job in $inFile){
    Write-Host
    Write-Host ("   {0} - Site: {1} List: {2}" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"), $job.srcSiteTitle, $job.srcListName) -ForegroundColor Magenta
    #$job # Uncomment to troubleshoot job information

    # Collect list information based on specified Mode
    switch ($mode) {
        #region 8.  Retreiving source SP on-premises list information
        {(($_.ToLower() -eq "sp on-premises") -or ($_.ToLower() -eq "both"))} {
            # Retreiving List Object
            write-host "      Retrieving source SharePoint on-premises list information" -NoNewline
            $errorState = $false
            try {
                $spOnPremWeb = Get-SPWeb $job.srcSiteURL -ErrorAction Stop
                $spOnPremList = $spOnPremWeb.Lists | Where-Object -Property Title -Value $job.srcListName -EQ
            }
            catch {
                # Error Handling for failures on retreiving the web/list objects
                Write-CatErrorInfo -funError $_ -funJobInfo $job -funErrorOn "SP On-premises" -funOutputPath $outputErrors

                # errorState is used to avoid logging the error twice
                $errorState = $true
            }

            # Output List information to CSV File
            Write-CatListInfo -funJob $job -funListCount $spOnPremList.Count -funListItemCount $spOnPremList.ItemCount -funOutputPath $outputSPOnPrem -funErrorState $errorState -funErrorEnv "SP On-premises"

            #Output to console
            Write-Host ("      Src Site URL: {0}" -f $job.srcSiteURL) -ForegroundColor DarkGray
            Write-Host ("      Site Template: {0}" -f $job.siteTemplate.ToLower()) -ForegroundColor DarkGray            
            Write-Host ("      Get Items in Srce List/Library: {0}" -f $job.srcListPath) -ForegroundColor DarkGray          
            Write-Host ("      Item Count: {0}" -f $spOnPremList.ItemCount) -ForegroundColor DarkGray

            # Cleaning up Variables
            If($null -ne $sponPremWeb) {Remove-Variable spOnPremWeb}
            If($null -ne $sponPremList) {Remove-Variable spOnPremList}
            Remove-Variable errorState
        }
        #endregion #>
        
        #region 9.  Retreiving destination SP Online list information
        {(($_.ToLower() -eq "spo") -or ($_.ToLower() -eq "both"))} {
            # Retreiving List Object
            write-host "      Retrieving destination SharePoint Online list information" -NoNewline
            $errorState = $false

            #region Set Destination URL
            if ($job.destSiteURL -like "https://*") {
                $destURL = $job.destSiteURL
            } else {
                $destURL = ("{0}/{1}" -f $job.classicSubSiteParentURL, $job.destSiteURL)
            }
            # Change $destURL if shared or private channel    
            switch ($job.destChannelType) {
                private {$destURL = ("{0}-{1}" -f $job.destSiteURL, ($job.destChannel -replace " ",""))}
                shared {$destURL = ("{0}{1}/{2}-{3}" -f $destRootSiteUrl, $destTeamsManagedPath, ($job.destSiteTitle -replace " ",""), ($job.destChannel -replace " ",""))}
            }
            #region
                    
            try {
                # Get List Information
                $spoConnection = Connect-PnPOnline $destURL -tenant $destTenantDomain -ClientId $destAppID -Thumbprint $destCert.Thumbprint

                # Determine which list to get the count based on siteTemplate (ex. if MSTeams get the Shared Documents library else get the list from the csv)
                if($job.siteTemplate.ToLower() -eq "msteams" -and ($job.destChannel.ToLower() -ne "none" -and $job.destChannel -ne $null)) {
                    $spoList = Get-PnPList -Identity "Documents" -ErrorAction Stop
                } else {
                    $spoList = Get-PnPList -Identity $job.destListName -ErrorAction Stop
                }                

                if ($job.destFolder) {
                    # Determine which folder in Documents to get the count from if siteTemplate is MSTeams
                    If($job.siteTemplate.ToLower() -eq "msteams" -and ($job.destChannel.ToLower() -ne "none" -and $job.destChannel -ne $null)) {
                        # Use Folder information in Channel's Documents library
                        $spoChannelFolderItems = Get-PnPFolderItem -FolderSiteRelativeUrl ("{0}/{1}/{2}" -f $spoList.RootFolder.ServerRelativeUrl.Replace(("{0}/" -f $spoList.ParentWebUrl),""), $job.destChannel,$job.destFolder) -Recursive
                        $spoItemCount = $spoChannelFolderItems.Count
                        Remove-Variable spoChannelFolderItems
                    } Else {
                        # Get Folder Information
                        $spoFolderItems = Get-PnPFolderItem -FolderSiteRelativeUrl ("{0}/{1}" -f $spoList.RootFolder.ServerRelativeUrl.Replace(("{0}/" -f $spoList.ParentWebUrl),""), $job.destFolder) -Recursive
                        $spoItemCount = $spoFolderItems.Count
                        Remove-Variable spoFolderItems
                    }                    
                } Else {
                    If($job.siteTemplate.ToLower() -eq "msteams") {
                        # Use Folder information in Channel's Documents library
                        $spoChannelItems = Get-PnPFolderItem -FolderSiteRelativeUrl ("{0}/{1}" -f $spoList.RootFolder.ServerRelativeUrl.Replace(("{0}/" -f $spoList.ParentWebUrl),""), $job.destChannel) -Recursive
                        $spoItemCount = $spoChannelItems.Count
                        Remove-Variable spoChannelItems
                    } Else {
                        # Use List information
                        $spoItemCount = $spoList.ItemCount   
                    }                
                }
                Disconnect-PnPOnline
            }
            catch {
                # Error Handling for failures on retreiving the web/list objects
                Write-CatErrorInfo -funError $_ -funJobInfo $job -funErrorOn "SP Online" -funOutputPath $outputErrors

                # errorState is used to avoid logging the error twice
                $errorState = $true
            }

            # Output List information to CSV File
            Write-CatListInfo -funJob $job -funListCount $spoList.Count -funListItemCount $spoItemCount -funOutputPath $outputSPO -funErrorState $errorState -funErrorEnv "SP Online"
            
            #Output to console
            Write-Host ("      Dest Site URL: {0}" -f $destURL) -ForegroundColor DarkGray
            Write-Host ("      Site Template: {0}" -f $job.siteTemplate.ToLower()) -ForegroundColor DarkGray
            If($job.siteTemplate.ToLower() -eq "msteams" -and ($job.destChannel.ToLower() -ne "none" -and $job.destChannel -ne $null))  {
                Write-Host ("      Channel Name: {0}" -f $job.destChannel) -ForegroundColor DarkGray
            }
            If($job.destFolder) {
                Write-Host ("      Get Items in Dest Folder: {0}" -f $job.destFolder) -ForegroundColor DarkGray
            } else {
                Write-Host ("      Get Items in List/Library Root: {0}" -f $job.destListPath) -ForegroundColor DarkGray
            }            
            Write-Host ("      Item Count: {0}" -f $spoItemCount) -ForegroundColor DarkGray

            # Cleaning up Variables
            If($null -ne $spoConnection) {Remove-Variable spoConnection}
            If($null -ne $spoList) {Remove-Variable spoList}
            #If($job.destFolder) {Remove-Variable spoFolderItems}
            Remove-Variable spoItemCount
            Remove-Variable errorState
        }
        #endregion #>
    }
}
#endregion

#region Completed Script
# Output Script Info
Complete-CatScript $scriptStart[0]

# Stop logging
Stop-Transcript
#endregion