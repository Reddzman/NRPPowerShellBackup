<#	Cat.PnP-CreationMigrationSites.ps1

For migration purposes, this script iterates through a CSV file and creates a site collections in SharePoint 
Online based on the exported template from the on-premises/online site. It activates features based on the on-premises site. 

Dependencies
	-	PnP PowerShell https://msdn.microsoft.com/en-us/pnp_powershell/pnp-powershell-overview
	-	CSV input file in _Input Files folder

Sample
	Cat.PnP-CreateMigrationSites -inputFile cat.CM-SiteList-Test.csv -adminURL https://nrpdev-admin.sharepoint.com -ownerPrimary nickp@nrpdev.com -ownerSecondary sapadmin@nrpdev.com -timezone 12

Table of Contents
	1.	Variables
	2.	Input & Output Files
	3.	Snapin(s), Modules & Environment Connection(s)
	4.	Main Script
	5.	Create New Site Collection
	6.	Get Provisioning Template
	7.	Applying Provisioning Template
#>

#region Script Parameters
Param(	[string] $inputFile = "cat.CM-SiteList-Test.csv",
		[string] $adminURL = "https://nrpdev-admin.sharepoint.com"
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
Start-Transcript -Path ("..\_Log FIles\CreateMigraitonSites-{0}.log" -f $scriptStart[1])
#endregion

#region 1.	Variables
# User Input Variables
Write-Host
Write-Host "User Input Variables" -ForegroundColor Cyan
Write-Host "   Enter credentials for source SharePoint Environment:"
$srcCred = Get-Credential -Message "Enter credentials for source SharePoint Environment:" 
Write-Host "   Enter credentials for destination SharePoint environment:"
$dstCred = Get-Credential -Message "Enter credentials for destination SharePoint environment:"

# Script Variables
$pnpHandlers = "Features" # used to determine what to export from on-premises site
$modernTemplates = @("GROUP#0","SITEPAGEPUBLISHING#0")
$modernCommTemplate = "SITEPAGEPUBLISHING#0"

# Output Parameters & Variables
Write-Host
Write-Host "Script Parameters" -ForegroundColor Cyan
Write-Host ("   Input File: {0}" -f $inputFile)
Write-Host ("   Admin URL: {0}" -f $adminURL)
Write-Host
Write-Host "User Inputs" -ForegroundColor Cyan
Write-Host ("   Source Credentials: {0}" -f $srcCred.UserName)
Write-Host ("   Destination Credentials: {0}" -f $dstCred.UserName)
Write-Host
Write-Host "Script Variables" -ForegroundColor Cyan
Write-Host ("   PnP Provisioning Handlers: {0}" -f $pnpHandlers)
#endregion

#region 2.	Input & Output Files
# File Paths
$inPath = ("..\_Input Files\{0}" -f $inputFile)

# Import Input File(s)
If (Test-Path ($inPath)){
	$inFile = Import-Csv $inPath -Encoding Unicode
}
else {
	Write-Host
	throw ("The jobs file does not exist: {0}" -f $inPath)	
}
#endregion

#region 3.	Snapin(s), Modules & Environment Connection(s)
# Load MS / PnP PowerShell Modules
Import-Module SharePointPnPPowerShellOnline -WarningAction SilentlyContinue

# Load Custom Modules
Import-Module "..\Reporting\Cat.ReportingFunctions.psm1"
#endregion
    
#region 4.	Main Script
Write-Host
Write-Host "Creating Site Collections" -ForegroundColor Cyan
ForEach ($site in $inFile){
	# Determine if site exists
    $siteExists = Test-CatPnPSiteExists -url $site.dstURL -cred $dstCred

    # Create Site
    If ($siteExists){
        Write-Host ("   Site Already Exists: {0}" -f $site.dstURL)
    }
    Else {
		Write-Host
		Write-Host ("   Creating destination site: {0}" -f $site.dstURL) -ForegroundColor Magenta
		
		#region 5.	Creating New Site
		Write-Host ("      Creating site: {0} ." -f $site.dstURL) -NoNewline
		$pnpAdminConnect = Connect-PnPOnline -Url $adminURL -Credentials $dstCred
		
		# Check for Modern Templates
		If ($modernTemplates -contains $site.template){
			# Determine Site Type
			If ($site.template -eq $modernCommTemplate){
				# Create Modern Communication Site
				$tempSite = New-PnPSite -Title $site.title -Url $site.dstURL -Type "CommunicationSite" -Connection $pnpAdminConnect
			}
			Else{
				# Determine Alias
				$siteAlias = $site.dstURL -split "/"
				$siteAlias = $siteAlias[($siteAlias.count -1)]

				# Create Modern Team Site
				$tempSite = New-PnPSite -Title $site.title -Alias $siteAlias -Type "TeamSite" -Connection $pnpAdminConnect
			}
		}
		Else{
			# Create Classic Site
			$tempSite = New-PnPTenantSite -Title $site.title -Url $site.dstURL -Owner $site.ownerPrimary -TimeZone $site.timezone -Template $site.template -Wait
		}
		Write-Host
		Disconnect-PnPOnline -Connection $pnpAdminConnect
		#endregion

		#region 6.	Get Provisioning Template
		Write-Host ("      Extracting provisioning template from source site: {0}" -f $site.srcURL)
		$pnpSrcConnect = Connect-PnPOnline -Url $site.srcURL -Credentials $srcCred
		$pnpTemplate = Get-PnPProvisioningTemplate -OutputInstance -Handlers  $pnpHandlers -Connection $pnpSrcConnect
		Disconnect-PnPOnline -Connection $pnpSrcConnect
		#endregion

		#region 7.	Applying Provisioning Template
		Write-Host "      Applying Provisioning Template to destination site collection"
		$pnpDstConnect = Connect-PnPOnline -Url $site.dstURL -Credentials $dstCred
		Apply-PnPProvisioningTemplate -InputInstance $pnpTemplate
		Disconnect-PnPOnline -Connection $pnpDstConnect
		#endregion

		# Screen Output for completion
		Write-Host ("      Successfully Created: {0}" -f $site.dstURL) -ForegroundColor Green
    }
}
#endregion

#region Completed Script
# Output Script Info
Complete-CatScript $scriptStart[0]

# Stop logging
Stop-Transcript
#endregion