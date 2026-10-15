<#  Cat.PnP-SiteAccess.ps1

This script can take a CSV file from the SharePoint Admin Center's active sites export or scan all sites in a tenant if 
a CSV file is not provided. For each site scanned, the script will
	-	Determine if the site is group or teams connected and output members and owners
	-	Iterate the SharePoint permissions and output user/groups and the associated permissions
	-	Iterate all lists and libraries to 
		*	Output permissions on list and libraries with broken inheritance
		*	Query for items with unique permissions and output the permissions of the individual files, including share links

The output is written to a SharePoint list and will contain repeat records, in the event that a user is a member of the group, 
listed specifically in site permission or in lists or items with unique permissions, that user will be listed several times in 
the resulting report.

Dependencies
	-	PnP.PowerShell Modules
	-	SPO PowerShell Modules
	-	Input CSV File - Example: ..\_Input FIles\Testing\Cat.Tech-Action-Test.csv
	-	Requires App Registration, Interactive login will prompt user for credentials multiple times per site

Call script with parameters
	cd c:\GitHub\CatPowerShell\Reporting
	.\Cat.PnP-SiteAccess.ps1 -clientTenant "nrpdev2" -reportingSiteURL "https://nrpdev2.sharepoint.com/sites/M365Admins" -siteAccessList "Site Access" -listCleanup $true -inputFile "Samples\Reporting\Cat.PnP-SiteAccess-Test.csv" -clientAppID "caf6b11a-cf03-416c-b0f0-67a09a8e30ee" -clientCertPath "C:\GitHub\NRPDev\AppRegistration\PowerShell App Registration.pfx" -clientInteractiveAppID "453f6927-9451-442f-96ab-7748ce984991"

Table of Contents
	1.	Snapin(s), Modules & Environment Connection(s)
	2.	Input & Output Files
	3.	Variables
	4.	Main Script
	5.	CSV Loop Example
	6.	...
#>

#region Script Parameters
Param(	[string] $clientTenant = $null, # IE. ClientTenant.sharepoint.com
		[string] $reportingSiteURL = "{URL of the site that contains the Site Access list}",
		[string] $siteAccessList = "{Title of list script should write to}",
		[boolean] $listCleanup = $false, # If you want to delete all list items before running scan, set to $true
		[string] $inputFile = $null, # Location from _Input Files folder - IE. "sample\filename.csv" or "filename.csv"
		[string] $clientAppID = $null, # Leave Blank for interactive login
		[string] $clientCertPath = $null,# Leave Blank for interactive login
		[string] $clientInteractiveAppID = $null # IE. PnP Now requires an app registration be deployed to support interactive login. Provide the appropriate App ID.
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
Start-Transcript -Path ("..\_Log Files\{0}.PnP-SiteAccess-{1}.log" -f $clientTenant, $scriptStart[1])
#endregion Start Date / Time

#region 1.	Snapin(s), Modules & Environment Connection(s)
# Load MS / PnP PowerShell Modules
Get-Module -ListAvailable | Where-Object{($_.Name -like "*SharePoint*") -or ($_.Name -like "*PnP*") } | Remove-Module
Import-Module PnP.PowerShell
#endregion Snapin(s), Modules & Environment Connection(s)

#region 2.	Input & Output Files
# File Paths
$inPath = ("..\_Input Files\{0}" -f $inputFile)
$errorPath = ("..\_Output Files\{0}.PnP-SiteAccess-Errors-{1}.csv" -f $clientTenant, $scriptStart[1])

# Import Input File(s)
# Load CSV file
If (($inputFile) -and (Test-Path ($inPath))){
	$inFile = Import-Csv $inPath -Encoding Unicode
	$CSVProvided = $true
}
else {
	$CSVProvided = $false
}
#endregion Input & Output Files

#region 3.	Variables
# User Inputs
If (($clientAppID) -or (-not $CSVProvided) -or ($listCleanup)) {
	Write-Host
	Write-Host "User Input Variables" -ForegroundColor Cyan
}
If (-not $CSVProvided) {
	Write-Host "By not providing a CSV file, all tenant sites will be scanned. Continue? (Yes/No): " -NoNewline
	$proceed = Read-Host

	# if user chose to proceed
	if ($proceed.ToLower() -eq "no") {
		Stop-Transcript
		throw "Terminating script so CSV file can be provided"
	}
	Clear-Variable proceed
}
If ($listCleanup) {
	Write-Host "Based on current settings, all current items in the Site Access list will be deleted. Continue? (Yes/No): " -NoNewline
	$proceed = Read-Host

	# if user chose to proceed
	if ($proceed.ToLower() -eq "no") {
		$listCleanup = $false
	}
	Clear-Variable proceed
}
if ($clientAppID) {
	$certPass = ConvertTo-SecureString -String (Read-Host -Prompt "Enter certificate password for the app registration") -AsPlainText -Force
	$cert = Get-PnPAzureCertificate -Path $clientCertPath -Password $certPass
	Remove-Variable certPass
	Remove-Variable clientCertPath
}

# Script Variables
$tenantDomain = ("{0}.onmicrosoft.com" -f $clientTenant)
$spoAdminUrl = ("https://{0}-admin.sharepoint.com" -f $clientTenant)
$maxItems = 4000

#region Output Parameters & Variables
Write-Host
Write-Host "Script Parameters" -ForegroundColor Cyan
Write-Host ("   Client Tenant Prefix: {0}" -f $clientTenant)
Write-Host ("   Reporting Site: {0}" -f $reportingSiteURL)
Write-Host ("   Site Access List: {0}" -f $siteAccessList)
Write-Host ("   List Clean up: {0}" -f $listCleanup)
Write-Host ("   Input File: {0}" -f $inputFile)
if ($clientAppID) {
	Write-Host ("   App ID: {0}" -f $clientAppID)
}
if ($clientInteractiveAppID) {
	Write-Host ("   Interactive App ID: {0}" -f $clientInteractiveAppID)
}
Write-Host
Write-Host "Output Variables" -ForegroundColor Cyan
Write-Host ("   Error Path: {0}" -f $errorPath)
if ($clientAppID) {
	Write-Host
	Write-Host "User Inputs" -ForegroundColor Cyan
	Write-Host ("   Certificate Thumbprint: {0}" -f "Only output for troubleshooting", $cert.Thumbprint)
}
Write-Host
Write-Host "Script Variables" -ForegroundColor Cyan
Write-Host ("   Microsoft Tenant: {0}" -f $tenantDomain)
Write-Host ("   SPO admin URL: {0}" -f $spoAdminUrl)
Write-Host ("   Max Items Page Size: {0}" -f $maxItems)
#endregion Output Parameters & Variables
#endregion
    
#region 4.	Main Script
#region 5.	Connecting To SP Online & Retrieving Sites
Write-Host
Write-Host "Connecting to SharePoint & Retrieving sites" -ForegroundColor Cyan
try {
	#region Connecting to SharePoint
	Write-Host ("   {0} - Connecting to SharePoint" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"))
	try {
		if ($clientAppID) {
			Connect-PnPOnline $spoAdminUrl -tenant $tenantDomain -ClientId $clientAppID -Thumbprint $cert.Thumbprint -ValidateConnection -ErrorAction Stop # PnP Modules w/ App Reg
		} else {
			Connect-PnPOnline -Url $spoAdminUrl -Interactive -ClientId $clientInteractiveAppID -ValidateConnection -ErrorAction Stop # Interactive Login requires PNP Interactive App Reg be deployed by a GA. 
		}
	}
	catch {
		Write-CatErrorCSV -funcCSVItem $null -funcAction "Connecting to SharePoint" -funcException $_.Exception -funcDetails $_.ErrorDetails -funcStackTrace $_.ScriptStackTrace -funcFilePath $errorPath
		Continue listLoop
	}
	#endregion Connecting to SharePoint
	
	#region Retrieving the List of Sites
	write-host ("   {0} - Retrieving the List of Sites" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"))
	if ($CSVProvided){
		$sites = $inFile
	} else {
		$sites = Get-PnPTenantSite
	}
	#endregion Retrieving the List of Sites

	#region Disconnecting from Admin Site
	Disconnect-PnPOnline
	#endregion
}
catch {
	Stop-Transcript
	throw "Unable to connect to tenant or retrieve sites. Terminating Script"
}
#endregion Connecting To SP Online & Retrieving Sites

#region 6.	List Data Cleanup
if ($listCleanup) {
	Write-Host
	Write-Host "Deleting items previously in Site Access list" -ForegroundColor Cyan

	#region Connecting to SharePoint
	write-host ("   {0} - Connecting to Reporting site" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"))
	try {
		if ($clientAppID) {
			Connect-PnPOnline $reportingSiteURL -tenant $tenantDomain -ClientId $clientAppID -Thumbprint $cert.Thumbprint -ValidateConnection -ErrorAction Stop # PnP Modules w/ App Reg
		} else {
			Connect-PnPOnline -Url $reportingSiteURL -Interactive -ClientId $clientInteractiveAppID -ValidateConnection -ErrorAction Stop # Interactive Login requires PNP Interactive App Reg be deployed by a GA. 
		}
	}
	catch {
		Write-CatErrorCSV -funcCSVItem $null -funcAction "Connecting to SharePoint" -funcException $_.Exception -funcDetails $_.ErrorDetails -funcStackTrace $_.ScriptStackTrace -funcFilePath $errorPath
		Continue listLoop
	}
	#endregion Connecting to SharePoint

	# Retrieving List
	write-host ("   {0} - Retrieving list" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"))
	$list = Get-PnPList -Identity $siteAccessList

	# Retrieve all items
	write-host ("   {0} - Retrieving items" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"))
	$items = Get-PnPListItem -List $siteAccessList -PageSize $maxItems
	
	# Deleting Items
	write-host ("   {0} - Deleting items" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"))
	foreach($item in $items)
	{
		Remove-PnPListItem -List $siteAccessList -Identity $item.Id -Force
	}

	# Variable Clean up 
	if ($list) {Remove-Variable list}
	if ($items) {Remove-Variable items}

	# Disconnecting from SharePoint
	Disconnect-PnPOnline
}
#endregion List Data Cleanup

#region Loop through sites
Write-Host
Write-Host "Iterating through SharePoint sites" -ForegroundColor Cyan
if ($sites.count -gt 0) {
	:siteLoop foreach ($site in $sites){
		Write-Host
		write-host ("{0} - Scanning: {1}" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"), $site.URL) -ForegroundColor Magenta
		
		#region Retrieve SharePoint Site
		write-host ("   {0} - Retrieve SharePoint Site Object" -f (Get-Date -Format "yy-MM-dd HH:mm:ss")) -NoNewline

		#region Connecting to SharePoint
		try {
			if ($clientAppID) {
				Connect-PnPOnline $site.Url -tenant $tenantDomain -ClientId $clientAppID -Thumbprint $cert.Thumbprint -ValidateConnection -ErrorAction Stop # PnP Modules w/ App Reg
			} else {
				Connect-PnPOnline -Url $site.Url -Interactive -ClientId $clientInteractiveAppID -ValidateConnection -ErrorAction Stop # Interactive Login requires PNP Interactive App Reg be deployed by a GA. 
			}
		}
		catch {
			Write-CatErrorCSV -funcCSVItem $site -funcAction "Connecting to SharePoint" -funcException $_.Exception -funcDetails $_.ErrorDetails -funcStackTrace $_.ScriptStackTrace -funcFilePath $errorPath
			Continue listLoop
		}
		#endregion Connecting to SharePoint

		#region Retrieving SharePoint Site
		try {
			$scanSite = Get-PnPSite -Includes RootWeb, GroupId
			Write-Host (": {0}" -f $scanSite.RootWeb.Title)
		}
		catch {
			Write-CatErrorCSV -funcCSVItem $site -funcAction "Retrieving SharePoint Site Object" -funcException $_.Exception -funcDetails $_.ErrorDetails -funcStackTrace $_.ScriptStackTrace -funcFilePath $errorPath
		}
		#endregion Retrieving SharePoint Site

		#endregion Retrieve SharePoint Site Object

		#region 7.	Retrieving M365 Group permissions
		if ($scanSite.GroupId.Guid -ne "00000000-0000-0000-0000-000000000000" ) {
			#region Retrieving M365 membership
			Write-Host ("   {0} - Retrieving M365 Group membership" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"))

			#region Connecting to SharePoint
			try {
				if ($clientAppID) {
					Connect-PnPOnline $scanSite.Url -tenant $tenantDomain -ClientId $clientAppID -Thumbprint $cert.Thumbprint -ValidateConnection -ErrorAction Stop # PnP Modules w/ App Reg
				} else {
					Connect-PnPOnline -Url $scanSite.Url -Interactive -ClientId $clientInteractiveAppID -ValidateConnection -ErrorAction Stop # Interactive Login requires PNP Interactive App Reg be deployed by a GA. 
				}
			}
			catch {
				Write-CatErrorCSV -funcCSVItem $site -funcAction "Connecting to SharePoint" -funcException $_.Exception -funcDetails $_.ErrorDetails -funcStackTrace $_.ScriptStackTrace -funcFilePath $errorPath
				Continue listLoop
			}
			#endregion Connecting to SharePoint

			#region Retrieving M365 group members
			Write-Host ("      {0} - Retrieving M365 Group members" -f (Get-Date -Format "yy-MM-dd HH:mm:ss")) -ForegroundColor DarkGray
			try {				
				$members = Get-PnPMicrosoft365GroupMember -identity $scanSite.GroupId.Guid
			}
			catch {
				Write-CatErrorCSV -funcCSVItem $site -funcAction "Retrieving M365 Group Members" -funcException $_.Exception -funcDetails $_.ErrorDetails -funcStackTrace $_.ScriptStackTrace -funcFilePath $errorPath
			}
			#endregion Retrieving M365 group members

			#region Retrieving M365 group owners
			Write-Host ("      {0} - Retrieving M365 Group owners" -f (Get-Date -Format "yy-MM-dd HH:mm:ss")) -ForegroundColor DarkGray
			try {
				$owners = Get-PnPMicrosoft365GroupOwner -identity $scanSite.GroupId.Guid
			}
			catch {
				Write-CatErrorCSV -funcCSVItem $site -funcAction "Retrieving M365 Group Owners" -funcException $_.Exception -funcDetails $_.ErrorDetails -funcStackTrace $_.ScriptStackTrace -funcFilePath $errorPath
			}
			#endregion Retrieving M365 group members

			# Disconnecting from SharePoint
			Disconnect-PnPOnline
			#endregion Retrieving M365 membership

			#region Writing M365 Group Permission to SharePoint List
			Write-Host ("   {0} - Writing M365 Group membership" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"))

			#region Connecting to SharePoint
			try {
				if ($clientAppID) {
					Connect-PnPOnline $reportingSiteURL -tenant $tenantDomain -ClientId $clientAppID -Thumbprint $cert.Thumbprint -ValidateConnection -ErrorAction Stop # PnP Modules w/ App Reg
				} else {
					Connect-PnPOnline -Url $reportingSiteURL -Interactive -ClientId $clientInteractiveAppID -ValidateConnection -ErrorAction Stop # Interactive Login requires PNP Interactive App Reg be deployed by a GA. 
				}
			}
			catch {
				Write-CatErrorCSV -funcCSVItem $site -funcAction "Connecting to SharePoint" -funcException $_.Exception -funcDetails $_.ErrorDetails -funcStackTrace $_.ScriptStackTrace -funcFilePath $errorPath
				Continue listLoop
			}
			#endregion Connecting to SharePoint

			#region Writing Members
			if ($members.count -gt 0){
				Write-Host ("      {0} - Writing M365 Group members" -f (Get-Date -Format "yy-MM-dd HH:mm:ss")) -ForegroundColor DarkGray
				foreach ($member in $members) {
					$spListItem = @{
						"Title" = ("{0}" -f $scanSite.RootWeb.Title);
						"SiteURL" = ("{0}" -f $scanSite.URL);
						"PermissionGivenThrough" = "M365 Group";
						"ObjectType" = "User"
						"ObjectName" = ("{0}" -f $member.UserPrincipalName);
						"ObjectEmail" = ("{0}" -f $member.Email);
						"PermissionGiven" = ("{0}" -f $member.UserType)
					}
					try {
						$temp = Add-PnPListItem -List $siteAccessList -Values $spListItem
					}
					catch {
						Write-CatErrorCSV -funcCSVItem $site -funcAction "Writing M365 Group Members to SharePoint" -funcException $_.Exception -funcDetails $_.ErrorDetails -funcStackTrace $_.ScriptStackTrace -funcFilePath $errorPath
					} 
					Clear-Variable spListItem
					Clear-variable temp
				}
			} else {
				Write-Host ("      {0} - No M365 group Members" -f (Get-Date -Format "yy-MM-dd HH:mm:ss")) -ForegroundColor DarkGray
			}
			#endregion Writing Members
			
			#region Writing Owners
			if ($owners.count -gt 0){
				Write-Host ("      {0} - Writing M365 Group owners" -f (Get-Date -Format "yy-MM-dd HH:mm:ss")) -ForegroundColor DarkGray
				foreach ($owner in $owners) {
					$spListItem = @{
						"Title" = ("{0}" -f $scanSite.RootWeb.Title);
						"SiteURL" = ("{0}" -f $scanSite.URL);
						"PermissionGivenThrough" = "M365 Group";
						"ObjectType" = "User"
						"ObjectName" = ("{0}" -f $owner.UserPrincipalName);
						"ObjectEmail" = ("{0}" -f $owner.Email);
						"PermissionGiven" = "Owner"
					}
					try {
						$temp = Add-PnPListItem -List $siteAccessList -Values $spListItem
					}
					catch {
						Write-CatErrorCSV -funcCSVItem $site -funcAction "Writing M365 Group Owners to SharePoint" -funcException $_.Exception -funcDetails $_.ErrorDetails -funcStackTrace $_.ScriptStackTrace -funcFilePath $errorPath
					}
					Clear-Variable spListItem
					Clear-variable temp
				}
			} else {
				Write-Host ("      {0} - No M365 group Owners" -f (Get-Date -Format "yy-MM-dd HH:mm:ss")) -ForegroundColor DarkGray
			}
			#endregion Writing Members

			# Disconnecting from SharePoint
			Disconnect-PnPOnline
			#endregion Writing M365 Group Permission to SharePoint List
			
			# variable clean up
			if ($spListItem) {Remove-Variable spListItem}
			if ($temp) {Remove-Variable temp}
			if ($members) {Remove-Variable members}
			if ($owners) {Remove-Variable owners}
		} else {
			Write-Host ("   {0} - Not M365 group connected" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"))
		}
		#endregion Retrieving M365 Group permissions #>

		#region 8.	Site Collection Admin permissions
		Write-Host ("   {0} - Site Collection Admin permissions" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"))

		#region Connecting to SharePoint
		try {
			if ($clientAppID) {
				Connect-PnPOnline $scanSite.Url -tenant $tenantDomain -ClientId $clientAppID -Thumbprint $cert.Thumbprint -ValidateConnection -ErrorAction Stop # PnP Modules w/ App Reg
			} else {
				Connect-PnPOnline -Url $scanSite.Url -Interactive -ClientId $clientInteractiveAppID -ValidateConnection -ErrorAction Stop # Interactive Login requires PNP Interactive App Reg be deployed by a GA. 
			}
		}
		catch {
			Write-CatErrorCSV -funcCSVItem $site -funcAction "Connecting to SharePoint" -funcException $_.Exception -funcDetails $_.ErrorDetails -funcStackTrace $_.ScriptStackTrace -funcFilePath $errorPath
			Continue listLoop
		}
		#endregion Connecting to SharePoint

		#region Retrieving Site Collection Administrators
		try {				
			$siteAdmins = Get-PnPSiteCollectionAdmin -Includes UserPrincipalName
		}
		catch {
			Write-CatErrorCSV -funcCSVItem $site -funcAction "Retrieve Site Collection Administrators" -funcException $_.Exception -funcDetails $_.ErrorDetails -funcStackTrace $_.ScriptStackTrace -funcFilePath $errorPath
		}

		# Disconnect from SharePoint
		Disconnect-PnPOnline
		#endregion Retrieve Site Collection Administrators

		#region Writing Site Collection Administrators to SharePoint List
		#region Connecting to SharePoint
		try {
			if ($clientAppID) {
				Connect-PnPOnline $reportingSiteURL -tenant $tenantDomain -ClientId $clientAppID -Thumbprint $cert.Thumbprint -ValidateConnection -ErrorAction Stop # PnP Modules w/ App Reg
			} else {
				Connect-PnPOnline -Url $reportingSiteURL -Interactive -ClientId $clientInteractiveAppID -ValidateConnection -ErrorAction Stop # Interactive Login requires PNP Interactive App Reg be deployed by a GA. 
			}
		}
		catch {
			Write-CatErrorCSV -funcCSVItem $site -funcAction "Connecting to SharePoint" -funcException $_.Exception -funcDetails $_.ErrorDetails -funcStackTrace $_.ScriptStackTrace -funcFilePath $errorPath
			Continue listLoop
		}
		#endregion Connecting to SharePoint
		
		#region Writing to SharePoint
		if ($siteAdmins.count -gt 0){
			foreach ($siteAdmin in $siteAdmins) {
				$spListItem = @{
					"Title" = ("{0}" -f $scanSite.RootWeb.Title);
					"SiteURL" = ("{0}" -f $scanSite.URL);
					"PermissionGivenThrough" = "SharePoint Administration";
					"ObjectType" = ("{0}" -f $siteAdmin.PrincipalType);
					"ObjectName" = ("{0}" -f $siteAdmin.UserPrincipalName);
					"ObjectEmail" = ("{0}" -f $siteAdmin.Email);
					"PermissionGiven" = "Site Collection Administrator"
				}
				try {
					$temp = Add-PnPListItem -List $siteAccessList -Values $spListItem
				}
				catch {
					Write-CatErrorCSV -funcCSVItem $site -funcAction "Writing Site Collection Administrators to SharePoint" -funcException $_.Exception -funcDetails $_.ErrorDetails -funcStackTrace $_.ScriptStackTrace -funcFilePath $errorPath
				} 
				Clear-Variable spListItem
				Clear-variable temp
			}
		} else {
			Write-Host ("      {0} - No Site Collection Administrators" -f (Get-Date -Format "yy-MM-dd HH:mm:ss")) -ForegroundColor DarkGray
		}
		#endregion Writing to SharePoint
		
		# Disconnecting from SharePoint
		Disconnect-PnPOnline

		#endregion Writing Site Collection Administrators to SharePoint List

		# Variable Clean up
		if ($siteAdmins) {Remove-Variable siteAdmins}
		if ($spListItem) {Remove-Variable spListItem}
		if ($temp) {Remove-Variable temp}
		#endregion Site Collection Admin permissions #>

		#region 9.	SharePoint Group permissions
		Write-Host ("   {0} - SharePoint Site permissions" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"))

		#region Connecting to SharePoint
		try {
			if ($clientAppID) {
				Connect-PnPOnline $scanSite.Url -tenant $tenantDomain -ClientId $clientAppID -Thumbprint $cert.Thumbprint -ValidateConnection -ErrorAction Stop # PnP Modules w/ App Reg
			} else {
				Connect-PnPOnline -Url $scanSite.Url -Interactive -ClientId $clientInteractiveAppID -ValidateConnection -ErrorAction Stop # Interactive Login requires PNP Interactive App Reg be deployed by a GA. 
			}
		}
		catch {
			Write-CatErrorCSV -funcCSVItem $site -funcAction "Connecting to SharePoint" -funcException $_.Exception -funcDetails $_.ErrorDetails -funcStackTrace $_.ScriptStackTrace -funcFilePath $errorPath
			Continue listLoop
		}
		#endregion Connecting to SharePoint

		#region Retrieve SharePoint Groups
		try {			
			$spGroups = Get-PnPSiteGroup -Site $scanSite.Url | Where{($_.Roles.count -gt 0) -and ($_.Users.count -gt 0)}
		}
		catch {
			Write-CatErrorCSV -funcCSVItem $site -funcAction "Retrieving SharePoint Groups" -funcException $_.Exception -funcDetails $_.ErrorDetails -funcStackTrace $_.ScriptStackTrace -funcFilePath $errorPath
		}

		# Disconnect from SharePoint
		Disconnect-PnPOnline
		#endregion Retrieve SharePoint Groups

		#region Looping through Groups	
		if ($spGroups.count -gt 0){
			foreach ($group in $spGroups) {
				#region Writing SharePoint Groups to SharePoint List
				$spListItem = @{
					"Title" = ("{0}" -f $scanSite.RootWeb.Title);
					"SiteURL" = ("{0}" -f $scanSite.URL);
					"PermissionGivenThrough" = "SharePoint Administration";
					"ObjectType" = ("{0}" -f "SharePoint Group");
					"ObjectName" = ("{0}" -f $group.Title);
					"PermissionGiven" = ("{0}" -f ($group.Roles -join ";"))
				}

				#region Connecting to SharePoint
				try {
					if ($clientAppID) {
						Connect-PnPOnline $reportingSiteURL -tenant $tenantDomain -ClientId $clientAppID -Thumbprint $cert.Thumbprint -ValidateConnection -ErrorAction Stop # PnP Modules w/ App Reg
					} else {
						Connect-PnPOnline -Url $reportingSiteURL -Interactive -ClientId $clientInteractiveAppID -ValidateConnection -ErrorAction Stop # Interactive Login requires PNP Interactive App Reg be deployed by a GA. 
					}
				}
				catch {
					Write-CatErrorCSV -funcCSVItem $site -funcAction "Connecting to SharePoint" -funcException $_.Exception -funcDetails $_.ErrorDetails -funcStackTrace $_.ScriptStackTrace -funcFilePath $errorPath
					Continue listLoop
				}
				#endregion Connecting to SharePoint

				#region Writting to SharePoint
				try {
					$temp = Add-PnPListItem -List $siteAccessList -Values $spListItem
				}
				catch {
					Write-CatErrorCSV -funcCSVItem $site -funcAction "Writing SharePoint Groups to SharePoint List" -funcException $_.Exception -funcDetails $_.ErrorDetails -funcStackTrace $_.ScriptStackTrace -funcFilePath $errorPath
				} 
				#endregion Writing to SharePoint

				# Variable Cleanup
				Clear-Variable spListItem
				Clear-variable temp

				# Disconnect from SharePoint
				Disconnect-PnPOnline
				#endregion Writing SharePoint Groups to SharePoint List

				#region Retrieve Group Membership
				#region Connecting to SharePoint
				try {
					if ($clientAppID) {
						Connect-PnPOnline $scanSite.Url -tenant $tenantDomain -ClientId $clientAppID -Thumbprint $cert.Thumbprint -ValidateConnection -ErrorAction Stop # PnP Modules w/ App Reg
					} else {
						Connect-PnPOnline -Url $scanSite.Url -Interactive -ClientId $clientInteractiveAppID -ValidateConnection -ErrorAction Stop # Interactive Login requires PNP Interactive App Reg be deployed by a GA. 
					}
				}
				catch {
					Write-CatErrorCSV -funcCSVItem $site -funcAction "Connecting to SharePoint" -funcException $_.Exception -funcDetails $_.ErrorDetails -funcStackTrace $_.ScriptStackTrace -funcFilePath $errorPath
					Continue listLoop
				}
				#endregion Connecting to SharePoint

				#region Retrieving Group membership
				try {			
					$members = Get-PnPGroupMember -Group $group.LoginName					
				}
				catch {
					Write-CatErrorCSV -funcCSVItem $site -funcAction "Retrieving Group Membership" -funcException $_.Exception -funcDetails $_.ErrorDetails -funcStackTrace $_.ScriptStackTrace -funcFilePath $errorPath
				}
				#endregion Retrieving Group membership
				
				# Disconnect from SharePoint
				Disconnect-PnPOnline
				#endregion Retrieve Group Membership

				#region Writing Group Membership to SharePoint List
				if ($members.count -gt 0) {
					#region Connecting to SharePoint
					try {
						if ($clientAppID) {
							Connect-PnPOnline $reportingSiteURL -tenant $tenantDomain -ClientId $clientAppID -Thumbprint $cert.Thumbprint -ValidateConnection -ErrorAction Stop # PnP Modules w/ App Reg
						} else {
							Connect-PnPOnline -Url $reportingSiteURL -Interactive -ClientId $clientInteractiveAppID -ValidateConnection -ErrorAction Stop # Interactive Login requires PNP Interactive App Reg be deployed by a GA. 
						}
					}
					catch {
						Write-CatErrorCSV -funcCSVItem $site -funcAction "Connecting to SharePoint" -funcException $_.Exception -funcDetails $_.ErrorDetails -funcStackTrace $_.ScriptStackTrace -funcFilePath $errorPath
						Continue listLoop
					}
					#endregion Connecting to SharePoint

					#region Looping through Members
					foreach ($member in $members) {
						$spListItem = @{
							"Title" = ("{0}" -f $scanSite.RootWeb.Title);
							"SiteURL" = ("{0}" -f $scanSite.URL);
							"PermissionGivenThrough" = ("{0}" -f $group.Title);
							"ObjectType" = ("{0}" -f $member.PrincipalType);
							"ObjectName" = ("{0}" -f $member.UserPrincipalName);
							"ObjectEmail" = ("{0}" -f $member.Email);
							"PermissionGiven" = ("{0}" -f ($group.Roles -join ";"))
						}
						try {
							$temp = Add-PnPListItem -List $siteAccessList -Values $spListItem
						}
						catch {
							Write-CatErrorCSV -funcCSVItem $site -funcAction "Writing SharePoint Groups to SharePoint List" -funcException $_.Exception -funcDetails $_.ErrorDetails -funcStackTrace $_.ScriptStackTrace -funcFilePath $errorPath
						}
						Clear-Variable spListItem
						Clear-variable temp
					}
					# Variable clean up
					Clear-Variable members
					#endregion Looping through Members

					# Disconnecting from SharePoint
					Disconnect-PnPOnline
				} else {
					Write-Host ("      {0} - No SharePoint Group Members" -f (Get-Date -Format "yy-MM-dd HH:mm:ss")) -ForegroundColor DarkGray
				}
				#endregion Writing Group Membership to SharePoint List
			}
			# Variable Clean up
			Clear-Variable spGroups
		} else {
			Write-Host ("      {0} - No SharePoint Groups" -f (Get-Date -Format "yy-MM-dd HH:mm:ss")) -ForegroundColor DarkGray
		}
		#endregion Looping through Groups
		
		# Variable Clean up
		if ($spGroups) {Remove-Variable spGroups}
		if ($members) {Remove-Variable members}
		if ($spListItem) {Remove-Variable spListItem}
		if ($temp) {Remove-Variable temp}
		#endregion SharePoint Group permissions #>

		#region 10.	Direct Access Permissions
		Write-Host ("   {0} - Direct Access Permissions" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"))

		#region Connecting to SharePoint
		try {
			if ($clientAppID) {
				Connect-PnPOnline $scanSite.Url -tenant $tenantDomain -ClientId $clientAppID -Thumbprint $cert.Thumbprint -ValidateConnection -ErrorAction Stop # PnP Modules w/ App Reg
			} else {
				Connect-PnPOnline -Url $scanSite.Url -Interactive -ClientId $clientInteractiveAppID -ValidateConnection -ErrorAction Stop # Interactive Login requires PNP Interactive App Reg be deployed by a GA. 
			}
		}
		catch {
			Write-CatErrorCSV -funcCSVItem $site -funcAction "Connecting to SharePoint" -funcException $_.Exception -funcDetails $_.ErrorDetails -funcStackTrace $_.ScriptStackTrace -funcFilePath $errorPath
			Continue listLoop
		}
		#endregion Connecting to SharePoint

		#region Retrieving Direct permissions
		try {				
			$users = Get-PnPUser -WithRightsAssignedDetailed -WarningAction SilentlyContinue -ProgressAction SilentlyContinue | Where{$_.Permissions.count -gt 0}
		}
		catch {
			Write-CatErrorCSV -funcCSVItem $site -funcAction "Retrieve Direct permissions" -funcException $_.Exception -funcDetails $_.ErrorDetails -funcStackTrace $_.ScriptStackTrace -funcFilePath $errorPath
		}

		# Disconnecting from SharePoint
		Disconnect-PnPOnline
		#endregion Retrieving Direct permissions

		#region Writing to SharePoint
		#region Connecting to SharePoint
		try {
			if ($clientAppID) {
				Connect-PnPOnline $reportingSiteURL -tenant $tenantDomain -ClientId $clientAppID -Thumbprint $cert.Thumbprint -ValidateConnection -ErrorAction Stop # PnP Modules w/ App Reg
			} else {
				Connect-PnPOnline -Url $reportingSiteURL -Interactive -ClientId $clientInteractiveAppID -ValidateConnection -ErrorAction Stop # Interactive Login requires PNP Interactive App Reg be deployed by a GA. 
			}
		}
		catch {
			Write-CatErrorCSV -funcCSVItem $site -funcAction "Connecting to SharePoint" -funcException $_.Exception -funcDetails $_.ErrorDetails -funcStackTrace $_.ScriptStackTrace -funcFilePath $errorPath
			Continue listLoop
		}
		#endregion Connecting to SharePoint
		
		#region Looping Users
		if ($users.count -gt 0){
			foreach ($user in $users) {
				$permEnum = $user.Permissions.GetEnumerator()

				#region Loop permissions
				while ($permEnum.MoveNext()) {
					$perm = $permEnum.Current
					$spListItem = @{
						"Title" = ("{0}" -f $scanSite.RootWeb.Title);
						"SiteURL" = ("{0}" -f $scanSite.URL);
						"DirectAccessURL" = ("{0}" -f $perm.Key);
						"PermissionGivenThrough" = "Direct Access";
						"ObjectType" = "User";
						"ObjectName" = ("{0}" -f $user.LoginName);
						"ObjectEmail" = ("{0}" -f $user.Email);
						"PermissionGiven" = ("{0}" -f $perm.Value);
					}
					try {
						$temp = Add-PnPListItem -List $siteAccessList -Values $spListItem
					}
					catch {
						Write-CatErrorCSV -funcCSVItem $site -funcAction "Writing Direct Permissions to SharePoint" -funcException $_.Exception -funcDetails $_.ErrorDetails -funcStackTrace $_.ScriptStackTrace -funcFilePath $errorPath
					}
					Clear-Variable perm
					Clear-Variable spListItem
					Clear-variable temp
				}

				Clear-Variable permEnum
				#endregion Loop permissions
			}
		} else {
			Write-Host ("      {0} - No Direct Permissions" -f (Get-Date -Format "yy-MM-dd HH:mm:ss")) -ForegroundColor DarkGray
		}
		#endregion Looping Users

		# Variable Clean up
		if ($users) {Remove-Variable users}
		if ($spListItem) {Remove-Variable spListItem}
		if ($temp) {Remove-Variable temp}

		# Disconnect From Reporting site
		Disconnect-PnPOnline
		#endregion Writing to SharePoint

		#endregion Direct Access Permissions #>

		#region 11.	Share Links Permissions
		Write-Host ("   {0} - Share Links Permissions" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"))

		#region Retrieve Share Links across all lists, libraries, and items
		#region Connecting to SharePoint
		try {
			if ($clientAppID) {
				Connect-PnPOnline $scanSite.Url -tenant $tenantDomain -ClientId $clientAppID -Thumbprint $cert.Thumbprint -ValidateConnection -ErrorAction Stop # PnP Modules w/ App Reg
			} else {
				Connect-PnPOnline -Url $scanSite.Url -Interactive -ClientId $clientInteractiveAppID -ValidateConnection -ErrorAction Stop # Interactive Login requires PNP Interactive App Reg be deployed by a GA. 
			}
		}
		catch {
			Write-CatErrorCSV -funcCSVItem $site -funcAction "Connecting to SharePoint" -funcException $_.Exception -funcDetails $_.ErrorDetails -funcStackTrace $_.ScriptStackTrace -funcFilePath $errorPath
			Continue listLoop
		}
		#endregion Connecting to SharePoint

		#region Retrieving Lists & Items
		Write-Host ("      {0} - Share Links Permissions" -f (Get-Date -Format "yy-MM-dd HH:mm:ss")) -ForegroundColor DarkGray
		$lists = Get-PnPList | where{($_.Hidden -eq $false) -and ($_.ItemCount -gt 0)}
		$items = @()
		foreach ($list in $lists) {
			Get-PnPListItem -List $list -PageSize $maxItems | ForEach-Object {Get-PnPProperty -ClientObject $_ -Property HasUniqueRoleAssignments, File, RoleAssignments; if ($_.HasUniqueRoleAssignments) {$items += $_}}
		}

		# Disconnect from SharePoint
		Disconnect-PnPOnline
		#endregion Retrieving Lists & Items

		#region Loop through Items
		foreach ($item in $items) {
			#region Connecting to SharePoint
			try {
				if ($clientAppID) {
					Connect-PnPOnline $scanSite.Url -tenant $tenantDomain -ClientId $clientAppID -Thumbprint $cert.Thumbprint -ValidateConnection -ErrorAction Stop # PnP Modules w/ App Reg
				} else {
					Connect-PnPOnline -Url $scanSite.Url -Interactive -ClientId $clientInteractiveAppID -ValidateConnection -ErrorAction Stop # Interactive Login requires PNP Interactive App Reg be deployed by a GA. 
				}
			}
			catch {
				Write-CatErrorCSV -funcCSVItem $site -funcAction "Connecting to SharePoint" -funcException $_.Exception -funcDetails $_.ErrorDetails -funcStackTrace $_.ScriptStackTrace -funcFilePath $errorPath
				Continue listLoop
			}
			#endregion Connecting to SharePoint

			#region Retrieving Share Links
			switch ($item.FileSystemObjectType) {
				File {
					$links = Get-PnPFileSharingLink -Identity $item.FieldValues.FileRef # throws an error that appears to be a false negative
					}
				Folder {
					$links = Get-PnPFolderSharingLink -Folder $item.FieldValues.FileRef # throws an error that appears to be a false negative
				}
			}
			Disconnect-PnPOnline
			#endregion Retrieving Share Links
			
			#region Connecting to SharePoint
			try {
				if ($clientAppID) {
					Connect-PnPOnline $reportingSiteURL -tenant $tenantDomain -ClientId $clientAppID -Thumbprint $cert.Thumbprint -ValidateConnection -ErrorAction Stop # PnP Modules w/ App Reg
				} else {
					Connect-PnPOnline -Url $reportingSiteURL -Interactive -ClientId $clientInteractiveAppID -ValidateConnection -ErrorAction Stop # Interactive Login requires PNP Interactive App Reg be deployed by a GA. 
				}
			}
			catch {
				Write-CatErrorCSV -funcCSVItem $site -funcAction "Connecting to SharePoint" -funcException $_.Exception -funcDetails $_.ErrorDetails -funcStackTrace $_.ScriptStackTrace -funcFilePath $errorPath
				Continue listLoop
			}
			#endregion Connecting to SharePoint

			#region Writing links to SharePoint List
			foreach ($link in $links) {
				$spListItem = @{
					"Title" = ("{0}" -f $scanSite.RootWeb.Title);
					"SiteURL" = ("{0}" -f $scanSite.URL);
					"ItemTitle" = ("{0}" -f $link.Parent.Name);
					"ItemURL" = ("{0}" -f $link.Parent.ServerRelativeUrl);
					"DirectAccessURL" = ("{0}" -f $link.link.WebUrl);
					"PermissionGivenThrough" = "Share Link";
					"ObjectType" = "User";
					"ObjectName" = ("{0}" -f $link.GrantedToIdentitiesV2.SiteUser.LoginName);
					"ObjectEmail" = ("{0}" -f $link.GrantedToIdentitiesV2.SiteUser.Email);
					"PermissionGiven" = ("{0}" -f ($links.Roles -join ";"));
				}
				try {
					$temp = Add-PnPListItem -List $siteAccessList -Values $spListItem
				}
				catch {
					Write-CatErrorCSV -funcCSVItem $site -funcAction "Writing ShareLinks to SharePoint" -funcException $_.Exception -funcDetails $_.ErrorDetails -funcStackTrace $_.ScriptStackTrace -funcFilePath $errorPath
				}
				Clear-Variable spListItem
				Clear-variable temp
			}			
			Clear-Variable links
			#endregion Writing links to SharePoint List

			# Disconnecting from SharePoint
			Disconnect-PnPOnline
		}
		#endregion Loop through Items
		
		# Variable Clean up
		if ($items) {Remove-Variable items}
		if ($lists) {Remove-Variable lists}
		if ($links) {Remove-Variable links}
		if ($spListItem) {Remove-Variable spListItem}
		if ($temp) {Remove-Variable temp}
		#endregion Retrieve Share Links across all lists, libraries, and items

		#endregion Share Links Permissions #>

		# Variable Clean up
		if ($scanSite) {Remove-Variable scanSite}
	} 
} else {
	Write-Host "*** No sites provided ***" -ForegroundColor Red
}
#endregion Loop through sites

# Variable Clean up
Remove-Variable sites
#endregion Main Script

#region Completed Script
# Output Script Info
Complete-CatScript $scriptStart[0]

# Stop logging
Stop-Transcript
#endregion Completed Script