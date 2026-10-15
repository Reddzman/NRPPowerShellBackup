<#  Cat.Tech-ScriptName.ps1 Description

This script takes a CSV file with a list of sites and reports on all locations where permissions inheritance has been broken. It 
Creates 2 reports: Locations & Permissions.  The Locations report can be used as an input for the Cat.PnP-ResetInheritanceByCSV.ps1
script (Configuration\Cat.PnP-ResetInheritanceByCSV.ps1)

Dependencies
	-	PnP.PowerShell Modules
	-	Input CSV File - Example: ..\_Input Files\Samples\Reporting\Cat.PnP-BrokenInheritanceBySite-Test.csv

Call script with parameters
	cd c:\GitHub\CatPowershell\Reporting
	.\Cat.PnP-BrokenInheritanceBySite.ps1  -inputFile Samples\Reporting\Cat.PnP-BrokenInheritanceBySite-Test.csv -clientAppID "caf6b11a-cf03-416c-b0f0-67a09a8e30ee" -clientCertPath "C:\GitHub\NRPDev\AppRegistration\PowerShell App Registration.pfx" -clientTenant "nrpdev2"

Table of Contents
	1.	Snapin(s), Modules & Environment Connection(s)
	2.	Input & Output Files
	3.	Variables
	4.	Main Script
	5.	CSV Loop Example
	6.	...
#>

#region Script Parameters
Param(	[string] $inputFile = "{Teams input file path relative to _Input Files}", # Location from _Input Files folder - IE. "sample\filename.csv" or "filename.csv"
		[string] $clientAppID = "{Client app registration ID from Azure}", # Leave Blank for interactive login
		[string] $clientCertPath = "{Path to client certificate file must be installed on local user}",# Leave Blank for interactive login
		[string] $clientTenant = "{Client tenant SharePoint prefix}" # IE. ClientTenant.sharepoint.com
)
#endregion Script Parameters

#region Start Date / Time
# Remove Existing Catapult Modules to ensure PS session has most up to date functions
Get-Module Cat.* | Remove-Module

# Load General Functions
Import-Module "..\General\Cat.GeneralFunctions.psm1"

# Get Start Date / Time
$scriptStart = Start-CatScript

# Start Transcript
Start-Transcript -Path ("..\_Log Files\{0}.PnP-BrokenInheritanceBySite-{1}.log" -f $clientTenant, $scriptStart[1])
#endregion Start Date / Time

#region 1.	Snapin(s), Modules & Environment Connection(s)
# Load MS / PnP PowerShell Modules
Get-Module -ListAvailable | Where-Object{($_.Name -like "*SharePoint*") -or ($_.Name -like "*PnP*") } | Remove-Module
Import-Module PnP.PowerShell
#endregion Snapin(s), Modules & Environment Connection(s)

#region 2.	Input & Output Files
# File Paths
$inPath = ("..\_Input Files\{0}" -f $inputFile)
$outFolder = ("..\_Output Files\{0}\{1}" -f $clientTenant,$scriptStart[1])
$outLocations = ("{0}\LocationsWithBrokenEnheritance.csv" -f $outFolder)
$outPermissions = ("{0}\PermissionsOnBrokenEnheritanceLocations.csv" -f $outFolder)
$errorPath = ("{0}\errors.csv" -f $outFolder)

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
If ($clientAppID) {
	Write-Host
	Write-Host "User Input Variables" -ForegroundColor Cyan
	$certPass = ConvertTo-SecureString -String (Read-Host -Prompt "   Enter certificate password for the app registration") -AsPlainText -Force
	$cert = Get-PnPAzureCertificate -Path $clientCertPath -Password $certPass
	Remove-Variable certPass
	Remove-Variable clientCertPath
}

# Script Variables
$tenantDomain = ("{0}.onmicrosoft.com" -f $clientTenant)
$spoURL = ("https://{0}.sharepoint.com" -f $clientTenant)
$specialCharacters = "[~``'`":^]" #(($pnpGroup.Description -replace $specialCharacters,'') -replace "`n",' ') -replace "`r",' '
$pageSize = 4500

#region Output Parameters & Variables
Write-Host
Write-Host "Script Parameters" -ForegroundColor Cyan
Write-Host ("   Input File: {0}" -f $inputFile)
if ($clientAppID) {
	Write-Host ("   App ID: {0}" -f $clientAppID)
}
Write-Host ("   Client Tenant Prefix: {0}" -f $clientTenant)
Write-Host
Write-Host "Output Variables" -ForegroundColor Cyan
Write-Host ("   Folder Path: {0}" -f $outFolder)
Write-Host ("   Location Path: {0}" -f $outLocations)
Write-Host ("   Permissions Path: {0}" -f $outPermissions)
Write-Host ("   Error Path: {0}" -f $errorPath)
if ($clientAppID) {
	Write-Host ("   Certificate Thumbprint: {0}" -f "Only output for troubleshooting", $cert.Thumbprint)
}
Write-Host
Write-Host "Script Variables" -ForegroundColor Cyan
Write-Host ("   Microsoft Tenant: {0}" -f $tenantDomain)
Write-Host ("   SPO URL: {0}" -f $spoURL)
Write-Host ("   Special Characters to Remove: {0}" -f $specialCharacters)
Write-Host ("   Page Size: {0}" -f $pageSize)
#endregion Output Parameters & Variables
#endregion Variables
    
#region 4.	Main Script
#region Creating directories for reports
Write-Host
Write-Host ("{0} - Creating Report Directories" -f (Get-Date -Format "yy-MM-dd HH:mm:ss")) -ForegroundColor Cyan
try {
	New-Item -Path $outFolder -ItemType Directory
}
catch {
	Write-CatErrorCSV -funcCSVItem $inFile[0] -funcAction "Creating Report Directories" -funcException $_.Exception -funcDetails $_.ErrorDetails -funcStackTrace $_.ScriptStackTrace -funcFilePath $errorPath
	Continue listLoop
}
#endregion Creating directories for reports

#region Loop through site list
Write-Host
Write-Host ("{0} - Iterating through site list ({1})" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"), $inFile.Count) -ForegroundColor Cyan
:loopSites foreach ($site in $inFile){
	Write-Host
	Write-Host ("   {0} - Checking Site for Broken Inheritiance - {1}" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"), $site.siteTitle) -ForegroundColor Magenta

	#region Connecting To SP Online
	Write-Host ("      {0} - Connecting to site" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"))
	try {
		if ($clientAppID) {
			Connect-PnPOnline $site.siteURL -tenant $tenantDomain -ClientId $clientAppID -Thumbprint $cert.Thumbprint -ValidateConnection -ErrorAction Stop # PnP Modules w/ App Reg
		} else {
			Connect-PnPOnline -Url $site.siteURL -Interactive -ValidateConnection -ErrorAction Stop # PnP Modules w/o App Reg
		} 
	}
	catch {
		Write-CatErrorCSV -funcCSVItem $site -funcAction "Connecting to site" -funcException $_.Exception -funcDetails $_.ErrorDetails -funcStackTrace $_.ScriptStackTrace -funcFilePath $errorPath
		Continue listLoop
	}
	#endregion Connecting To SP Online

	#region Retrieving all lists and libraries
	Write-Host ("      {0} - Retrieving lists and libraries" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"))
	try {
		$lists = Get-PnPList -Includes HasUniqueRoleAssignments, ParentWeb, Fields, RoleAssignments | where{$_.Hidden -eq $false}
	}
	catch {
		Write-CatErrorCSV -funcCSVItem $site -funcAction "Retrieving lists and libraries" -funcException $_.Exception -funcDetails $_.ErrorDetails -funcStackTrace $_.ScriptStackTrace -funcFilePath $errorPath
		Continue loopSites
	}
	#endregion Retrieving lists and libraries

	#region Loop through lists and libraries
	Write-Host ("      {0} - Iterating through lists and libraries ({1})" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"), $lists.Count)
	:loopLists foreach ($list in $lists){
		if ($list.HasUniqueRoleAssignments) {
			#region Write List to CSV if it has unique permissions
			Write-Host ("         {0} - {1} - Outputing Location Information" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"), $list.Title) -ForegroundColor DarkGray -NoNewline
			$csvItem = New-Object -TypeName PSObject -Property (@{`
				objectType=$list.BaseType;`
				siteTitle=((($list.ParentWeb.Title -replace $specialCharacters,'') -replace "`n",' ') -replace "`r",' ');`
				siteURL=$list.ParentWeb.URL;`
				listTitle=((($list.Title -replace $specialCharacters,'') -replace "`n",' ') -replace "`r",' ');`
				listURL=("{0}{1}" -f $spoURL,$list.RootFolder.ServerRelativeUrl);`
				itemTitle="";`
				itemURL="";`
				itemID="";`
			})
			$csvItem | Select-Object -Property objectType, siteTitle, siteURL, listTitle, listURL, itemTitle, itemURL, itemID | Export-Csv -Path $outLocations -UseQuotes AsNeeded -NoTypeInformation -Append -Encoding Unicode
			Remove-Variable csvItem
			#endregion Write List to CSV if it has unique permissions

			#region Writing unique permissions
			Write-Host " - Outputing Permission Information" -ForegroundColor DarkGray
			:loopListPerms foreach ($assignment in $list.RoleAssignments) {
				try {
					$member = Get-PnPProperty -ClientObject $assignment -Property Member
					$permission = Get-PnPListPermissions -Identity $list -PrincipalId $member.Id
				}
				catch {
					Write-CatErrorCSV -funcCSVItem $site -funcAction "Outputing List Permission Information" -funcException $_.Exception -funcDetails $_.ErrorDetails -funcStackTrace $_.ScriptStackTrace -funcFilePath $errorPath
					Continue loopListPerms
				}
				$csvItem = New-Object -TypeName PSObject -Property (@{`
					objectType=$list.BaseType;`
					objectTitle=((($list.Title -replace $specialCharacters,'') -replace "`n",' ') -replace "`r",' ');`
					objectURL=("{0}{1}" -f $spoURL,$list.RootFolder.ServerRelativeUrl);`
					LoginType=$member.PrincipalType;`
					LoginName=((($member.LoginName -replace $specialCharacters,'') -replace "`n",' ') -replace "`r",' ');`
					Permission=$permission.Name;`
				})
				$csvItem | Select-Object -Property objectType, objectTitle, objectURL, LoginType, LoginName, Permission | Export-Csv -Path $outPermissions -UseQuotes AsNeeded -NoTypeInformation -Append
				Remove-Variable csvItem
			}
			#endregion Writing unique permissions
		}
		
		#region Retrieving List and libary items with unique permissions
		Write-Host ("         {0} - {1} - Retrieving list or library items" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"), $list.Title) -ForegroundColor DarkGray -NoNewline
		try {
			$items = @()
			Get-PnPListItem -List $list -PageSize $pageSize | ForEach-Object {Get-PnPProperty -ClientObject $_ -Property HasUniqueRoleAssignments, File, RoleAssignments; if ($_.HasUniqueRoleAssignments) {$items += $_}}
		}
		catch {
			Write-CatErrorCSV -funcCSVItem $site -funcAction "Retrieving items" -funcException $_.Exception -funcDetails $_.ErrorDetails -funcStackTrace $_.ScriptStackTrace -funcFilePath $errorPath
			Continue loopSites
		}
		#endregion Retrieving List and libary items with unique permissions

		#region Loop through list and library items
		If ($items.Count -gt 0) {
			Write-Host
			Write-Host ("         {0} - {1} - Iterating through list and library items ({2})" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"), $list.Title, $items.Count) -ForegroundColor DarkGray
			:loopItems foreach ($item in $items) {
				#region Write Item to CSV
				Write-Host ("            {0} - {1} - Outputing Location Information" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"), ((($item.FieldValues.FileLeafRef -replace $specialCharacters,'') -replace "`n",' ') -replace "`r",' ')) -ForegroundColor DarkBlue -NoNewline
				$csvItem = New-Object -TypeName PSObject -Property (@{`
					objectType=switch ($item.FileSystemObjectType) {Folder {$item.FileSystemObjectType} File {If($null -ne $item.File.Name){$item.FileSystemObjectType} Else {"Item"}}};`
					siteTitle=((($list.ParentWeb.Title -replace $specialCharacters,'') -replace "`n",' ') -replace "`r",' ');`
					siteURL=$list.ParentWeb.URL;`
					listTitle=((($list.Title -replace $specialCharacters,'') -replace "`n",' ') -replace "`r",' ');`
					listURL=("{0}{1}" -f $spoURL,$list.RootFolder.ServerRelativeUrl);`
					itemTitle=((($item.FieldValues.FileLeafRef -replace $specialCharacters,'') -replace "`n",' ') -replace "`r",' ');`
					itemURL=("{0}{1}" -f $spoURL,$item.FieldValues.FileRef);`
					itemID=$item.Id;`
				})
				$csvItem | Select-Object -Property objectType, siteTitle, siteURL, listTitle, listURL, itemTitle, itemURL, itemID | Export-Csv -Path $outLocations -UseQuotes AsNeeded -NoTypeInformation -Append -Encoding Unicode
				Remove-Variable csvItem
				#endregion Write Item to CSV

				#region Retrieving item permissions
				try {
					$permissions = Get-PnPListItemPermission -List $list -Identity $item.Id
				}
				catch {
					Write-CatErrorCSV -funcCSVItem $site -funcAction "Outputing Item Permission Information" -funcException $_.Exception -funcDetails $_.ErrorDetails -funcStackTrace $_.ScriptStackTrace -funcFilePath $errorPath
					Continue loopItems
				}
				#endregion Retrieving item permissions

				#region Writing permissions
				If ($permissions.Count -gt 0) {
					Write-Host " - Outputing Permission Information" -ForegroundColor DarkBlue
					:loopItemPerms foreach ($permission in $permissions.Permissions) {
						$csvItem = New-Object -TypeName PSObject -Property (@{`
							objectType=switch ($item.FileSystemObjectType) {Folder {$item.FileSystemObjectType} File {If($null -ne $item.File.Name){$item.FileSystemObjectType} Else {"Item"}}};`
							objectTitle=((($item.FieldValues.FileLeafRef -replace $specialCharacters,'') -replace "`n",' ') -replace "`r",' ');`
							objectURL=("{0}{1}" -f $spoURL,$item.FieldValues.FileRef);`
							LoginType=$permission.PrincipalType;`
							LoginName=$permission.PrincipalName;`
							Permission=$permission.PermissionLevels.Name -join ";";`
						})
						$csvItem | Select-Object -Property objectType, objectTitle, objectURL, LoginType, LoginName, Permission | Export-Csv -Path $outPermissions -UseQuotes AsNeeded -NoTypeInformation -Append
						Remove-Variable csvItem
					}
				} else {
					Write-Host " - No permission found" -ForegroundColor DarkBlue
				}
				#endregion Writing permissions
			}
		} Else {
			Write-Host " - No items with custom permissions" -ForegroundColor DarkGray
		}
		Remove-Variable items
		#endregion Retrieving List and libary items with unique permissions
	}
	Remove-Variable lists
	#endregion Loop through lists and libraries

	#region Disconnect from SP Online
	Write-Host ("      {0} - Disconnecting to site" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"))
	Disconnect-PnPOnline
	#endregion Disconnect from SP Online
} 
#endregion Loop through site list

#endregion Main Script

#region Completed Script
# Output Script Info
Complete-CatScript $scriptStart[0]

# Stop logging
Stop-Transcript
#endregion Completed Script