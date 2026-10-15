<#  Cat.PnP-SharePointDiscovery.ps1 Description

This script leverages the list sites exported from SharePoint Administration / Active Sites to peform additional discovery. The script 
iterates through the list and outputs the information listed below.

	-	Site Level permissions (Owners/Members) for both sharepoint permissions and M365 Group membership
	-	TODO: External Sharing, Public Sharing Links (annonymous)

Dependencies
	-	PnP.PowerShell Modules
	-	Input CSV File - Example: ..\_Input Files\Samples\Reporting\Cat.PnP-SharePointDiscovery-Test01.csv

Call script with parameters
	cd c:\GitHub\CatPowerShell\Reporting
	.\Cat.PnP-SharePointDiscovery.ps1  -inputFile Samples\Reporting\Cat.PnP-SharePointDiscovery-Test02.csv -clientAppID "caf6b11a-cf03-416c-b0f0-67a09a8e30ee" -clientCertPath "C:\GitHub\NRPDev\AppRegistration\PowerShell App Registration.pfx" -clientTenant "nrpdev2" -outputLocation "C:\GitHub\CatPowerShell\_Output Files"

Table of Contents
	1.	Snapin(s), Modules & Environment Connection(s)
	2.	Input & Output Files
	3.	Variables
	4.	Main Script
	5.	Looping Sites
	6.	Collect Site Admins
	7.	Collect SharePoint Groups
	8.	Collect SharePoint group user permissions
	9.	Collect M365 Group membership
#>

#region Script Parameters
Param(	[string] $inputFile = "{Teams input file path relative to _Input Files}", # Location from _Input Files folder - IE. "sample\filename.csv" or "filename.csv"; Must be saved as UTF-8 in excel
		[string] $clientAppID = "{Client app registration ID from Azure}", # Leave Blank for interactive login
		[string] $clientCertPath = "{Path to client certificate file must be installed on local user}",# Leave Blank for interactive login
		[string] $clientTenant = "{Client tenant SharePoint prefix}", # IE. ClientTenant.sharepoint.com
		[string] $outputLocation = "{c:\path to output files location}" # A /clienttenant/datetime/ folder will be created for output files
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
Start-Transcript -Path ("..\_Log Files\{0}.PnP-SharePointDiscovery-{1}.log" -f $clientTenant, $scriptStart[1])
#endregion Start Date / Time

#region 1.	Snapin(s), Modules & Environment Connection(s)
# Load MS / PnP PowerShell Modules
Get-Module -ListAvailable | Where-Object{($_.Name -like "*SharePoint*") -or ($_.Name -like "*PnP*") } | Remove-Module
Import-Module PnP.PowerShell
#endregion Snapin(s), Modules & Environment Connection(s)

#region 2.	Input & Output Files
# File Paths
$inPath = ("..\_Input Files\{0}" -f $inputFile)
$outputPath = ("{0}\{1}\{2}" -f $outputLocation, $clientTenant, $scriptStart[1])
$errorPath = ("{0}\errors.csv" -f $outputPath)
$outputPerms = ("{0}\Permissions.csv" -f $outputPath)

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
If (($clientAppID) -or ($boolean)) {
	Write-Host
	Write-Host "User Input Variables" -ForegroundColor Cyan
	$certPass = ConvertTo-SecureString -String (Read-Host -Prompt "   Enter certificate password for the app registration") -AsPlainText -Force
	$cert = Get-PnPAzureCertificate -Path $clientCertPath -Password $certPass
	Remove-Variable certPass
	Remove-Variable clientCertPath
}

# Script Variables
$tenantDomain = ("{0}.onmicrosoft.com" -f $clientTenant)
$spoAdminUrl = ("https://{0}-admin.sharepoint.com" -f $clientTenant)
$specialCharacters = "[``'`":]" #(($pnpGroup.Description -replace $specialCharacters,'') -replace "`n",' ') -replace "`r",' '

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
Write-Host ("   Output Path: {0}" -f $outputPath)
Write-Host ("   Error Path: {0}" -f $errorPath)
Write-Host ("   Permissions Path: {0}" -f $outputPerms)
If (($clientAppID) -or ($boolean)) {
	Write-Host
	Write-Host "User Inputs" -ForegroundColor Cyan
	Write-Host ("   Certificate Thumbprint: {0}" -f "Only output for troubleshooting", $cert.Thumbprint)
}
Write-Host
Write-Host "Script Variables" -ForegroundColor Cyan
Write-Host ("   Microsoft Tenant: {0}" -f $tenantDomain)
Write-Host ("   SPO admin URL: {0}" -f $spoAdminUrl)
Write-Host ("   Special Characters to Remove: {0}" -f $specialCharacters)
#endregion Output Parameters & Variables
#endregion

#region 4.	Main Script
#region Created directories for reports
Write-Host
Write-Host ("{0} - Creating Report Directories" -f (Get-Date -Format "yy-MM-dd HH:mm:ss")) -ForegroundColor Cyan
New-Item -Path $outputPath -ItemType Directory
#endregion Created directories for reports

#region	5.	Looping Sites
Write-Host
Write-Host ("{0} - Looping through sites from SharePoint Administration/Active Sites export" -f (Get-Date -Format "yy-MM-dd HH:mm:ss")) -ForegroundColor Cyan
:siteLoop foreach ($site in $inFile){
	#region Connecting To SP Online
	Write-Host
	Write-Host ("   {0} - Connecting to SharePoint site: {1}" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"), $site.'Site name') -ForegroundColor Magenta
	try {
		if ($clientAppID) {
			Connect-PnPOnline $site.URL -tenant $tenantDomain -ClientId $clientAppID -Thumbprint $cert.Thumbprint -ValidateConnection -ErrorAction Stop # PnP Modules w/ App Reg
		} else {
			Connect-PnPOnline -Url $site.URL -Interactive -ValidateConnection -ErrorAction Stop # PnP Modules w/o App Reg
		} 
	}
	catch {
		Write-CatErrorCSV -funcCSVItem $site -funcAction "Connecting to SharePoint" -funcException $_.Exception -funcDetails $_.ErrorDetails -funcStackTrace $_.ScriptStackTrace -funcFilePath $errorPath
		Continue siteLoop
	}
	#endregion Connecting To SP Online

	#region	6.	Collect Site Admins
	try {
		Write-Host ("      {0} - Retrieving site collection administrators" -f (Get-Date -Format "yy-MM-dd HH:mm:ss")) -ForegroundColor DarkGray
		$siteAdmins = Get-PnPSiteCollectionAdmin -Includes UserPrincipalName
		foreach ($admin in $siteAdmins) {
			$csvItem = New-Object -TypeName PSObject -Property ([ordered]@{`
				siteTitle=$site.'Site name';`
				siteURL=$site.URL;`
				title=$admin.Title;`
				principalType=$admin.PrincipalType;`
				LoginName=$admin.UserPrincipalName;`
				Email=$admin.Email;`
				IsSiteAdmin=$admin.IsSiteAdmin;`
				Groups="";`
				Permissions="Site Collection Administrator";
			})
			$csvItem | Export-Csv -Path $outputPerms -UseQuotes AsNeeded -NoTypeInformation -Append
			Remove-Variable csvItem
		}
		Remove-Variable siteAdmins
	}
	catch {
		Write-CatErrorCSV -funcCSVItem $site -funcAction "Retrieving site collection administrators" -funcException $_.Exception -funcDetails $_.ErrorDetails -funcStackTrace $_.ScriptStackTrace -funcFilePath $errorPath
	}	
	#endregion 6.	Collect Site Permission

	#region 7.	Collect SharePoint Groups
	try {
		Write-Host ("      {0} - Retrieving SharePoint groups" -f (Get-Date -Format "yy-MM-dd HH:mm:ss")) -ForegroundColor DarkGray
		$spGroups = Get-PnPSiteGroup
		foreach ($group in $spGroups) {
			$csvItem = New-Object -TypeName PSObject -Property ([ordered]@{`
				siteTitle=$site.'Site name';`
				siteURL=$site.URL;`
				title=$group.Title;`
				principalType="SharePointGroup";`
				LoginName="";`
				Email="";`
				IsSiteAdmin=$false;`
				Groups="";`
				Permissions=($group.Roles -join ";");
			})
			$csvItem | Export-Csv -Path $outputPerms -UseQuotes AsNeeded -NoTypeInformation -Append
			Remove-Variable csvItem
			Remove-Variable group
		}
		Remove-Variable spGroups
	}
	catch {
		Write-CatErrorCSV -funcCSVItem $site -funcAction "Retrieving SharePoint groups" -funcException $_.Exception -funcDetails $_.ErrorDetails -funcStackTrace $_.ScriptStackTrace -funcFilePath $errorPath
	}	
	#endregion 7.	Collect SharePoint Groups

	#region 8.	Collect SharePoint group user permissions
	try {
		Write-Host ("      {0} - Retrieving users" -f (Get-Date -Format "yy-MM-dd HH:mm:ss")) -ForegroundColor DarkGray
		$users = Get-PnPUser -WithRightsAssignedDetailed -WarningAction SilentlyContinue
		foreach ($user in $users) {
			$userDetail = Get-PnPUser -WithRightsAssigned -Includes UserPrincipalName | Where {$user.LoginName -eq $_.LoginName}
			$csvItem = New-Object -TypeName PSObject -Property ([ordered]@{`
				siteTitle=$site.'Site name';`
				siteURL=$site.URL;`
				title=$user.Title;`
				principalType=$userDetail.PrincipalType;`
				LoginName=$userDetail.UserPrincipalName;`
				Email=$user.Email;`
				IsSiteAdmin=$userDetail.IsSiteAdmin;`
				Groups=($user.Groups -join ";");`
				Permissions=($user.Permissions.Values -join ";");
			})
			$csvItem | Export-Csv -Path $outputPerms -UseQuotes AsNeeded -NoTypeInformation -Append
			Remove-Variable csvItem
			Remove-Variable userDetail
			Remove-variable user
		}
		Remove-Variable users
	}
	catch {
		Write-CatErrorCSV -funcCSVItem $site -funcAction "Retrieving users" -funcException $_.Exception -funcDetails $_.ErrorDetails -funcStackTrace $_.ScriptStackTrace -funcFilePath $errorPath
	}
	#endregion 8.	Collect SharePoint group user permissions

	#region 9.	Collect M365 Group membership
	#region Members
	try {
		if ($site.'Microsoft 365 group' -eq "TRUE") {
			Write-Host ("      {0} - Retrieving M365 Group members" -f (Get-Date -Format "yy-MM-dd HH:mm:ss")) -ForegroundColor DarkGray
			$members = Get-PnPMicrosoft365GroupMember -identity $site.'Site name'
			foreach ($member in $members) {
				$csvItem = New-Object -TypeName PSObject -Property ([ordered]@{`
					siteTitle=$site.'Site name';`
					siteURL=$site.URL;`
					title=$site.'Site name';`
					principalType="M365GroupMembership";`
					LoginName=$member.UserPrincipalName;`
					Email=$member.Email;`
					IsSiteAdmin="";`
					Groups=$site.'Site name';`
					Permissions=$member.UserType;
				})
				$csvItem | Export-Csv -Path $outputPerms -UseQuotes AsNeeded -NoTypeInformation -Append
				Remove-Variable csvItem
				Remove-Variable member
			}
			Remove-Variable members	
		}
	}
	catch {
		Write-CatErrorCSV -funcCSVItem $site -funcAction "Retrieving M365 Group members" -funcException $_.Exception -funcDetails $_.ErrorDetails -funcStackTrace $_.ScriptStackTrace -funcFilePath $errorPath
	}
	#endregion Members

	#region Owners
	try {
		if ($site.'Microsoft 365 group' -eq "TRUE") {
			Write-Host ("      {0} - Retrieving M365 Group owners" -f (Get-Date -Format "yy-MM-dd HH:mm:ss")) -ForegroundColor DarkGray
			$owners = Get-PnPMicrosoft365GroupOwner -identity $site.'Site name'
			foreach ($owner in $owners) {
				$csvItem = New-Object -TypeName PSObject -Property ([ordered]@{`
					siteTitle=$site.'Site name';`
					siteURL=$site.URL;`
					title=$site.'Site name';`
					principalType="M365GroupMembership";`
					LoginName=$owner.UserPrincipalName;`
					Email=$owner.Email;`
					IsSiteAdmin="";`
					Groups=$site.'Site name';`
					Permissions=$owner.UserType;
				})
				$csvItem | Export-Csv -Path $outputPerms -UseQuotes AsNeeded -NoTypeInformation -Append
				Remove-Variable csvItem
				Remove-Variable owner
			}
			Remove-Variable owners	
		}
	}
	catch {
		Write-CatErrorCSV -funcCSVItem $site -funcAction "Retrieving M365 Group owners" -funcException $_.Exception -funcDetails $_.ErrorDetails -funcStackTrace $_.ScriptStackTrace -funcFilePath $errorPath
	}
	#endregion Owners
	#endregion 9.	Collect M365 Group membership

	# Disconnect from Services
	Write-Host ("      {0} - Disconnecting from SharePoint site - {1}" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"), $site.'Site name') -ForegroundColor DarkGray
	Disconnect-PnPOnline
} 
#endregion 5.	Looping Sites
#endregion 4.	Main Script

#region Completed Script
# Output Script Info
Complete-CatScript $scriptStart[0]

# Stop logging
Stop-Transcript
#endregion Completed Script