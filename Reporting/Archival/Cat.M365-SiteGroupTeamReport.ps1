<#	Cat.M365-SiteGroupTeam.ps1

This script iterates through all SharePoint Online Sites to generate a report on Sites, O365 Groups, and Teams that existing in
Microsoft 365. Because this script iterates through SharePoint as the starting point, 0365 groups not tied to SharePoint Sites 
are not included in this report.

Dependencies:
	-	AzureADPreview PowerShell Module
	-	Microsoft.Online.SharePoint.PowerShell PowerShell Module
	-	SharePointPnPPowershellOnline PowerShell Module
	-	MicrosoftTeams PowerShell Module PowerShell Module

Table of Contents
	1.	Input & Output Files
	2.	Variables
	3.	Main Script
#>

#region Script Parameters
Param(	[string] $spAdminURL = "https://nrpdev-admin.sharepoint.com",
		[string] $client = "Testing",
		[bool] $includePersonalSite = $false)
#endregion

#region Start Date / Time
# Remove Existing Catapult Modules to ensure PS session has most up to date functions
Get-Module Cat.* | Remove-Module

# Load General Functions
Import-Module "..\General\Cat.GeneralFunctions.psm1"

# Get Start Date / Time
$scriptStart = Start-CatScript

# Start Transcript
Start-Transcript -Path ("..\_Log FIles\{0}.M365-CollabSiteReport-{1}.log" -f $client, $scriptStart[1])
#endregion

#region 1.	Input & Output Files
# File Paths
$outSitesPath = ("..\_Output Files\{0}.CollabSiteReport-{1}.txt" -f $client, $scriptStart[1])
$outUsersPath = ("..\_Output Files\{0}.CollabUsersReport-{1}.txt" -f $client, $scriptStart[1])

# Create Output File(s)
Out-File -FilePath $outSitesPath -InputObject "Type|Name|URL|SPModified|SPStorage|SPID|SPCreated|SPOwner|SPSecondaryOwner|SPTemplate|IsHubSite|HubSiteID|IsGroup|GroupID|RelatedGroupID|IsTeam|MailNickName|PWAEnabled"
Out-File -FilePath $outUsersPath -InputObject "SPGUID|ID|User|Name|RoleOrGroups"
#endregion

#region 2.	Variables
# Variable Output
Write-Host
Write-Host "Parameters" -ForegroundColor Cyan
Write-Host ("   Admin URL: {0}" -f $spAdminURL)
Write-Host ("   Client: {0}" -f $client)
Write-Host ("   Include Personal Sites: {0}" -f $includePersonalSite)
Write-Host
Write-Host "Output Files" -ForegroundColor Cyan
Write-Host ("   Sites Report: {0}" -f $outSitesPath)
Write-Host ("   Users Report: {0}" -f $outSitesPath)
#endregion

#region 3.	Snapin(s), Modules & Environment Connection(s)
# Load custom modules
Import-Module "..\Reporting\Cat.ReportingFunctions.psm1"
#endregion

#region 3.	Main Script
#region Connecting to M365 services
Write-Host
Write-Host "Connecting to M365 Services" -ForegroundColor Cyan
Write-Host ("   Connecting to SharePoint Online with the SharePoint Management Shell: {0}" -f $spAdminURL)
Connect-SPOService -Url $spAdminURL
Write-Host ("   Connecting to SharePoint Online with PnP PowerShell: {0}" -f $spAdminURL)
$pnpConnect = Connect-PnPOnline -Url $spAdminURL.Replace("-admin","").Replace("-Admin","") -PnPO365ManagementShell -ReturnConnection -LaunchBrowser # *** The -LaunchBrowser flag may have issues with MFA ***
$pnpToken = Get-PnPAccessToken
Disconnect-PnPOnline
Write-Host
Write-Host "   Connecting to Microsoft Teams"
$teamsConnect = Connect-MicrosoftTeams
Write-Host "   Connecting to Azure AD"
$teamsConnect = Connect-AzureAD
#endregion

# Screen Output
Write-Host
Write-Host "Scanning M365 Environment" -ForegroundColor Magenta

# Get site collection objects
Write-Host "   Retrieving site collections"
$spSites = Get-CatSPOAllSites -personal $includePersonalSite

# Loop through site collections for web application
ForEach ($site in $spSites)
{
	# Setting Default Site Type
	$siteType = "SharePoint"
	$siteTitle = ""

	# Screen Output for Site
	Write-Host ("   Scanning site collection: {0}" -f $site.Title)

	# Connecting to Site
	Write-Host ("      Connecting to Site: {0}" -f $site.URL) -ForegroundColor DarkGray
	$pnpConnect = Connect-PnPOnline -Url $site.URL -AccessToken $pnpToken -SkipTenantAdminCheck

	# Collect Site Collection Info
	Write-Host ("      Retreiving Site Information: {0}" -f $site.URL) -ForegroundColor DarkGray
	$pnpSite = Get-PnPSite -Includes RootWeb, Owner, SecondaryContact, RootWeb.LastItemUserModifiedDate, Usage, ID, IsHubSite, HubSiteID, GroupID, RelatedGroupID -Connection $pnpConnect
	if ($pnpSite.RootWeb.Title -eq $null){
		$siteTitle = ""
	}
	else {
		$siteTitle = $pnpSite.RootWeb.Title.Replace("|","\")
	}
	try {
		$tempTeam = Get-Team -GroupID $pnpSite.GroupID
		$isTeam = $true
	}
	catch {
		$isTeam = $false
	}
	if ($isTeam) {
		# Set Site Information
		$siteType = "Teams"
		$siteTitle = $tempTeam.DisplayName.Replace("|","\")
		$mailNickName = $tempTeam.MailNickName.Replace("|","\")

		# Get Member Information
		Write-Host ("      Retrieving Teams Users: {0}" -f $tempTeam.GroupID) -ForegroundColor DarkGray
		$users = Get-TeamUser -GroupID $tempTeam.GroupID

		# Write Users To CSV
		ForEach ($user in $users) {
			Write-Host ("         Writting User Information: {0}" -f $user.User) -ForegroundColor Blue
			Out-File -FilePath $outUsersPath -Append -InputObject ("{0}|{1}|{2}|{3}|{4}" -f 
				$pnpSite.Id,
				$user.UserId, 
				$user.User, 
				$user.Name, 
				$user.Role)
		}
	}
	Else {
		try {
			$tempGroup = Get-AzureADGroup -ObjectId $pnpSite.GroupID
			$isGroup = $true
		}
		catch {
			$isGroup = $false
		}
		if ($isGroup) {
			$siteType = "O365Group"
			$siteTitle = $tempGroup.DisplayName.Replace("|","\")
			$mailNickName = $tempGroup.MailNickName.Replace("|","\")

			# Get Member Information
			Write-Host ("      Retrieving Group Users: {0}" -f $tempGroup.ObjectId) -ForegroundColor DarkGray
			$users = @()
			$users += Get-AzureADGroupOwner -ObjectId $tempGroup.ObjectId
			ForEach($user in $users) {$user.UserType = "Owner"}
			$users += Get-AzureADGroupMember -ObjectId $tempGroup.ObjectId | Where{$users -notcontains $_}

			# Write Users To CSV
			ForEach ($user in $users) {
				Write-Host ("         Writting User Information: {0}" -f $user.UserPrincipalName) -ForegroundColor Blue
				Out-File -FilePath $outUsersPath -Append -InputObject ("{0}|{1}|{2}|{3}|{4}" -f 
					$pnpSite.Id,
					$user.ObjectId, 
					$user.UserPrincipalName, 
					$user.DisplayName, 
					$user.UserType)
			}
		}
	}
	
	# Output SharePoint Users
	If ((-not $isTeam) -and (-not $isGroup)) {
		# Get Member Information
		Write-Host ("      Retrieving SharePoint Users: {0}" -f $pnpSite.ID) -ForegroundColor DarkGray
		$users = Get-PnPUser -WithRightsAssigned -Web $pnpSite.RootWeb

		# Write Users To CSV
		ForEach ($user in $users) {
			Write-Host ("         Writting User Information: {0}" -f $user.LoginName) -ForegroundColor Blue
			If ($user.IsSiteAdmin){
				$userType = "Admin"
			}
			else {
				$userType = $user.TypedObject.Groups.loginName -join ","
			}
			Out-File -FilePath $outUsersPath -Append -InputObject ("{0}|{1}|{2}|{3}|{4}" -f 
				$pnpSite.Id,
				$user.LoginName.Replace("|","\"), 
				$user.Email, 
				$user.Title, 
				$userType)
		}
	}

	# Create Site Record
	Write-Host "      Creating Site Record" -ForegroundColor DarkGray
	Out-File -FilePath $outSitesPath -Append -InputObject ("{0}|{1}|{2}|{3}|{4}|{5}|{6}|{7}|{8}|{9}|{10}|{11}|{12}|{13}|{14}|{15}|{16}|{17}" -f 
		$siteType,
		$siteTitle, 
		$pnpSite.Url, 
		$pnpSite.RootWeb.LastItemUserModifiedDate, 
		$pnpSite.Usage.Storage,
		$pnpSite.ID,
		$pnpSite.RootWeb.Created,
		$pnpSite.Owner.Email,
		$pnpSite.SecondaryContact.Email,
		$site.Template,
		$pnpSite.IsHubSite,
		$pnpSite.HubSiteID,
		$isGroup,
		$pnpSite.GroupID,
		$pnpSite.RelatedGroupID,
		$isTeam,
		$mailNickName,
		$site.PWAEnabled)

	# Disconnect PnP Online
	Write-Host ("      Disconnecting from Site: {0}" -f $site.URL) -ForegroundColor DarkGray
	Disconnect-PnPOnline

	# Site Collection Complete Output
	Write-Host ("      Site collection complete: {0}" -f $site.Title) -ForegroundColor Green
} # ForEach ($site in $spSites)
#endregion

#region Completed Screen Output
# Disconneting from Servcies
Write-Host
Write-Host "Disconnecting from M365 Services" -ForegroundColor Cyan
Write-Host ("   Disconnecting from SharePoint Online: {0}" -f $spAdminURL)
Disconnect-SPOService
Write-Host "   Disconnecting from Microsoft Teams"
Disconnect-MicrosoftTeams
Write-Host "   Disconnecting from Azure AD"
Disconnect-AzureAD

# Finish Script Output
Complete-CatScript $scriptStart[0]

# Ending Transcription
Stop-Transcript
#endregion