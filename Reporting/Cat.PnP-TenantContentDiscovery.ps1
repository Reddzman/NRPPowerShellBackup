<#	Cat.M365-ContentDiscovery.ps1

This script scans a customers M365 tenant generating the following reports.
	-	SharePoint Sites
	-	Microsoft Teams
	-	M365 group connected Sites
	-	User Permission / Group Membership

Dependencies:
	-	PnP.PowerShell Module
	-	Client Azure App Registration and associated Certificate information

Commands
	cd C:\GitHub\CatPowerShell\Reporting\
	.\Cat.PnP-TenantContentDiscovery.ps1 -clientTenant "nrpdev2" -includePersonalSite $true -clientAppID "caf6b11a-cf03-416c-b0f0-67a09a8e30ee" -clientCertPath "C:\GitHub\NRPDev\AppRegistration\PowerShell App Registration.pfx"

Table of Contents
	1.	Input & Output Files
	2.	Variables
	3.	Main Script
	4.	Connecting to M365 services
#>

#region Script Parameters
Param(	[string] $clientTenant = "client",
		[boolean] $includePersonalSite = $true,
		[String] $clientAppID = "{ClientAppRegID}",
		[String] $clientCertPath = "{FullCertificatePath}"
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
Start-Transcript -Path ("..\_Log FIles\{0}-ContentDiscovery-{1}.log" -f $clientTenant,$scriptStart[1])
#endregion

#region 1.	Input & Output Files
# File Paths
$outputBasePath = ("..\_Output Files\{0}\{1}" -f $clientTenant, $scriptStart[1])
$outputErrorsPath = ("{0}\Errors.csv" -f $outputBasePath)
$outputSitesPath = ("{0}\SPOSites.csv" -f $outputBasePath)
$outputGroupsPath = ("{0}\M365Groups.csv" -f $outputBasePath)
$outputTeamsPath = ("{0}\MSTeams.csv" -f $outputBasePath)
$outputUsersPath = ("{0}\UsersReport.csv" -f $outputBasePath)
#endregion

#region 2.	Variables
# Script Variables
$spoTenantUrl = ("https://{0}.sharepoint.com/" -f $clientTenant)
$clientTenantDomain = ("{0}.onmicrosoft.com" -f $clientTenant)
$specialCharacters = "[~``!@#$%^&*()_-]"

# User Input Variables
$clientCertPass = ConvertTo-SecureString -String (Read-Host -Prompt "   Enter certificate password for the app registration") -AsPlainText -Force

# Variable Output
Write-Host
Write-Host "Parameters" -ForegroundColor Cyan
Write-Host ("   Client Tenant: {0}" -f $clientTenant)
Write-Host ("   Include Personal Sites: {0}" -f $includePersonalSite)
Write-Host ("   App ID: {0}" -f $spAdminURL)
Write-Host ("   Certificate: {0}" -f $clientCertPath)
Write-Host
Write-Host "Variables" -ForegroundColor Cyan
Write-Host ("   SPO Tenant URL: {0}" -f $spoTenantUrl)
Write-Host ("   Tenant Domain: {0}" -f $clientTenantDomain)
Write-Host ("   Special Characters to Remove: {0}" -f $specialCharacters)
Write-Host
Write-Host "Reports" -ForegroundColor Cyan
Write-Host ("   Base Directory: {0}" -f $outputSitesPath)
Write-Host ("   Errors: {0}" -f $outputErrorsPath)
Write-Host ("   SPO Sites: {0}" -f $outputSitesPath)
Write-Host ("   Teams: {0}" -f $outputTeamsPath)
Write-Host ("   M365 Groups: {0}" -f $outputGroupsPath)
Write-Host ("   Users: {0}" -f $outputUsersPath)
#endregion

#region 3.	Snapin(s), Modules & Environment Connection(s)
# Load Community / Microsoft Modules
Import-Module PnP.PowerShell

# Load custom modules
Import-Module "..\Reporting\Cat.ReportingFunctions.psm1"
#endregion

#region 4.	Main Script
#region 5.	Creating Reporting Directory
Write-Host
Write-Host ("{0} - Creating Report Directory" -f (Get-Date -Format "yy-MM-dd HH:mm:ss")) -ForegroundColor Cyan
New-Item -Path $outputBasePath -ItemType Directory
#endregion

# Screen Output
#region 6.	Building Site, Team, Group reports
Write-Host
Write-Host "Building Site, Team, Group reports" -ForegroundColor Magenta

#region Retreiving SPO Sites
# Connecting to SPO
Write-Host ("   Connecting to SharePoint Online with PnP PowerShell: {0}" -f $spoTenantUrl)
$cert = Get-PnPAzureCertificate -CertificatePath $clientCertPath -CertificatePassword $clientCertPass
$pnpRootConnect = Connect-PnPOnline $spoTenantUrl -tenant $clientTenantDomain -ClientId $clientAppID -Thumbprint $cert.Thumbprint

# retrieving all Site Collections except redirect sites
Write-Host "   Retrieving site collections"
$spoSites= Get-PnPTenantSite -Detailed -IncludeOneDriveSites:$includePersonalSite | Where-Object Template -ne "REDIRECTSITE#0"

# Disconnecting from SPO
Write-Host "   Disconnecting from SharePoint Online"
Disconnect-PnPOnline -Connection $pnpRootConnect
#endregion

# Looping all site objects
foreach ($site in $spoSites) {
	#region SP Site Details
	Write-Host
	Write-Host ("   {0} - Retreiving Site Information: {1}" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"), $site.URL)
	$pnpSiteConnect = Connect-PnPOnline $site.URL -tenant $clientTenantDomain -ClientId $clientAppID -Thumbprint $cert.Thumbprint
	try {
		$pnpSite = Get-PnPSite -Includes AllowCreateDeclarativeWorkflow, `
		AllowDesigner, `
		AllowExternalEmbeddingWrapper, `
		AllowMasterPageEditing, `
		AllowRevertFromTemplate, `
		AllowSaveDeclarativeWorkflowAsTemplate, `
		AllowSavePublishDeclarativeWorkflow, `
		AllowSelfServiceUpgrade, `
		AllowSelfServiceUpgradeEvaluation, `
		AuditLogTrimmingRetention, `
		CanSyncHubSitePermissions, `
		CanUpgrade, `
		ChannelGroupId, `
		Classification, `
		CommentsOnSitePagesDisabled, `
		CompatibilityLevel, `
		ComplianceAttribute, `
		DisableAppViews, `
		DisableCompanyWideSharingLinks, `
		DisableFlows, `
		ExternalSharingTipsEnabled, `
		ExternalUserExpirationInDays, `
		GeoLocation, `
		GroupId, `
		HubSiteId, `
		HubSiteSynchronizableVisitorGroup, `
		Id, `
		IsHubSite, `
		LockIssue, `
		MaxItemsPerThrottledOperation, `
		NeedsB2BUpgrade, `
		Owner.Email, `
		PrimaryUri, `
		ReadOnly, `
		RelatedGroupId, `
		RequiredDesignerVersion, `
		RootWeb, `
		SandboxedCodeActivationCapability, `
		SearchBoxInNavBar, `
		SearchBoxPlaceholderText, `
		SecondaryContact.Email, `
		SensitivityLabel, `
		ServerRelativeUrl, `
		ShareByEmailEnabled, `
		ShareByLinkEnabled, `
		ShowPeoplePickerSuggestionsForGuestUsers, `
		ShowUrlStructure, `
		SocialBarOnSitePagesDisabled, `
		StatusBarLink, `
		StatusBarText, `
		ThicketSupportDisabled, `
		TrimAuditLog, `
		UIVersionConfigurationEnabled, `
		UpgradeReminderDate, `
		UpgradeScheduled, `
		UpgradeScheduledDate, `
		Upgrading, `
		Url, `
		Usage -ErrorAction Stop
	}
	catch {
		$pnpSite = $null
		Write-Host ("      {0} - Unable to access Site Information: {1}" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"), $site.URL) -ForegroundColor Red
		$obj = New-Object -TypeName PSObject -Property(@{SharePointUrl=$site.URL;`
			errorAction="Retrieving Site Object";`
			errorException=$_.Exception;`
			errorDetails=$_.ErrorDetails;`
			errorStackTrace=$_.ScriptStackTrace;})
		$obj | Export-Csv -Path $outputErrorsPath -NoTypeInformation -Append
	}	
	#endregion
	
	if ($null -ne $pnpSite) {
		#region 6a.	Determine if site is a Team, Group or SPO Site
		$siteType = "SharePoint"

		# Checking for a M365 Group
		try {
			$pnpGroup = Get-PnPAzureADGroup -Identity $pnpSite.GroupID.Guid
			$siteType = "M365 Group"
		}
		catch {$siteType = $siteType}

		# Checking for a Team
		try {
			$pnpTeam = Get-PnPTeamsTeam -Identity $pnpSite.GroupId.Guid
			$siteType = "MS Team"
		}
		catch {$siteType = $siteType}
		#endregion
		
		#region SP Site Record
		Write-Host "      Writing SharePoint site record" -ForegroundColor DarkGray
		$obj = New-Object -TypeName PSObject -Property(@{SiteType=$siteType;`
			Url=$pnpSite.Url;`
			RootWeb_Title=$pnpSite.RootWeb.Title -replace $specialCharacters,'';`
			Id=$pnpSite.Id;`
			AllowCreateDeclarativeWorkflow=$pnpSite.AllowCreateDeclarativeWorkflow;`
			AllowDesigner=$pnpSite.AllowDesigner;`
			AllowExternalEmbeddingWrapper=$pnpSite.AllowExternalEmbeddingWrapper;`
			AllowMasterPageEditing=$pnpSite.AllowMasterPageEditing;`
			AllowRevertFromTemplate=$pnpSite.AllowRevertFromTemplate;`
			AllowSaveDeclarativeWorkflowAsTemplate=$pnpSite.AllowSaveDeclarativeWorkflowAsTemplate;`
			AllowSavePublishDeclarativeWorkflow=$pnpSite.AllowSavePublishDeclarativeWorkflow;`
			AllowSelfServiceUpgrade=$pnpSite.AllowSelfServiceUpgrade;`
			AllowSelfServiceUpgradeEvaluation=$pnpSite.AllowSelfServiceUpgradeEvaluation;`
			AuditLogTrimmingRetention=$pnpSite.AuditLogTrimmingRetention;`
			CanSyncHubSitePermissions=$pnpSite.CanSyncHubSitePermissions;`
			CanUpgrade=$pnpSite.CanUpgrade;`
			ChannelGroupId=$pnpSite.ChannelGroupId;`
			Classification=$pnpSite.Classification;`
			CommentsOnSitePagesDisabled=$pnpSite.CommentsOnSitePagesDisabled;`
			CompatibilityLevel=$pnpSite.CompatibilityLevel;`
			ComplianceAttribute=$pnpSite.ComplianceAttribute;`
			DisableAppViews=$pnpSite.DisableAppViews;`
			DisableCompanyWideSharingLinks=$pnpSite.DisableCompanyWideSharingLinks;`
			DisableFlows=$pnpSite.DisableFlows;`
			ExternalSharingTipsEnabled=$pnpSite.ExternalSharingTipsEnabled;`
			ExternalUserExpirationInDays=$pnpSite.ExternalUserExpirationInDays;`
			GeoLocation=$pnpSite.GeoLocation;`
			GroupId=$pnpSite.GroupId;`
			HubSiteId=$pnpSite.HubSiteId;`
			HubSiteSynchronizableVisitorGroup=$pnpSite.HubSiteSynchronizableVisitorGroup;`
			IsHubSite=$pnpSite.IsHubSite;`
			LockIssue=$pnpSite.LockIssue;`
			MaxItemsPerThrottledOperation=$pnpSite.MaxItemsPerThrottledOperation;`
			NeedsB2BUpgrade=$pnpSite.NeedsB2BUpgrade;`
			Owner_Email=$pnpSite.Owner.Email;`
			PrimaryUri=$pnpSite.PrimaryUri;`
			ReadOnly=$pnpSite.ReadOnly;`
			RelatedGroupId=$pnpSite.RelatedGroupId;`
			RequiredDesignerVersion=$pnpSite.RequiredDesignerVersion;`
			RootWeb_Language=$pnpSite.RootWeb.Language;`
			RootWeb_LastItemModifiedDate=$pnpSite.RootWeb.LastItemModifiedDate;`
			RootWeb_LastItemUserModifiedDate=$pnpSite.RootWeb.LastItemUserModifiedDate;`
			RootWeb_NoCrawl=$pnpSite.RootWeb.NoCrawl;`
			SandboxedCodeActivationCapability=$pnpSite.SandboxedCodeActivationCapability;`
			SearchBoxInNavBar=$pnpSite.SearchBoxInNavBar;`
			SearchBoxPlaceholderText=$pnpSite.SearchBoxPlaceholderText;`
			SecondaryContact_Email=$pnpSite.SecondaryContact.Email;`
			SensitivityLabel_GUID=$pnpSite.SensitivityLabel.GUID;`
			ServerRelativeUrl=$pnpSite.ServerRelativeUrl;`
			ShareByEmailEnabled=$pnpSite.ShareByEmailEnabled;`
			ShareByLinkEnabled=$pnpSite.ShareByLinkEnabled;`
			ShowPeoplePickerSuggestionsForGuestUsers=$pnpSite.ShowPeoplePickerSuggestionsForGuestUsers;`
			ShowUrlStructure=$pnpSite.ShowUrlStructure;`
			SocialBarOnSitePagesDisabled=$pnpSite.SocialBarOnSitePagesDisabled;`
			StatusBarLink=$pnpSite.StatusBarLink;`
			StatusBarText=$pnpSite.StatusBarText;`
			ThicketSupportDisabled=$pnpSite.ThicketSupportDisabled;`
			TrimAuditLog=$pnpSite.TrimAuditLog;`
			UIVersionConfigurationEnabled=$pnpSite.UIVersionConfigurationEnabled;`
			UpgradeReminderDate=$pnpSite.UpgradeReminderDate;`
			UpgradeScheduled=$pnpSite.UpgradeScheduled;`
			UpgradeScheduledDate=$pnpSite.UpgradeScheduledDate;`
			Upgrading=$pnpSite.Upgrading;`
			Usage_Bandwidth=$pnpSite.Usage.Bandwidth;`
			Usage_DiscussionStorage=$pnpSite.Usage.DiscussionStorage;`
			Usage_Hits=$pnpSite.Usage.Hits;`
			Usage_Storage=$pnpSite.Usage.Storage;`
			Usage_StoragePercentageUsed=$pnpSite.Usage.StoragePercentageUsed;`
			Usage_Visits=$pnpSite.Usage.Visits;`
			Usage_TypeId=$pnpSite.Usage.TypeId;})
		$obj | Export-Csv -Path $outputSitesPath -NoTypeInformation -Append
		#endregion

		switch ($siteType) {
			"SharePoint" {
				#region SharePoint Users
				Write-Host "      Writing SharePoint User records" -ForegroundColor DarkGray
				$spGroups = @()
				$pnpUsers = @()
				$spGroups = Get-PnPGroup
				foreach ($group in $spGroups) {
					$spGroupPermissions = Get-PnPGroupPermissions -Identity $group
					$spPermission = $spGroupPermissions.Name -join ";"
					$spUsers = Get-PnPGroupMember -Group $group
					$spUsers | Add-Member -NotePropertyName "GroupTitle" -NotePropertyValue $group.Title
					$spUsers | Add-Member -NotePropertyName "GroupId" -NotePropertyValue $group.Id
					$spUsers | Add-Member -NotePropertyName "GroupPermissions" -NotePropertyValue $spPermission
					$pnpUsers += $spUsers
				}
				foreach ($user in $pnpUsers) {
					$obj = New-Object -TypeName PSObject -Property(@{SharePointUrl=$site.URL;`
						Group=$user.GroupTitle -replace $specialCharacters,'';`
						GroupId=$user.GroupId;`
						UserId=$user.Id;`
						DisplayName=$user.Title -replace $specialCharacters,'';`
						UserPrincipalName=$user.UserPrincipalName;`
						UserPermissions=$user.GroupPermissions;})
					$obj | Export-Csv -Path $outputUsersPath -NoTypeInformation -Append
				}	
				#endregion
			}
			{($_ -eq "M365 Group") -or ($_ -eq "MS Team")} {
				#region M365 Group Record
				Write-Host "      Writing M365 Group record" -ForegroundColor DarkGray
				$obj = New-Object -TypeName PSObject -Property(@{SharePointUrl=$site.URL;`
					Description=(($pnpGroup.Description -replace $specialCharacters,'') -replace "`n",' ') -replace "`r",' ';`
					DisplayName=$pnpGroup.DisplayName -replace $specialCharacters,'';`
					GroupTypes=$pnpGroup.GroupTypes;`
					GroupId=$pnpGroup.Id;`
					Mail=$pnpGroup.Mail;`
					MailEnabled=$pnpGroup.MailEnabled;`
					MailNickname=$pnpGroup.MailNickname;`
					SecurityEnabled=$pnpGroup.SecurityEnabled;})
				$obj | Export-Csv -Path $outputGroupsPath -NoTypeInformation -Append
				#endregion
			}
			"M365 Group" {
				#region M365 Group Users Records
				Write-Host "      Writing M365 Group User records" -ForegroundColor DarkGray
				$pnpUsers = @()
				$pnpUsers += Get-PnPAzureADGroupOwner -Identity $pnpGroup.Id
				$pnpUsers | Add-Member -NotePropertyName "UserType" -NotePropertyValue "Owner"
				$pnpUsers += Get-PnPAzureADGroupMember -Identity $pnpGroup.Id | Where{$pnpUsers -notcontains $_}
				foreach ($user in $pnpUsers) {
					$obj = New-Object -TypeName PSObject -Property(@{SharePointUrl=$site.URL;`
						Group=$pnpGroup.DisplayName -replace $specialCharacters,'';`
						GroupId=$pnpGroup.Id;`
						UserId=$user.Id;`
						DisplayName=$user.DisplayName -replace $specialCharacters,'';`
						UserPrincipalName=$user.UserPrincipalName;`
						UserPermissions=If($user.UserType -ne "Owner"){"Member"}else {$user.UserType};})
					$obj | Export-Csv -Path $outputUsersPath -NoTypeInformation -Append
				}
				#endregion
			}
			"MS Team" {
				#region MS Teams Record
				Write-Host "      Writing Microsoft Teams record" -ForegroundColor DarkGray
				$obj = New-Object -TypeName PSObject -Property(@{SharePointURL=$site.URL;`
					GroupId=$pnpTeam.GroupId;`
					DisplayName=$pnpTeam.DisplayName -replace $specialCharacters,'';`
					Classification=$pnpTeam.Classification;`
					FunSettings_AllowGiphy=$pnpTeam.FunSettings.AllowGiphy;`
					FunSettings_GiphyContentRating=$pnpTeam.FunSettings.GiphyContentRating;`
					FunSettings_AllowStickersAndMemes=$pnpTeam.FunSettings.AllowStickersAndMemes;`
					FunSettings_AllowCustomMemes=$pnpTeam.FunSettings.AllowCustomMemes;`
					GuestSettings_AllowCreateUpdateChannels=$pnpTeam.GuestSettings.AllowCreateUpdateChannels;`
					GuestSettings_AllowDeleteChannels=$pnpTeam.GuestSettings.AllowDeleteChannels;`
					MemberSettings_AllowCreateUpdateChannels=$pnpTeam.MemberSettings.AllowCreateUpdateChannels;`
					MemberSettings_AllowDeleteChannels=$pnpTeam.MemberSettings.AllowDeleteChannels;`
					MemberSettings_AllowAddRemoveApps=$pnpTeam.MemberSettings.AllowAddRemoveApps;`
					MemberSettings_AllowCreateUpdateRemoveTabs=$pnpTeam.MemberSettings.AllowCreateUpdateRemoveTabs;`
					MemberSettings_AllowCreateUpdateRemoveConnectors=$pnpTeam.MemberSettings.AllowCreateUpdateRemoveConnectors;`
					MemberSettings_AllowCreatePrivateChannels=$pnpTeam.MemberSettings.AllowCreatePrivateChannels;`
					MessagingSettings_AllowUserEditMessages=$pnpTeam.MessagingSettings.AllowUserEditMessages;`
					MessagingSettings_AllowUserDeleteMessages=$pnpTeam.MessagingSettings.AllowUserDeleteMessages;`
					MessagingSettings_AllowOwnerDeleteMessages=$pnpTeam.MessagingSettings.AllowOwnerDeleteMessages;`
					MessagingSettings_AllowTeamMentions=$pnpTeam.MessagingSettings.AllowTeamMentions;`
					MessagingSettings_AllowChannelMentions=$pnpTeam.MessagingSettings.AllowChannelMentions;`
					DiscoverySettings_ShowInTeamsSearchAndSuggestions=$pnpTeam.DiscoverySettings.ShowInTeamsSearchAndSuggestions;`
					Security=$pnpTeam.Security;`
					Specialization=$pnpTeam.Specialization;`
					CloneFrom=$pnpTeam.CloneFrom;`
					IsArchived=$pnpTeam.IsArchived;`
					MailNickname=$pnpTeam.MailNickname;`
					Description=(($pnpTeam.Description -replace $specialCharacters,'') -replace "`n",' ') -replace "`r",' ';`
					Visibility=$pnpTeam.Visibility;})
				$obj | Export-Csv -Path $outputTeamsPath -NoTypeInformation -Append
				#endregion

				#region Teams Users Records
				Write-Host "      Writing MS Teams User records" -ForegroundColor DarkGray
				$pnpUsers = @()
				$pnpUsers = Get-PnPTeamsUser -Team $pnpTeam
				foreach ($user in $pnpUsers) {
					$obj = New-Object -TypeName PSObject -Property(@{SharePointUrl=$site.URL;`
						Group=$pnpTeam.DisplayName -replace $specialCharacters,'';`
						GroupId=$pnpTeam.GroupID;`
						UserId=$user.Id;`
						DisplayName=$user.DisplayName -replace $specialCharacters,'';`
						UserPrincipalName=$user.UserPrincipalName;`
						UserPermissions=$user.UserType;})
					$obj | Export-Csv -Path $outputUsersPath -NoTypeInformation -Append
				}	
				#endregion
			}
		}
	}
}
#endregion

#endregion

#region Completed Screen Output
# Finish Script Output
Complete-CatScript $scriptStart[0]

# Ending Transcription
Stop-Transcript
#endregion