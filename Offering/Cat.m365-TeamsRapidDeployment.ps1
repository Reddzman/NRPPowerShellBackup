<#  Cat.MT-DeployCrisisTeam.ps1

This script deploys the configurations for remote collaboration in a crisis scenario for Microsoft Teams (MT).  The follow configurations
are added to the target tenant based on an input XML file.
	-	Creates Crisis Communication Team for the crisis communication PowerApp
	-	Creates a Organization Wide Team for remote collaboration
	-	Creates teams channels for departments/groups
	-	Configures O365 Tenant based on Best Practices and client discussions

Dependencies
	-	Microsoft Teams Module
	-	PnP PowerShell - https://docs.microsoft.com/en-us/powershell/sharepoint/sharepoint-pnp/sharepoint-pnp-cmdlets?view=sharepoint-ps
	-	SharePoint Online Module
	-	SkypeOnlineConnector Module - https://www.microsoft.com/en-us/download/details.aspx?id=39366
	-	ExchangeOnlineManagement - https://www.powershellgallery.com/packages/ExchangeOnlineManagement/0.4368.1
	-	AzureADPreview Module
	-	General Functions - ..\General\Cat.PnP-GeneralFunctionsFunctions.psm1
	-	Input XML File - Example:  ..\_Input Files\Testing\Cat.MT-CrisisTeam-Test.xml

Call script with parameters
	.\Cat.MT-DeployCrisisTeam.ps1 -inputFile Clients\Cat.MT-DeployCrisisTeam-Client.xml -client "ClientNameNoSpacesOrAbriviation"

Table of Contents
	1.	Input & Output Files
	2.	Variables
	3.	Snapin(s), Modules & Environment Connection(s)
	4.	Main Script
#>

#region Script Parameters
Param(	[string] $inputFile = "Testing\Cat.m365-TeamsRapidDeployment-Test.xml",
		[string] $client = "Testing"
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
Start-Transcript -Path ("..\_Log FIles\{0}.m365-TeamsRapidDeployment-{1}.log" -f $client, $scriptStart[1])
#endregion

#region 1.	Input & Output Files
# File Paths
$inPath = ("..\_Input Files\{0}" -f $inputFile)

# Import Input File(s)
# Load CSV file
If (Test-Path ($inPath)){
	[xml]$inFile = Get-Content $inPath 
}
else {
	Write-Host
	throw ("The input file does not exist: {0}" -f $inPath)	
}
#endregion

#region 2.	Variables
# User Inputs
Write-Host
Write-Host "User Input Variables" -ForegroundColor Cyan
Write-Host "   Enter credentials for O365 Global Admin:"
$o365Cred = Get-Credential -Message "Enter your credentials for O365 Global Admin:" 

# Load Settings from XML
# Script Variables
$pause = 60
$azureADExernalAllowBlockPolicyDefinition = @("{`"B2BManagementPolicy`":{`"InvitationsAllowedAndBlockedDomainsPolicy`":{`"AllowedDomains`": [{ClientAllowDomains}],`"BlockedDomains`": [{ClientBlockDomains}]}}}")

#   O365 Settings
$o365SettingTeamsAllowGuestUsers = If ($inFile.CrisisTeamDeployment.TenantSettings.TeamsAllowGuestUsers.ToLower() -eq "true") {$true} else {$false}
$o365SettingTeamsAllowEgnyteFiles = If ($inFile.CrisisTeamDeployment.TenantSettings.TeamsAllowEgnyteFiles.ToLower() -eq "true") {$true} else {$false}
$o365SettingTeamsAllowDropBoxFiles = If ($inFile.CrisisTeamDeployment.TenantSettings.TeamsAllowDropBoxFiles.ToLower() -eq "true") {$true} else {$false}
$o365SettingTeamsAllowBoxFiles = If ($inFile.CrisisTeamDeployment.TenantSettings.TeamsAllowBoxFiles.ToLower() -eq "true") {$true} else {$false}
$o365SettingTeamsAllowGoogleFiles = If ($inFile.CrisisTeamDeployment.TenantSettings.TeamsAllowGoogleFiles.ToLower() -eq "true") {$true} else {$false}
$o365SettingTeamsAllowSkypeForBusiness = If ($inFile.CrisisTeamDeployment.TenantSettings.TeamsAllowSkypeForBusiness.ToLower() -eq "true") {$true} else {$false}
# This capability isn't displaying properly in the UI yet.  Documenation says it is preview  $o365SettingTeamsAllowSkypeForConsumers = If ($inFile.CrisisTeamDeployment.TenantSettings.TeamsAllowSkypeForConsumers.ToLower() -eq "true") {$true} else {$false}
$o365SettingTeamsAllowBroadcastScheduling = If ($inFile.CrisisTeamDeployment.TenantSettings.TeamsAllowBroadcastScheduling.ToLower() -eq "true") {$true} else {$false}
$domains = $inFile.CrisisTeamDeployment.TenantSettings.AzureADExternalDomains.Domain -join '","'
$domains = '"{0}"' -f $domains
if ($inFile.CrisisTeamDeployment.TenantSettings.AzureADExternalAllowBlockList.ToLower() -eq "allow") {
	$azureADExernalAllowBlockPolicyDefinition = $azureADExernalAllowBlockPolicyDefinition.Replace("{ClientAllowDomains}",$domains).Replace("{ClientBlockDomains}","")
}
Else {
	$azureADExernalAllowBlockPolicyDefinition = $azureADExernalAllowBlockPolicyDefinition.Replace("{ClientAllowDomains}","").Replace("{ClientBlockDomains}",$domains)
}
$o365SettingAzureADAllowGroupCreation = If ($inFile.CrisisTeamDeployment.TenantSettings.O365GroupCreators.AllowGroupCreation.ToLower() -eq "true") {$true} else {$false}

#	Team Workspace Settings
$tSettingAllowGiphy = If ($inFile.CrisisTeamDeployment.TeamSettings.AllowGiphy.ToLower() -eq "true") {$true} else {$false}
$tSettingGiphyContentRating = $inFile.CrisisTeamDeployment.TeamSettings.GiphyContentRating
$tSettingAllowStickersAndMemes = If ($inFile.CrisisTeamDeployment.TeamSettings.AllowStickersAndMemes.ToLower() -eq "true") {$true} else {$false}
$tSettingAllowCustomMemes = If ($inFile.CrisisTeamDeployment.TeamSettings.AllowCustomMemes.ToLower() -eq "true") {$true} else {$false}
$tSettingAllowGuestCreateUpdateChannels = If ($inFile.CrisisTeamDeployment.TeamSettings.AllowGuestCreateUpdateChannels.ToLower() -eq "true") {$true} else {$false}
$tSettingAllowGuestDeleteChannels = If ($inFile.CrisisTeamDeployment.TeamSettings.AllowGuestDeleteChannels.ToLower() -eq "true") {$true} else {$false}
$tSettingAllowCreateUpdateChannels = If ($inFile.CrisisTeamDeployment.TeamSettings.AllowCreateUpdateChannels.ToLower() -eq "true") {$true} else {$false}
$tSettingAllowDeleteChannels = If ($inFile.CrisisTeamDeployment.TeamSettings.AllowDeleteChannels.ToLower() -eq "true") {$true} else {$false}
$tSettingAllowAddRemoveApps = If ($inFile.CrisisTeamDeployment.TeamSettings.AllowAddRemoveApps.ToLower() -eq "true") {$true} else {$false}
$tSettingAllowCreateUpdateRemoveTabs = If ($inFile.CrisisTeamDeployment.TeamSettings.AllowCreateUpdateRemoveTabs.ToLower() -eq "true") {$true} else {$false}
$tSettingAllowCreateUpdateRemoveConnectors = If ($inFile.CrisisTeamDeployment.TeamSettings.AllowCreateUpdateRemoveConnectors.ToLower() -eq "true") {$true} else {$false}
$tSettingAllowUserEditMessages = If ($inFile.CrisisTeamDeployment.TeamSettings.AllowUserEditMessages.ToLower() -eq "true") {$true} else {$false}
$tSettingAllowUserDeleteMessages = If ($inFile.CrisisTeamDeployment.TeamSettings.AllowUserDeleteMessages.ToLower() -eq "true") {$true} else {$false}
# *** There appears to be a bug in this setting via PowerShell *** $tSettingAllowOwnerDeleteMessages = If ($inFile.CrisisTeamDeployment.TeamSettings.AllowOwnerDeleteMessages.ToLower() -eq "true") {$true} else {$false}
$tSettingAllowTeamMentions = If ($inFile.CrisisTeamDeployment.TeamSettings.AllowTeamMentions.ToLower() -eq "true") {$true} else {$false}
$tSettingAllowChannelMentions = If ($inFile.CrisisTeamDeployment.TeamSettings.AllowChannelMentions.ToLower() -eq "true") {$true} else {$false}
# *** There appears to be a bug in this setting via PowerShell *** $tSettingShowInTeamsSearchAndSuggestions = If ($inFile.CrisisTeamDeployment.TeamSettings.ShowInTeamsSearchAndSuggestions.ToLower() -eq "true") {$true} else {$false}
$tSettingRetainCreatedGroup = If ($inFile.CrisisTeamDeployment.TeamSettings.RetainCreatedGroup.ToLower() -eq "true") {$true} else {$false}

# Output Parameters & Variables
Write-Host
Write-Host "Script Parameters" -ForegroundColor Cyan
Write-Host ("   Input File: {0}" -f $inputFile)
Write-Host ("   Client: {0}" -f $client)
Write-Host 
Write-Host ("Office 365 Settings") -ForegroundColor Cyan
Write-Host ("   SPO Admin URL: {0}" -f $inFile.CrisisTeamDeployment.TenantSettings.SPOAdminURL)
Write-Host ("   SPO External Sharing: {0}" -f $inFile.CrisisTeamDeployment.TenantSettings.SPOExternalSharing)
Write-Host ("   Teams Allow Guest Users - XML: {0} Variable: {1}" -f $inFile.CrisisTeamDeployment.TenantSettings.TeamsAllowGuestUsers, $o365SettingTeamsAllowGuestUsers)
Write-Host ("   Teams Egnyte Files - XML: {0} Variable: {1}" -f $inFile.CrisisTeamDeployment.TenantSettings.TeamsAllowEgnyteFiles, $o365SettingTeamsAllowEgnyteFiles)
Write-Host ("   Teams DropBox Files - XML: {0} Variable: {1}" -f $inFile.CrisisTeamDeployment.TenantSettings.TeamsAllowDropBoxFiles, $o365SettingTeamsAllowDropBoxFiles)
Write-Host ("   Teams Box Files - XML: {0} Variable: {1}" -f $inFile.CrisisTeamDeployment.TenantSettings.TeamsAllowBoxFiles, $o365SettingTeamsAllowBoxFiles)
Write-Host ("   Teams Communicate with Skype for Business - XML: {0} Variable: {1}" -f $inFile.CrisisTeamDeployment.TenantSettings.TeamsAllowSkypeForBusiness, $o365SettingTeamsAllowSkypeForBusiness)
# This capability isn't displaying properly in the UI yet.  Documenation says it is preview  Write-Host ("   Teams Communicate with Skype for Consumers - XML: {0} Variable: {1}" -f $inFile.CrisisTeamDeployment.TenantSettings.TeamsAllowSkypeForConsumers, $o365SettingTeamsAllowSkypeForConsumers)
Write-Host ("   Teams Auto Admit Users: {0}" -f $inFile.CrisisTeamDeployment.TenantSettings.TeamsAutoAdmitUsers)
Write-Host ("   Teams Allow Broadcast Scheduling - XML: {0} Variable: {1}" -f $inFile.CrisisTeamDeployment.TenantSettings.TeamsAllowBroadcastScheduling, $o365SettingTeamsAllowBroadcastScheduling)
Write-Host ("   Azure AD External Allow Block List: {0}" -f $inFile.CrisisTeamDeployment.TenantSettings.AzureADExternalAllowBlockList)
Write-Host ("   Azure AD External Policy Name: {0}" -f $inFile.CrisisTeamDeployment.TenantSettings.AzureADExternalPolicyName)
Write-Host ("   Azure AD Site/Group Creation Group: {0}" -f $inFile.CrisisTeamDeployment.TenantSettings.O365GroupCreators.DisplayName)
Write-Host ("   Azure AD Allow Group Creation - XML: {0} Variable: {1}" -f $inFile.CrisisTeamDeployment.TenantSettings.O365GroupCreators.AllowGroupCreation, $o365SettingAzureADAllowGroupCreation)
Write-Host
Write-Host ("Team Workspace Settings") -ForegroundColor Cyan
Write-Host ("   Allow Giphy - XML: {0} Variable: {1}" -f $inFile.CrisisTeamDeployment.TeamSettings.AllowGiphy, $tSettingAllowGiphy)
Write-Host ("   Giphy Content Rating - XML: {0} Variable: {1}" -f $inFile.CrisisTeamDeployment.TeamSettings.GiphyContentRating, $tSettingGiphyContentRating)
Write-Host ("   Allow Sickers And Memes - XML: {0} Variable: {1}" -f $inFile.CrisisTeamDeployment.TeamSettings.AllowStickersAndMemes, $tSettingAllowStickersAndMemes)
Write-Host ("   Allow Custom Memes - XML: {0} Variable: {1}" -f $inFile.CrisisTeamDeployment.TeamSettings.AllowCustomMemes, $tSettingAllowCustomMemes)
Write-Host ("   Allow Guest Create update Channels - XML: {0} Variable: {1}" -f $inFile.CrisisTeamDeployment.TeamSettings.AllowGuestCreateUpdateChannels, $tSettingAllowGuestCreateUpdateChannels)
Write-Host ("   Allow Guest Delete Channels - XML: {0} Variable: {1}" -f $inFile.CrisisTeamDeployment.TeamSettings.AllowGuestDeleteChannels, $tSettingAllowGuestDeleteChannels)
Write-Host ("   Allow Create Update Channels - XML: {0} Variable: {1}" -f $inFile.CrisisTeamDeployment.TeamSettings.AllowCreateUpdateChannels, $tSettingAllowCreateUpdateChannels)
Write-Host ("   Allow Delete Channels - XML: {0} Variable: {1}" -f $inFile.CrisisTeamDeployment.TeamSettings.AllowDeleteChannels, $tSettingAllowDeleteChannels)
Write-Host ("   Allow Add Remove Apps - XML: {0} Variable: {1}" -f $inFile.CrisisTeamDeployment.TeamSettings.AllowAddRemoveApps, $tSettingAllowAddRemoveApps)
Write-Host ("   Allow Create Update Remove Tabs - XML: {0} Variable: {1}" -f $inFile.CrisisTeamDeployment.TeamSettings.AllowCreateUpdateRemoveTabs, $tSettingAllowCreateUpdateRemoveTabs)
Write-Host ("   Allow Create Update Remove Connectors - XML: {0} Variable: {1}" -f $inFile.CrisisTeamDeployment.TeamSettings.AllowCreateUpdateRemoveConnectors, $tSettingAllowCreateUpdateRemoveConnectors)
Write-Host ("   Allow User Edit Messages - XML: {0} Variable: {1}" -f $inFile.CrisisTeamDeployment.TeamSettings.AllowUserEditMessages, $tSettingAllowUserEditMessages)
Write-Host ("   Allow User Delete Messages - XML: {0} Variable: {1}" -f $inFile.CrisisTeamDeployment.TeamSettings.AllowUserDeleteMessages, $tSettingAllowUserDeleteMessages)
# *** There appears to be a bug in this setting via PowerShell *** Write-Host ("   Allow Owner Delete Messages - XML: {0} Variable: {1}" -f $inFile.CrisisTeamDeployment.TeamSettings.AllowOwnerDeleteMessages, $tSettingAllowOwnerDeleteMessages)
Write-Host ("   Allow Team Mentions - XML: {0} Variable: {1}" -f $inFile.CrisisTeamDeployment.TeamSettings.AllowTeamMentions, $tSettingAllowTeamMentions)
Write-Host ("   Allow Channel Mentions - XML: {0} Variable: {1}" -f $inFile.CrisisTeamDeployment.TeamSettings.AllowChannelMentions, $tSettingAllowChannelMentions)
# *** There appears to be a bug in this setting via PowerShell *** Write-Host ("   Show In Teams Search And Suggestions - XML: {0} Variable: {1}" -f $inFile.CrisisTeamDeployment.TeamSettings.ShowInTeamsSearchAndSuggestions, $tSettingShowInTeamsSearchAndSuggestions)
Write-Host ("   Retain Created Group - XML: {0} Variable: {1}" -f $inFile.CrisisTeamDeployment.TeamSettings.RetainCreatedGroup, $tSettingRetainCreatedGroup)
Write-Host
Write-Host "User Provided Variables" -ForegroundColor Cyan
Write-Host ("   Credentials: {0}" -f $o365Cred.UserName)
Write-Host
Write-Host "Script Variables" -ForegroundColor Cyan
Write-Host ("   Pause (Seconds): {0}" -f $pause)
Write-Host
Write-Host "Azure AD External Policy Definitiion" -ForegroundColor Cyan
$azureADExernalAllowBlockPolicyDefinition
#endregion

#region 3.	Snapin(s), Modules & Environment Connection(s)
# Load MS / PnP PowerShell Modules
Import-module MicrosoftTeams #-erroraction SilentlyContinue
#endregion
    
#region 4.	Main Script
# Screen Output
Write-Host
Write-Host ("Configuring Tenant Settings: {0}" -f $inputFile) -ForegroundColor Magenta

# Connecting to SPO
Write-Host ("   Connecting to SPO Tenant: {0}" -f $inFile.CrisisTeamDeployment.TenantSettings.SPOAdminURL)
Connect-SPOService -Url $inFile.CrisisTeamDeployment.TenantSettings.SPOAdminURL -Credential $o365Cred

# Setting External Sharing
Write-Host ("      Configuring SPO External Sharing: {0}" -f $inFile.CrisisTeamDeployment.TenantSettings.SPOExternalSharing) -ForegroundColor DarkGray
Set-SPOTenant -SharingCapability $inFile.CrisisTeamDeployment.TenantSettings.SPOExternalSharing

# Disconnecting SharePoint Session
Write-Host "      Disconnecting from SharePoint Online" -ForegroundColor DarkGray
Disconnect-SPOService #>

# Connecting to Skype Online
Write-Host ("   Connecting to Skype Online")
$skypeSession = New-CsOnlineSession -Credential $o365Cred
$skypeActiveSession = Import-PSSession -Session $skypeSession -AllowClobber

# Updating Guest Access and External Files
Write-Host ("      Configuring Teams Guest Access and External Files") -ForegroundColor DarkGray
$tConfig = Get-CsTeamsClientConfiguration
Set-CsTeamsClientConfiguration -Identity $tConfig.Identity `
	-AllowGuestUser $o365SettingTeamsAllowGuestUsers `
	-AllowDropBox $o365SettingTeamsAllowDropBoxFiles `
	-AllowEgnyte $o365SettingTeamsAllowEgnyteFiles `
	-AllowBox $o365SettingTeamsAllowBoxFiles `
	-AllowGoogleDrive $o365SettingTeamsAllowGoogleFiles

# Updating Skype Communication settings
Write-Host ("      Configuring Teams External Access - Skype for Business: {0}; Skype for Consumers: {1}" -f $inFile.CrisisTeamDeployment.TenantSettings.TeamsAllowSkypeForBusiness, $inFile.CrisisTeamDeployment.TenantSettings.TeamsAllowSkypeForConsumers) -ForegroundColor DarkGray
Set-CsTenantFederationConfiguration -Tenant $inFile.CrisisTeamDeployment.TenantSettings.TenantID -AllowFederatedUsers $o365SettingTeamsAllowSkypeForBusiness # This capability isn't displaying properly in the UI yet.  Documenation says it is preview  -AllowPublicUsers $o365SettingTeamsAllowSkypeForConsumers

# Updating Meeting Policy settings
Write-Host ("      Configuring Teams Meeting Policy: {0}" -f $inFile.CrisisTeamDeployment.TenantSettings.TeamsAutoAdmitUsers) -ForegroundColor DarkGray
Set-CsTeamsMeetingPolicy -AutoAdmittedUsers $inFile.CrisisTeamDeployment.TenantSettings.TeamsAutoAdmitUsers

# Updating Live Event Schduling settings
Write-Host ("      Configuring Live Event Scheduling: {0}" -f $o365SettingTeamsAllowBroadcastScheduling) -ForegroundColor DarkGray
Set-CsTeamsMeetingBroadcastPolicy -AllowBroadcastScheduling $o365SettingTeamsAllowBroadcastScheduling

# Disconnecting Skype Session
Write-Host "      Disconnecting from Skype Online" -ForegroundColor DarkGray
Remove-PSSession -Session $skypeSession #>

# Connecting to Exchange Online
Write-Host ("   Connecting to Exchange Online")
Connect-ExchangeOnline -Credential $o365Cred -ShowProgress $false

# Creating group for Site/Group creation
$tempGroup = Get-UnifiedGroup | Where{$_.DisplayName -eq $inFile.CrisisTeamDeployment.TenantSettings.O365GroupCreators.DisplayName}
If ($null -eq $tempGroup.Id) {
	Write-Host ("      Creating Office 365 Group for site creation: {0}" -f $inFile.CrisisTeamDeployment.TenantSettings.O365GroupCreators.Alias) -ForegroundColor DarkGray
	$groupOwners = $inFile.CrisisTeamDeployment.TenantSettings.O365GroupCreators.Owners.Owner
	$groupOwnersCount = $groupOwners.Count
	$groupOwner = If($groupOwnersCount -eq 1) {$inFile.CrisisTeamDeployment.TenantSettings.O365GroupCreators.Owners.Owner} Else {$inFile.CrisisTeamDeployment.TenantSettings.O365GroupCreators.Owners.Owner[0]}
	$groupMembers = $inFile.CrisisTeamDeployment.TenantSettings.O365GroupCreators.Members.Member -join ","
	$tempGroup = New-UnifiedGroup -DisplayName $inFile.CrisisTeamDeployment.TenantSettings.O365GroupCreators.DisplayName `
		-Alias $inFile.CrisisTeamDeployment.TenantSettings.O365GroupCreators.Alias `
		-Owner $groupOwner
}

# Updating Site/Group creation Owners
Write-Host ("      Updating Group Owners: {0}" -f $inFile.CrisisTeamDeployment.TenantSettings.O365GroupCreators.Alias) -ForegroundColor DarkGray
foreach ($user in $inFile.CrisisTeamDeployment.TenantSettings.O365GroupCreators.Owners.Owner){
	Write-Host ("         Adding User: {0}" -f $user) -ForegroundColor Blue
	Add-UnifiedGroupLinks -Identity $tempGroup.Id -LinkType Member -Links $user
	Add-UnifiedGroupLinks -Identity $tempGroup.Id -LinkType Owner -Links $user
}

# Updating Site/Group creation Members
Write-Host ("      Updating Group Members: {0}" -f $inFile.CrisisTeamDeployment.TenantSettings.O365GroupCreators.Alias) -ForegroundColor DarkGray
foreach ($user in $inFile.CrisisTeamDeployment.TenantSettings.O365GroupCreators.Members.Member){
	Write-Host ("         Adding User: {0}" -f $user) -ForegroundColor Blue
	Add-UnifiedGroupLinks -Identity $tempGroup.Id -LinkType Member -Links $user
}

# Disconnecting Exchange Session
Write-Host "      Disconnecting from Exchange Online" -ForegroundColor DarkGray
$exchangeSession = Get-PSSession | Where{$_.Name -like "ExchangeOnline*"}
Remove-PSSession -Session $exchangeSession #>

# Connecting to Azure AD
Write-Host ("   Connecting to Azure AD")
$aadConnection = Connect-AzureAD -Credential $o365Cred

# Configuring Group Creation Permissions
# Check if settings already exists
$groupCreationSetting = Get-AzureADDirectorySetting | Where{$_.DisplayName -eq "Group.Unified"}
if($null -eq $groupCreationSetting.Id)
{
	Write-Host "      Creating Unified Group Settings" -ForegroundColor DarkGray
	$groupCreationSettingTemplate = Get-AzureADDirectorySettingTemplate | Where-object {$_.displayname -eq "group.unified"}
    $groupCreationSettingCopy = $groupCreationSettingTemplate.CreateDirectorySetting()
    $groupCreationSetting = New-AzureADDirectorySetting -DirectorySetting $groupCreationSettingCopy
}

# Update settings
Write-Host ("      Updating Unified Group Settings for Site Creation: {0}"  -f $groupCreationSetting.Id) -ForegroundColor DarkGray
$groupCreationSettingCopy = Get-AzureADDirectorySetting -Id $groupCreationSetting.Id
$groupCreationSettingCopy["EnableGroupCreation"] = $o365SettingAzureADAllowGroupCreation
$groupCreationSettingGroup = Get-AzureADGroup -SearchString $inFile.CrisisTeamDeployment.TenantSettings.O365GroupCreators.DisplayName
$groupCreationSettingCopy["GroupCreationAllowedGroupId"] = $groupCreationSettingGroup.Id
Set-AzureADDirectorySetting -Id $groupCreationSetting.Id -DirectorySetting $groupCreationSettingCopy

# Configuring Azure External Collaboration Allow List
$tempADPolicy = Get-AzureADPolicy | Where{$_.DisplayName -eq $inFile.CrisisTeamDeployment.TenantSettings.AzureADExternalPolicyName}
If ($null -eq $tempADPolicy.Id) {
	Write-Host ("      Creating Azure AD External Collaboration Policies: {0}" -f $inFile.CrisisTeamDeployment.TenantSettings.AzureADExternalAllowBlockList) -ForegroundColor DarkGray
	$tempADPolicy = New-AzureADPolicy -Definition $azureADExernalAllowBlockPolicyDefinition -DisplayName $inFile.CrisisTeamDeployment.TenantSettings.AzureADExternalPolicyName -Type B2BManagementPolicy -IsOrganizationDefault $true
}
else {
	Write-Host ("      Configuring Azure AD External Collaboration Policies: {0}" -f $inFile.CrisisTeamDeployment.TenantSettings.AzureADExternalAllowBlockList) -ForegroundColor DarkGray
	Set-AzureADPolicy -Id $tempADPolicy.Id -Definition $azureADExernalAllowBlockPolicyDefinition
}

# Disconnecting AzureAD
Write-Host "      Disconnecting from Azure AD" -ForegroundColor DarkGray
Disconnect-AzureAD #>

# Screen Output
Write-Host "   Tenant Settings Configuration Complete" -ForegroundColor Green

# Screen Output
Write-Host
Write-Host ("Configuring Microsoft Teams: {0}" -f $inputFile) -ForegroundColor Magenta

# Conneting to Microsoft Teams
Write-Host ("   Connecting to Microsoft Teams: {0}" -f $o365Cred.UserName)
$tConnect = Connect-MicrosoftTeams -Credential $o365Cred

# Looping through XML for teams workspaces
Write-Host "   Creating Team Workspaces"
ForEach ($workspace in $inFile.CrisisTeamDeployment.CreateTeams.Team){
	Write-Host ("      Creating Workspace: {0}" -f $workspace.DisplayName) -ForegroundColor DarkGray
	$tOwners = $workspace.Owners.Owner
	$tOwnersCount = $tOwners.Count
	$tOwner = If($tOwnersCount -eq 1) {$workspace.Owners.Owner} Else {$workspace.Owners.Owner[0]}
	$team = New-Team -DisplayName $workspace.DisplayName `
		-Description $workspace.Description `
		-MailNickName $workspace.MailNickName `
		-Visibility $workspace.Visibility `
		-Owner $tOwner `
		-AllowGiphy $tSettingAllowGiphy `
		-GiphyContentRating $tSettingGiphyContentRating `
		-AllowStickersAndMemes $tSettingAllowStickersAndMemes `
		-AllowCustomMemes $tSettingAllowCustomMemes `
		-AllowGuestCreateUpdateChannels $tSettingAllowGuestCreateUpdateChannels `
		-AllowGuestDeleteChannels $tSettingAllowGuestDeleteChannels `
		-AllowCreateUpdateChannels $tSettingAllowCreateUpdateChannels `
		-AllowDeleteChannels $tSettingAllowDeleteChannels `
		-AllowAddRemoveApps $tSettingAllowAddRemoveApps `
		-AllowCreateUpdateRemoveTabs $tSettingAllowCreateUpdateRemoveTabs `
		-AllowCreateUpdateRemoveConnectors $tSettingAllowCreateUpdateRemoveConnectors `
		-AllowUserEditMessages $tSettingAllowUserEditMessages `
		-AllowUserDeleteMessages $tSettingAllowUserDeleteMessages `
		<# *** There appears to be a bug in this setting via PowerShell *** -AllowOwnerDeleteMessages $tSettingAllowOwnerDeleteMessages #>`
		-AllowTeamMentions $tSettingAllowTeamMentions `
		-AllowChannelMentions $tSettingAllowChannelMentions `
		<# *** There appears to be a bug in this setting via PowerShell *** -ShowInTeamsSearchAndSuggestions $tSettingShowInTeamsSearchAndSuggestions #>`
		-RetainCreatedGroup:$tSettingRetainCreatedGroup
	
	# Adding additional owners if there are multiple
	If($tOwnersCount -gt 1){
		$tOwnerIndex = 1
		do {
			Write-Host ("         Adding Owner: {0}" -f $workspace.Owners.Owner[$tOwnerIndex]) -ForegroundColor Blue
			Add-Teamuser -GroupID $team.GroupID -User $workspace.Owners.Owner[$tOwnerIndex] -Role Owner
			$tOwnerIndex++
		} until ($tOwnerIndex -ge $tOwnersCount)
	}

	# Adding members to the team
	$tMembers = $workspace.Members.Member
	$tMembersCount = $tMembers.Count
	If($tMembersCount -gt 0){
		ForEach ($tMember in $workspace.Members.Member) {
			Write-Host ("         Adding Members: {0}" -f $tMember) -ForegroundColor Blue
			Add-Teamuser -GroupID $team.GroupID -User $tMember -Role Member
		}
	}
	Else {
		Write-Host "         You must manually update the team to make it an Org-Wide team" -ForegroundColor Blue
		$orgwideWarning = new-object -comobject wscript.shell
		$temp = $orgwideWarning.popup("You must manually update the team to make it an Org-Wide team.  Press any key to continue.",0,"Org-Wide Teams",1)
	}

	# Adding Additional Channels
	ForEach ($tChannel in $workspace.Channels.Channel) {
		Write-Host ("         Adding Channel: {0}" -f $tChannel.DisplayName) -ForegroundColor Blue
		$temp = New-TeamChannel -GroupId $team.GroupId -DisplayName $tChannel.DisplayName -Description $tChannel.Description
	}

	# Pausing in between workspaces
	Start-Sleep -Seconds $pause
} #>
# Screen Output
Write-Host "   Teams Configuration Complete" -ForegroundColor Green
#endregion

#region Completed Script
# Disconnect from SharePoint
Disconnect-MicrosoftTeams

# Output Script Info
Complete-CatScript $scriptStart[0]

# Stop logging
Stop-Transcript
#endregion