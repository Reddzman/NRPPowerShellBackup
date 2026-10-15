<#  Cat.MT-DeployCrisisTeam.ps1

This script creates Microsoft Teams and Channels based on an input XML file.

Dependencies
	-	PnP PowerShell - https://docs.microsoft.com/en-us/powershell/sharepoint/sharepoint-pnp/sharepoint-pnp-cmdlets?view=sharepoint-ps
	-	General Functions - ..\General\Cat.PnP-GeneralFunctionsFunctions.psm1
	-	Input XML File - Example:  ..\_Input Files\Testing\Cat.PnP-CreateTeamsAndChannels-Test.xml

Call script with parameters
	cd C:\GitHub\CatPowerShell\Configuration\
	.\Cat.PnP-CreateTeamsAndChannels.ps1 -inputFile "Samples\Configuration\Cat.PnP-CreateTeamsAndChannels-Test.xml" -clientAppID "caf6b11a-cf03-416c-b0f0-67a09a8e30ee" -clientCertPath "C:\GitHub\NRPDev\AppRegistration\PowerShell App Registration.pfx" -clientTenant "nrpdev2" -addChannelsToExistingTeams $false

Table of Contents
	1.	Input & Output Files
	2.	Variables
	3.	Snapin(s), Modules & Environment Connection(s)
	4.	Main Script
#>

#region Script Parameters
Param(	[string] $inputFile = "Samples\Configuration\Cat.PnP-CreateTeamsAndChannels-Test.xml",
		[string] $clientAppID = "{Client app registration ID from Azure}", # Leave Blank for interactive login
		[string] $clientCertPath = "{Path to client certificate file must be installed on local user}",# Leave Blank for interactive login
		[string] $clientTenant = "nrpdev2", # IE. {ClientTenant}.sharepoint.com
		[boolean] $addChannelsToExistingTeams = $false, # If set to false, script will identify existing teams and move on
		[int] $creationPause = 0 # This will add a pause between team workspaces. Great for testing/troubleshooting
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
Start-Transcript -Path ("..\_Log FIles\{0}.m365-TeamsRapidDeployment-{1}.log" -f $clientTenant, $scriptStart[1])
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
if ($addChannelsToExistingTeams) {
	# Get confirmation to remove permissions
	Write-Host
	Write-Host "*** When Adding channels to existing teams, existing members will see new channels. Communication with existing teams is recommended. ***" -ForegroundColor Yellow
	$userProceed = Read-Host "   Are you sure you want to proceed with adding to existing teams (Yes or No)" 
	if ($userProceed.ToLower() -eq "yes") {
		$addChannelsToExistingTeams = $true
	}
	else {
		$addChannelsToExistingTeams = $false
	}
}
if ($clientAppID) {
	Write-Host
	Write-Host "User Input Variables" -ForegroundColor Cyan
	$certPass = ConvertTo-SecureString -String (Read-Host -Prompt "   Enter certificate password for the app registration") -AsPlainText -Force
	$cert = Get-PnPAzureCertificate -Path $clientCertPath -Password $certPass
	Remove-Variable certPass
}

# Load Settings from XML
# Script Variables
$tenantDomain = ("{0}.onmicrosoft.com" -f $clientTenant)
$spoAdminURL = ("https://{0}-admin.sharepoint.com" -f $clientTenant)

# Output Parameters & Variables
Write-Host
Write-Host "Script Parameters" -ForegroundColor Cyan
Write-Host ("   Input File: {0}" -f $inputFile)
Write-Host ("   App ID: {0}" -f $clientAppID)
Write-Host ("   Cert Path: {0}" -f $clientCertPath)
Write-Host ("   client Tenant: {0}" -f $clientTenant)
Write-Host ("   Add channels to existing Teams: {0}" -f $addChannelsToExistingTeams)
Write-Host ("   Teams Creation Pause (Seconds): {0}" -f $creationPause)
if ($clientAppID) {
	Write-Host
	Write-Host "User Inputs" -ForegroundColor Cyan
	Write-Host ("   Certificate Thumbprint: {0}" -f $cert.Thumbprint)
}
Write-Host
Write-Host "Script Variables" -ForegroundColor Cyan
Write-Host ("   Microsoft Tenant: {0}" -f $tenantDomain)
Write-Host ("   SPO Admin URL: {0}" -f $spoAdminURL)
#endregion

#region 3.	Snapin(s), Modules & Environment Connection(s)
# Load MS / PnP PowerShell Modules
Get-Module -ListAvailable | Where-Object{($_.Name -like "*SharePoint*") -or ($_.Name -like "*PnP*") } | Remove-Module
Import-Module PnP.PowerShell
#endregion
    
#region 4.	Main Script
#region Connecting To SP Online
Write-Host
Write-Host "Connecting to Services" -ForegroundColor Cyan
Write-Host ("   Connecting to SharePoint Online Administraton: {0}" -f $spoAdminURL)
try {
	if ($clientAppID) {
		Connect-PnPOnline $spoAdminURL -tenant $tenantDomain -ClientId $clientAppID -Thumbprint $cert.Thumbprint # PnP Modules w/ App Reg
	} else {
		Connect-PnPOnline -Url $spoAdminURL -Interactive # PnP Modules w/o App Reg
	} 
}
catch {
	Write-Host
	throw ("Unable to connect: {0}" -f $_.Exception)	
}
#endregion

# Screen Output
Write-Host
Write-Host ("Creating Microsoft Teams: {0}" -f $inputFile) -ForegroundColor Cyan

# Looping through XML for teams workspaces
ForEach ($workspace in $inFile.CreateTeams.Team){
	Write-Host
	Write-Host ("   Creating Workspace: {0}" -f $workspace.DisplayName) -ForegroundColor Magenta

	#region Checking if team exists
	$team = Get-PnPTeamsTeam -Identity $workspace.DisplayName
	if ($team) {
		Write-Host ("      Team workspace already exists") -ForegroundColor Yellow
		$teamExists = $true
	} else {
		$team = Get-PnPTeamsTeam -Identity $workspace.MailNickName
		if ($team) {
			Write-Host ("      Team workspace already exists") -ForegroundColor Yellow
			$teamExists = $true
		} else {
			$teamExists = $false
		}
	}
	#endregion

	#region Create team workspace if one doesn't already exist
	if (-not $teamExists) {
		#region Determine workspace settings
		Write-Host "      Determining Settings"
		if ($workspace.TeamSettings.AllowAddRemoveApps.ToLower() -eq "true") {$tAllowAddRemoveApps = $true} Else {$tAllowAddRemoveApps = $false}
		if ($workspace.TeamSettings.AllowChannelMentions.ToLower() -eq "true") {$tAllowChannelMentions = $true} Else {$tAllowChannelMentions = $false}
		if ($workspace.TeamSettings.AllowCreateUpdateChannels.ToLower() -eq "true") {$tAllowCreateUpdateChannels = $true} Else {$tAllowCreateUpdateChannels = $false}
		if ($workspace.TeamSettings.AllowCreateUpdateRemoveConnectors.ToLower() -eq "true") {$tAllowCreateUpdateRemoveConnectors = $true} Else {$tAllowCreateUpdateRemoveConnectors = $false}
		if ($workspace.TeamSettings.AllowCreateUpdateRemoveTabs.ToLower() -eq "true") {$tAllowCreateUpdateRemoveTabs = $true} Else {$tAllowCreateUpdateRemoveTabs = $false}
		if ($workspace.TeamSettings.AllowCustomMemes.ToLower() -eq "true") {$tAllowCustomMemes = $true} Else {$tAllowCustomMemes = $false}
		if ($workspace.TeamSettings.AllowDeleteChannels.ToLower() -eq "true") {$tAllowDeleteChannels = $true} Else {$tAllowDeleteChannels = $false}
		if ($workspace.TeamSettings.AllowGiphy.ToLower() -eq "true") {$tAllowGiphy = $true} Else {$tAllowGiphy = $false}
		if ($workspace.TeamSettings.AllowGuestCreateUpdateChannels.ToLower() -eq "true") {$tAllowGuestCreateUpdateChannels = $true} Else {$tAllowGuestCreateUpdateChannels = $false}
		if ($workspace.TeamSettings.AllowGuestDeleteChannels.ToLower() -eq "true") {$tAllowGuestDeleteChannels = $true} Else {$tAllowGuestDeleteChannels = $false}
		if ($workspace.TeamSettings.AllowOwnerDeleteMessages.ToLower() -eq "true") {$tAllowOwnerDeleteMessages = $true} Else {$tAllowOwnerDeleteMessages = $false}
		if ($workspace.TeamSettings.AllowStickersAndMemes.ToLower() -eq "true") {$tAllowStickersAndMemes = $true} Else {$tAllowStickersAndMemes = $false}
		if ($workspace.TeamSettings.AllowTeamMentions.ToLower() -eq "true") {$tAllowTeamMentions = $true} Else {$tAllowTeamMentions = $false}
		if ($workspace.TeamSettings.AllowUserDeleteMessages.ToLower() -eq "true") {$tAllowUserDeleteMessages = $true} Else {$tAllowUserDeleteMessages = $false}
		if ($workspace.TeamSettings.AllowUserEditMessages.ToLower() -eq "true") {$tAllowUserEditMessages = $true} Else {$tAllowUserEditMessages = $false}
		if ($workspace.TeamSettings.ShowInTeamsSearchAndSuggestions.ToLower() -eq "true") {$tShowInTeamsSearchAndSuggestions = $true} Else {$tShowInTeamsSearchAndSuggestions = $false}
		if ($workspace.TeamSettings.Classification -eq "") {$tClassification = $null} Else {$tClassification = $workspace.TeamSettings.Classification}
		if ($workspace.TeamSettings.ResourceBehaviorOptions -eq "") {$tResourceBehaviorOptions = $null} Else {$tResourceBehaviorOptions = $workspace.TeamSettings.ResourceBehaviorOptions}
		if ($workspace.TeamSettings.SensitivityLabels -eq "") {$tSensitivityLabels = $null} Else {$tSensitivityLabels = $workspace.TeamSettings.SensitivityLabels}
		#endregion

		#region Provision Team
		Write-Host "      Creating Team"
		$team = New-PnPTeamsTeam `
			-DisplayName $workspace.DisplayName `
			-Description $workspace.Description `
			-MailNickName $workspace.MailNickName `
			-AllowAddRemoveApps $tAllowAddRemoveApps `
			-AllowChannelMentions $tAllowChannelMentions `
			-AllowCreateUpdateChannels $tAllowCreateUpdateChannels `
			-AllowCreateUpdateRemoveConnectors $tAllowCreateUpdateRemoveConnectors `
			-AllowCreateUpdateRemoveTabs $tAllowCreateUpdateRemoveTabs `
			-AllowCustomMemes $tAllowCustomMemes `
			-AllowDeleteChannels $tAllowDeleteChannels `
			-AllowGiphy $tAllowGiphy `
			-AllowGuestCreateUpdateChannels $tAllowGuestCreateUpdateChannels `
			-AllowGuestDeleteChannels $tAllowGuestDeleteChannels `
			-AllowOwnerDeleteMessages $tAllowOwnerDeleteMessages `
			-AllowStickersAndMemes $tAllowStickersAndMemes `
			-AllowTeamMentions $tAllowTeamMentions `
			-AllowUserDeleteMessages $tAllowUserDeleteMessages `
			-AllowUserEditMessages $tAllowUserEditMessages `
			-GiphyContentRating $workspace.TeamSettings.GiphyContentRating `
			-Visibility $workspace.TeamSettings.Visibility `
			-ShowInTeamsSearchAndSuggestions $tShowInTeamsSearchAndSuggestions `
			-Classification $tClassification `
			-Owners $workspace.Owners.Owner `
			-Members $workspace.Members.Member `
			-ResourceBehaviorOptions $tResourceBehaviorOptions `
			-SensitivityLabels $tSensitivityLabels
		#endregion
	}
	#endregion

	#region Adding additional channels
	if ((-not $teamExists) -or ($teamExists -and $addChannelsToExistingTeams)) {
		ForEach ($channel in $workspace.Channels.Channel) {
			$channelExists = Get-PnPTeamsChannel -Team $team -Identity $channel.DisplayName
			if ($channelExists) {
				Write-Host ("      Teams Channel already exists: {0}" -f $channel.DisplayName) -ForegroundColor Yellow
			} else {
				Write-Host ("      Adding Channel: {0}" -f $channel.DisplayName)
				if ($channel.IsFavoriteByDefault.ToLower() -eq "true") {$cIsFavoriteByDefault = $true} Else {$cIsFavoriteByDefault = $false}
				switch ($channel.ChannelType) {
					Standard {$temp = Add-PnPTeamsChannel -Team $team -DisplayName $channel.DisplayName -Description $channel.Description -IsFavoriteByDefault $cIsFavoriteByDefault}
					Private {$temp = Add-PnPTeamsChannel -Team $team -DisplayName $channel.DisplayName -Private -OwnerUPN $channel.OwnerUPN -Description $channel.Description}
					#Shared {$temp = Add-PnPTeamsChannel -Team $team -DisplayName $channel.DisplayName -ChannelType Shared -OwnerUPN $channel.OwnerUPN -Description $channel.Description -IsFavoriteByDefault $cIsFavoriteByDefault} # This is prerelease capability. Documation ahead of module
					Default {Write-Host ("      Channel Type not valid for {0}" -f $channel.DisplayName) -ForegroundColor Red}
				}
			}
		}
	} else {
		Write-Host "      Skipping Channel Creation" -ForegroundColor Yellow
	}
	#endregion

	# Pausing in between workspaces
	Write-Host ("      Team Completed, Pausing for {0} seconds" -f $creationPause) -ForegroundColor Green
	if ($creationPause -gt 0) {Start-Sleep -Seconds $creationPause}
}

# Screen Output
Write-Host
Write-Host "Teams Configuration Complete" -ForegroundColor Green
#endregion

#region Completed Script
# Disconnect from SharePoint
Disconnect-PnPOnline

# Output Script Info
Complete-CatScript $scriptStart[0]

# Stop logging
Stop-Transcript
#endregion