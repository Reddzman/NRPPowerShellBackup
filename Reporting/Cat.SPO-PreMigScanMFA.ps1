<# Cat.SPO_PreMigScanMFA.ps1

Requirements
	-	Account needs to have access to use the "PnP Office 365 Management Shell"
		(Azure AD -> Enterprise Applications -> All applications ->  PnP Office 365 Management Shell -> Permissions)
	-	Account needs to have SharePoint Admin role

This script generates reports for across all site collections to identify the following information:
	-	Site Collection Information
	-	Web Information
	-	Large Lists
	-	InfoPath Forms
	-	Workflows

Potential Improvements
	-	What Usage (Page load, Last Accessed, etc.) data can be collected w/ PS?
	-	InfoPath Report - Add list modified information.  What other InfoPath Data could we grab?

Table of Contents
	1.	Input & Output Files
	2.	Variables
	3.	Snapin(s), Modules & Environment Connection(s)
	4.	Main Script
	5.	Collect Site Collection Info
	6.	Collect Web Info
	7.	Check for Large Lists
	8.	Check for InfoPath Forms Libraries & Lists
	9.	Check for Workflows
	10.	Checked Out Documents
#>

#region Script Parameters
Param(
	[string] $client = "westloan",
	[String] $adminURL = "https://westloan-admin.sharepoint.com/",
	[boolean] $includePersonalSite = $false,
	[String] $tenantID = "e19bb33a-929a-4d93-985a-1a08f7fa8e3d"
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
Start-Transcript -Path ("..\_Log FIles\PreMigScan-{0}.log" -f $scriptStart[1])
#endregion

#region 1.	Input & Output Files
# File Paths
$fileSites = ("..\_Output Files\{0}-SiteCollections-{1}.txt" -f $client, $scriptStart[1])
$fileWebs = ("..\_Output Files\{0}-Webs-{1}.txt" -f $client, $scriptStart[1])
$fileLists = ("..\_Output Files\{0}-LargeLists-{1}.txt" -f $client, $scriptStart[1])
$fileInfoPath = ("..\_Output Files\{0}-InfoPathFile-{1}.txt" -f $client, $scriptStart[1])
$fileWorkflows = ("..\_Output Files\{0}-Workflows-{1}.txt" -f $client, $scriptStart[1])
$fileCheckedOut = ("..\_Output Files\{0}-CheckedOut-{1}.txt" -f$client,  $scriptStart[1])

# Create Output File(s)
Out-File -FilePath $fileSites -InputObject "Site|SiteURL|LastContentModified|Storage|ID|Owner|SecondaryOwner|Template|HubSite|HubSiteID|GroupID|RelatedGroupID|IsTeam|PWAEnabled"
Out-File -FilePath $fileWebs -InputObject "Site|SiteURL|Web|WebURL|LastModified|MasterPage|NoCrawl|Template"
Out-File -FilePath $fileLists -InputObject "Site|SiteURL|Web|WebURL|List|ListURL|ItemCount|Type|Created|ID|LastItemDeleted|LastItemModifed|MajorVersionLimit|MinorVersionLimit|NoCrawl"
Out-File -FilePath $fileInfoPath -InputObject "Site|SiteURL|Web|WebURL|List|ListURL"
Out-File -FilePath $fileWorkflows -InputObject "Site|SiteURL|Web|WebURL|List|ListURL|Workflow|TaskList|HistoryList|Created|Modified"
Out-File -FilePath $fileCheckedOut -InputObject "Site|SiteURL|Web|WebURL|List|ListURL|Document|CheckedOutBy|ID|CreatedBy|Version|ModifiedBy|Created|Modified|ItemURL"
#endregion

#region 2.	Variables
# User Inputs
#$spCred = Get-Credential -Message "Enter your credentials for SharePoint:"

#region Connecting to SPO services
Write-Host
Write-Host "Connecting to SPO Services" -ForegroundColor Cyan
Write-Host ("   Connecting to SharePoint Online with the SharePoint Management Shell: {0}" -f $AdminURL)
Connect-SPOService -Url $AdminURL
Write-Host ("   Connecting to SharePoint Online with PnP PowerShell: {0}" -f $AdminURL)
$pnpConnect = Connect-PnPOnline -Url $AdminURL.Replace("-admin","").Replace("-Admin","") -PnPO365ManagementShell -ReturnConnection -LaunchBrowser # *** The -LaunchBrowser flag may have issues with MFA ***
$pnpToken = Get-PnPAccessToken
Disconnect-PnPOnline


# Script Variables
$maxItemCount = 0 # 5000 = SPO List view limit
$skipLists = @(
	"Content and Structure Reports";
	"Customized Reports";
	"Form Templates";
	"Master Page Gallery";
	"Reusable Content";
	"Style Library";
	"Workflow Tasks";
	"MicroFeed",
	"Social"
)
# Checked out Query
#$queryCheckedOut = New-Object Microsoft.SharePoint.Client.CamlQuery
$queryCheckedOut = "@
<View>
	<Query>
		<Where>
			<IsNotNull>
				<FieldRef Name='CheckoutUser' />
			</IsNotNull>
		</Where>
	</Query>
</View>"

# Output Parameters & Variables
Write-Host
Write-Host "Script Parameters" -ForegroundColor Cyan
Write-Host ("   Client: {0}" -f $client)
Write-Host ("   Admin URL: {0}" -f $adminURL)
Write-Host ("   Include Personal Sites: {0}" -f $includePersonalSite)
Write-Host ("   Tenant ID: {0}" -f $tenantID)
Write-Host
Write-Host "User Inputs" -ForegroundColor Cyan
Write-Host ("   Credentials: {0}" -f $spCred.UserName)
Write-Host
Write-Host "Script Variables" -ForegroundColor Cyan
Write-Host ("   Max List Item Count: {0}" -f $maxItemCount)
Write-Host
Write-Host "Lists to be skipped" -ForegroundColor Cyan
$skipLists
Write-Host
Write-Host "Checked Out Documents Query" -ForegroundColor Cyan
$queryCheckedOut
#endregion

#region 3.	Snapin(s), Modules & Environment Connection(s)
# Load SharePoint PowerShell Snap-In(s)
Import-Module Microsoft.Online.SharePoint.PowerShell -DisableNameChecking -ErrorAction SilentlyContinue
Import-Module SharePointPnPPowerShellOnline -DisableNameChecking
Import-Module MicrosoftTeams

# Load Custom Modules
Import-Module "..\Reporting\Cat.ReportingFunctions.psm1"
#endregion

#region 4.	Main Script
# Screen Output
Write-Host
Write-Host "Starting SharePoint Online Scan" -ForegroundColor Cyan
	
# Get site collection object
Write-Host "   Retrieving Site Collections"
Connect-SPOService -Url $adminURL
$sites = Get-SPOSite -Limit all -IncludePersonalSite:$includePersonalSite | Where{$_.Url -like "*sharepoint.com*"} | Sort-Object -Property Url

# Loop through site collections for web application
ForEach ($site in $sites)
{
	# Screen Output for Site
	Write-Host ("   Scanning site collection: {0}" -f $site.URL) -ForegroundColor Magenta

	# Adding Administrator
    Write-Host ("      Adding Permissions: {0}" -f $spCred.UserName)
    $tempUser = Set-SPOUser -Site $site.Url -LoginName $spCred.UserName -IsSiteCollectionAdmin $true

	#region 5.	Collect Site Collection Info
	Write-Host "      Creating Site Record" 
	$pnpSiteConnect = Connect-PnPOnline -Url $site.Url -AccessToken $pnpToken -ReturnConnection
	$pnpSite = Get-PnPSite -Includes RootWeb, Owner, SecondaryContact, RootWeb.LastItemUserModifiedDate, Usage, ID, IsHubSite, HubSiteID, GroupID, RelatedGroupID -Connection $pnpSiteConnect
	if ($pnpSite.RootWeb.Title -eq $null){
		$siteTitle = ""
	}
	else {
		$siteTitle = $pnpSite.RootWeb.Title.Replace("|","\")
	}
	$teamsConnect = Connect-MicrosoftTeams -TenantID $tenantID -Credential $spCred
	try {
		$tempTeam = Get-Team -GroupID $pnpSite.GroupID
		$isTeam = $true
	}
	catch {
		$isTeam = $false
	}
	Out-File -FilePath $fileSites -Append -InputObject ("{0}|{1}|{2}|{3}|{4}|{5}|{6}|{7}|{8}|{9}|{10}|{11}|{12}|{13}" -f 
		$siteTitle, 
		$pnpSite.Url, 
		$pnpSite.RootWeb.LastItemUserModifiedDate, 
		$pnpSite.Usage.Storage,
		$pnpSite.ID,
		$pnpSite.Owner.Email,
		$pnpSite.SecondaryContact.Email,
		$site.Template,
		$pnpSite.IsHubSite,
		$pnpSite.HubSiteID,
		$pnpSite.GroupID,
		$pnpSite.RelatedGroupID,
		$isTeam,
		$site.PWAEnabled)
	#endregion

	# Loop through sub sites
	$webs = @()
	$webs += Get-PnPWeb -Includes LastItemUserModifiedDate, MasterUrl, NoCrawl, WebTemplate -Connection $pnpSiteConnect
	$webs += Get-PnPSubWebs -Recurse -Includes LastItemUserModifiedDate, MasterUrl, NoCrawl, WebTemplate -Connection $pnpSiteConnect | Sort-Object -Property Url
	ForEach ($web in $webs)
	{
		# Screen output for web
		Write-Host ("      Scanning Web: {0}" -f $web.Title) -ForegroundColor Blue

		#region 6.	Collect Web Info		
		# Get Modified Dates
		Write-Host "         Creating Web Record"
		if ($web.Title -eq $null){
			$webTitle = ""
		}
		else {
			$webTitle = $web.Title.Replace("|","\")
		}
		Out-File -FilePath $fileWebs -Append -InputObject ("{0}|{1}|{2}|{3}|{4}|{5}|{6}|{7}" -f 
			$siteTitle, 
			$pnpSite.Url, 
			$webTitle,
			$web.Url, 
			$web.LastItemUserModifiedDate,
			$web.MasterURL,
			$web.NoCrawl,
			$web.WebTemplate)
		#endregion
		
		# Loop through all lists and libraries in a sub site
		$lists = Get-PnPList -Web $web -Includes RootFolder.ServerRelativePath, ContentTypes, WorkflowAssociations, BaseType, BaseTemplate -Connection $pnpSiteConnect | Where{($_.Hidden -eq $false) -and ($skipLists -notcontains $_.Title)}
		ForEach ($list in $lists)
		{
			# Screen output for list
			Write-Host ("         Scanning List: {0}" -f $list.Title) -ForegroundColor DarkGray
			if ($list.Title -eq $null){
				$listTitle = ""
			}
			else {
				$listTitle = $list.Title.Replace("|","\")
			}
			
			#region 7.	Check for Large Lists
			If ($list.ItemCount -gt $maxItemCount)
			{
				# Creating Large List Record
				Write-Host ("            Creating List record: {0}" -f $list.Title)
				Out-File -FilePath $fileLists -Append -InputObject ("{0}|{1}|{2}|{3}|{4}|{5}|{6}|{7}|{8}|{9}|{10}|{11}|{12}|{13}|{14}" -f 
					$siteTitle, 
					$site.Url, 
					$webTitle,
					$web.Url, 
					$listTitle, 
					$list.RootFolder.ServerRelativePath.DecodedUrl, 
					$list.ItemCount,
					$list.BaseType,
					$list.Created,
					$list.ID,
					$list.LastItemDeletedDate,
					$list.LastItemModifiedDate,
					$list.MajorVersionLimit,
					$list.MajorWithMinorVersionsLimit,
					$list.NoCrawl)
			} # If ($list.Items.Count -gt $MaxItems)
			#endregion

			#region 8.	Check for InfoPath Forms Libraries
			# Checking based on Library type
			If (($list.BaseType -eq "DocumentLibrary") -and ($list.BaseTemplate -eq "115")) {
				# Creating InfoPath Record
				Write-Host ("            Creating InfoPath Library record: {0}" -f $list.Title)
				Out-File -FilePath $fileInfoPath -Append -InputObject ("{0}|{1}|{2}|{3}|{4}|{5}" -f 
					$siteTitle, 
					$pnpSite.Url, 
					$webTitle,
					$web.Url, 
					$listTitle, 
					$list.RootFolder.ServerRelativePath.DecodedUrl)
			} # If (InfoPath)
			try {
				$listInfoPath = Get-PnPList -Identity $list -Includes Forms -Connection $pnpSiteConnect -ErrorAction Stop
				If ($listInfoPath.Forms.ServerRelativeURL -contains "*.displayifs.aspx")
				{
					# Creating InfoPath Record
					Write-Host ("            Creating InfoPath List Form record: {0}" -f $list.Title)
					Out-File -FilePath $fileInfoPath -Append -InputObject ("{0}|{1}|{2}|{3}|{4}|{5}" -f 
						$siteTitle, 
						$pnpSite.Url, 
						$webTitle,
						$web.Url, 
						$listTitle, 
						$list.RootFolder.ServerRelativePath.DecodedUrl)
				} # If (InfoPath)
			}
			catch {
				Write-Host ("            List doesn't support Infopath edited forms: {0}" -f $list.Title)
				<#Write-Host "!! EXCEPTION OCCURRED !!" -ForegroundColor Red
				Write-Host "!! Exception Type: $($_.Exception.GetType().FullName)" -ForegroundColor Red
				Write-Host "!! Exception Message: $($_.Exception.Message)" -ForegroundColor Red #>
			}
			#endregion

			#region 9.	Check for SP Workflows
			ForEach ($workflow in $list.WorkflowAssociations)
			{
				If ($workflow -ne $null)
				{
					Write-Host ("            Creating Workflow record: {0}" -f $workflow.Name) -ForegroundColor DarkRed
					if ($workflow.Name -eq $null){
						$wfName = ""
					}
					else {
						$wfName = $workflow.Name.Replace("|","\")
					}
					Out-File -FilePath $fileWorkflows -Append -InputObject ("{0}|{1}|{2}|{3}|{4}|{5}|{6}|{7}|{8}|{9}|{10}" -f 
						$siteTitle, 
						$pnpSite.Url, 
						$webTitle,
						$web.Url, 
						$listTitle, 
						$list.RootFolder.ServerRelativePath.DecodedUrl, 
						$wfName,
						$workflow.TaskListTitle,
						$workflow.HistoryListtitle,
						$workflow.Created,
						$workflow.Modified)
				}
			} # ForEach ($workflow in $list.WorkflowAssociations)
			#endregion

			#region 10.	Checked Out Documents
			If (($list.BaseType -eq "DocumentLibrary") -and ($list.ItemCount -gt 0))
			{
				# Loop through Items in List
				Write-Host "            Scanning for checked out documents" -ForegroundColor Yellow
				$items = Get-PnPListItem -Web $web -List $list -Query $queryCheckedOut -Connection $pnpSiteConnect
				ForEach ($item in $items)
				{
					Write-Host ("               Creating Checked Out record: {0}" -f $item["FileLeafRef"])
					If ($item["Modified_x0020_By"] -eq $null){
						$userModified = ""
					}
					else {
						$userModified = $item["Modified_x0020_By"].Replace("|","/")
					}
					Out-File -FilePath $fileCheckedOut -Append -InputObject ("{0}|{1}|{2}|{3}|{4}|{5}|{6}|{7}|{8}|{9}|{10}|{11}|{12}|{13}|{14}" -f 
						$siteTitle, 
						$site.Url, 
						$webTitle,
						$web.Url, 
						$listTitle, 
						$list.RootFolder.ServerRelativePath.DecodedUrl, 
						$item["FileLeafRef"].Replace("|","\"), 
						$item["CheckoutUser"].Email,
						$item.ID,
						$item["Author"].Email ,
						$item["_UIVersionString"],
						$userModified,
						$item["Created"],
						$item["Modified"],
						$item["FileRef"])
				}
			}
			#endregion
		} # ForEach ($list in $web.Lists)
			
		# Screen Output for Web
		Write-Host ("      Web Scan Complete" -f $web.Title) -ForegroundColor Green
	} # ForEach ($web in $webs)

	# Dispose of site object
	Write-Host ("   Site Scan Complete" -f $web.Title) -ForegroundColor Green
	Disconnect-PnPOnline  -Connection $pnpSiteConnect
} # ForEach ($site in $sites)

# Screen Output for Farm
Write-Host ("Scan Complete" -f $web.Title) -ForegroundColor Green
#endregion

#region Completed Screen Output
Disconnect-SPOService
Complete-CatScript $scriptStart[0]
Stop-Transcript
#endregion