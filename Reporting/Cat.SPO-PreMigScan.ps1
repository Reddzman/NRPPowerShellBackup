<# Notes
#region Cat.SPO_PreMigScan.ps1

This script generates reports for across all site collections to identify the following information:
	-	Site Collection Information
	-	Web Information
	-	Large Lists
	-	InfoPath Forms
	-	Workflows
	-	Complex Columns
	-	Content Type Hub Content Types
	-	Managed Metadata
	-	Guest Users by Site

Dependencies
	-	This script is designed to leverage the new PnP PowerShell - https://pnp.github.io/powershell/cmdlets/released/Add-PnPAlert.html
	-	This script is designed to use App Registration permissions

Potential Improvements
	-	What Usage (Page load, Last Accessed, etc.) data can be collected w/ PS?
	-	InfoPath Report - Add list modified information.  What other InfoPath Data could we grab?
	-	Convert Data structure to CSV using Export-CSV

Table of Contents
	1.	Input & Output Files
	2.	Variables
	3.	Snapin(s), Modules & Environment Connection(s)
	4.	Main Script
	5.	Custom Managed Metadata
	6.	Content Type Hub
	7.	External Users
	8.	Collect Site Collection Info
	9.	Collect Web Info
	10.	Check for Large Lists
	11.	Check for InfoPath Forms Libraries & Lists
	12.	Check for Workflows
	13.	Checked Out Documents
	14.	Complex Columns

Sample Execution Command: .\Cat.SPO-PreMigScan.ps1 -clientTenant {ClientTenant} -includePersonalSite $true -tenantID {tenantIDGUID} -appID {AppIDGUID} -certPath c:\github\CatPowerShell\_InputFiles\{CertName}.pfx

#endregion #>

#region Script Parameters
Param(
	[string] $clientTenant = "nrpdev2",
	[boolean] $includePersonalSite = $false,
	# TODO: Use these for App Registrationsget-host
	[String] $appID = "57577043-3de0-40c0-9c68-7960e2cebe8b",
	[String] $certPath = "C:\GitHub\NRPDev\AppRegistration\PnP PowerShell Authentication.pfx"
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
Start-Transcript -Path ("..\_Log FIles\{0}-PreMigScan-{1}.log" -f $clientTenant,$scriptStart[1])
#endregion

#region 1.	Input & Output Files
# File Paths
$fileSites = ("..\_Output Files\{0}-SiteCollections-{1}.txt" -f $clientTenant, $scriptStart[1])
$fileWebs = ("..\_Output Files\{0}-Webs-{1}.txt" -f $clientTenant, $scriptStart[1])
$fileLists = ("..\_Output Files\{0}-LargeLists-{1}.txt" -f $clientTenant, $scriptStart[1])
$fileInfoPath = ("..\_Output Files\{0}-InfoPathFile-{1}.txt" -f $clientTenant, $scriptStart[1])
$fileWorkflows = ("..\_Output Files\{0}-Workflows-{1}.txt" -f $clientTenant, $scriptStart[1])
$fileCheckedOut = ("..\_Output Files\{0}-CheckedOut-{1}.txt" -f$clientTenant, $scriptStart[1])
$fileComplexColumns = ("..\_Output Files\{0}-ComplexColumns-{1}.txt" -f$clientTenant, $scriptStart[1])
$fileContentTypeHub = ("..\_Output Files\{0}-ContentTypeHub-{1}.txt" -f$clientTenant, $scriptStart[1])
$fileManagedMetaDataTermGroups = ("..\_Output Files\{0}-ManagedMetaDataTermGroups-{1}.txt" -f$clientTenant, $scriptStart[1])
$fileExternalUsers = ("..\_Output Files\{0}-ExternalUsers-{1}.txt" -f$clientTenant, $scriptStart[1])

# Create Output File(s)
Out-File -FilePath $fileSites -InputObject "Site|SiteURL|LastItemModified|Storage|ID|Owner|SecondaryOwner|Template|HubSite|HubSiteID|GroupID|RelatedGroupID|IsTeam|PWAEnabled"
Out-File -FilePath $fileWebs -InputObject "Site|SiteURL|Web|WebURL|LastModified|MasterPage|NoCrawl|Template"
Out-File -FilePath $fileLists -InputObject "Site|SiteURL|Web|WebURL|List|ListURL|ItemCount|Type|Created|ID|LastItemDeleted|LastItemModifed|MajorVersionLimit|MinorVersionLimit|NoCrawl|ListViewCount|UnfilteredLargeListViews|UnfilteredLargeListViewCount|CustomContentTypes|CustomContentTypeCount"
Out-File -FilePath $fileInfoPath -InputObject "Site|SiteURL|Web|WebURL|List|ListURL"
Out-File -FilePath $fileWorkflows -InputObject "Site|SiteURL|Web|WebURL|List|ListURL|Workflow|TaskList|HistoryList|Created|Modified"
Out-File -FilePath $fileCheckedOut -InputObject "Site|SiteURL|Web|WebURL|List|ListURL|Document|CheckedOutBy|ID|CreatedBy|Version|ModifiedBy|Created|Modified|ItemURL"
Out-File -FilePath $fileComplexColumns -InputObject "Site|SiteURL|Web|WebURL|List|ListURL|ColumnName|ColumnType"
Out-File -FilePath $fileContentTypeHub -InputObject "SiteURL|ContentType|Group|Parent"
Out-File -FilePath $fileManagedMetaDataTermGroups -InputObject "Site|SiteURL|TermGroupName|SiteCollectionGroup|NumberOfTermSets|NumberOfTerms"
Out-File -FilePath $fileExternalUsers -InputObject "Site|SiteURL|DisplayName|Email|AcceptedAs|WhenCreated|InvitedBy|UniqueId|UserId"
#endregion

#region 2.	Variables
# User Inputs
Write-Host
Write-Host "User Input Variables" -ForegroundColor Cyan
# TODO: Use this for app registrations
$certPass = ConvertTo-SecureString -String (Read-Host -Prompt "   Enter certificate password for the app registration") -AsPlainText -Force #>
<# TODO: Use this for basic authentications
Write-Host "   Enter credentials for SharePoint:"
$spCred = Get-Credential -Message "Enter your credentials for SharePoint:" #>

# Script Variables
$adminURL = ("https://{0}-admin.sharepoint.com/" -f $clientTenant)
$rootSiteUrl = ("https://{0}.sharepoint.com/" -f $clientTenant)
# TODO: Use this for app registrations
$tenant = ("{0}.onmicrosoft.com" -f $clientTenant) #>
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
# TODO: The following URLs are actually hard coded later in this script. Need to come up with a way to use the array instead.
$skipSites = @(
	"-public.sharepoint.com";
	".sharepoint.com/portals/hub"
)
$skipTermGroups = @(
	"Companies";
	"People";
	"Search Dictionaries"
)
$skipCTGroups = @(
	"_Hidden";
	"Business Intelligence";
	"Publishing Content Types";
	"Page Layout Content Types";
	"List Content Types";
	"Group Work Content Types";
	"SharePoint Online Content Types";
	"Document Content Types";
	"Display Template Content Types";
	"Digital Asset Content Types";
	"Special Content Types";
	"Folder Content Types";
	"Document Set Content Types"
)
$skipContentTypes = @(
	"Item";
	"Folder";
	"Document";
	"Document Set";	
	"Basic Page";	
	"Site Page";
	"Repost Page";
	"Wiki Page";
	"Web Part Page";
	"Event";
	"Page";
	"Article Page";
	"Welcome Page";
	"Error Page";
	"Picture";
	"Request for app purchase approval";
	"Task";
	"Announcement";
	"Image";
	"Audio";
	"Video";
	"Video Rendition";
	"Workflow Task";
	"SharePoint Server Workflow Task";
	"Workflow Task (SharePoint 2013)";
	"Summary Task";
	"Tabs in Search Results";
	"Static Page";
	"News Page";
	"Event Page";
	"Site Links";
	"SharePoint App";
	"Quick Links";
	"Promoted Links";
	"Office Data Connection File";
	"Universal Data Connection File";
	"Link to a Document";
	"Link";
	"Issue";
	"Home Page Promoted Links";
	"Discussion";
	"Message";
	"Contact"
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
$colComplexTypes = @("Calculated";"TaxonomyFieldType";"Lookup")
$skipColumns = @("Item Child Count";"Folder Child Count";"App Created By";"App Modified By";"Check In Comment";"Source Version (Converted Document)";
			"Source Name (Converted Document)";"Check In Comment";"DisplayVersion";"Label setting";"Retention label";"Comment count";"Like count";
			"Sensitivity";"Label applied by";"Retention label Applied")

# Output Parameters & Variables
Write-Host
Write-Host "Script Parameters" -ForegroundColor Cyan
Write-Host ("   Client: {0}" -f $clientTenant)
Write-Host ("   Include Personal Sites: {0}" -f $includePersonalSite)
Write-Host ("   Tenant ID: {0}" -f $tenantID)
# TODO: Use this for app registration
Write-Host ("   App ID: {0}" -f $appID)
Write-Host ("   Certificate: {0}" -f $certificate) #>
Write-Host
Write-Host "User Inputs" -ForegroundColor Cyan
# TODO: Use this for app registration
Write-Host ("   Certificate Password Legnth: {0}" -f $certPass.Length) #>
<# TODO: Use this for basic authentication
Write-Host ("   Credentials: {0}" -f $spCred.UserName) #>
Write-Host
Write-Host "Script Variables" -ForegroundColor Cyan
Write-Host ("   Admin URL: {0}" -f $adminURL)
Write-Host ("   Root Site URL: {0}" -f $rootSiteUrl)
Write-Host ("   Certificate Path: {0}" -f $certPath)
Write-Host ("   Tenant Domain: {0}" -f $tenant)
Write-Host ("   Max List Item Count: {0}" -f $maxItemCount)
<# TODO: Output for troubleshooting
Write-Host
Write-Host "Lists to be skipped" -ForegroundColor Cyan
$skipLists
Write-Host
Write-Host "Sites to be skipped" -ForegroundColor Cyan
$skipSites 
Write-Host
Write-Host "Term Groups to be skipped" -ForegroundColor Cyan
$skipTermGroups 
Write-Host
Write-Host "Content Type Groups to be skipped" -ForegroundColor Cyan
$skipCTGroups 
Write-Host
Write-Host "Content Type Groups to be skipped" -ForegroundColor Cyan
$skipContentTypes 
Write-Host
Write-Host "Checked Out Documents Query" -ForegroundColor Cyan
$queryCheckedOut #>
#endregion

#region 3.	Snapin(s), Modules & Environment Connection(s)
# Load SharePoint PowerShell Snap-In(s)
Import-Module PnP.PowerShell

# Load Custom Modules
Import-Module "..\Reporting\Cat.ReportingFunctions.psm1"
#endregion

#region 4.	Main Script
# Screen Output
Write-Host
Write-Host "Starting SharePoint Online Scan" -ForegroundColor Cyan

# Connect to SPO admin & get all site collections
Write-Host
Write-Host "   Retrieving Site Collections" -ForegroundColor Magenta
Write-Host ("      Connecting to SPO admin: {0}" -f $adminURL)
# TODO: Use this for app registration
Connect-PnPOnline -ClientId $appID -Url $adminURL -Tenant $tenant -CertificatePath $certPath -CertificatePassword $certPass #>
<# TODO: Use this for basic Authtentication	
$pnpSiteConnect = Connect-PnPOnline -Url $adminURL -Credentials $spCred #>
Write-Host "      Retrieving site collections"
$sites = Get-PnPTenantSite -IncludeOneDriveSites:$includePersonalSite -Detailed:$true -Filter "Url -notlike '-public.sharepoint.com' -and Url -notlike 'sharepoint.com/portals/'"

#region 5.	Custom Managed Metadata	
Write-Host
Write-Host "   Retrieving Custom Term Sets" -ForegroundColor Magenta

# Get Site Object
Write-Host "      Retrieving Site Object"
$pnpSite = Get-PnPSite -Includes RootWeb, Owner, SecondaryContact, RootWeb.LastItemUserModifiedDate, Usage, ID, IsHubSite, HubSiteID, GroupID, RelatedGroupID -Connection $pnpSiteConnect

# Get Site Object
Write-Host "      Retrieving Term Groups"
$mmdTermGroups = Get-PnPTermGroup | Where{($_.IsSystemGroup -ne $true ) -and ($skipTermGroups -notcontains $_.Name)} 

# Outputing terms
Write-Host "      Creating Managed Meta Data records"
Foreach ($termGroup in $mmdTermGroups)
{
	# Reset Term Counter
	$termCount = 0

	# Get term sets
	Write-Host "         Retrieving Term Sets" -ForegroundColor DarkGray
	$termSets = Get-PnPTermSet -TermGroup $termGroup.Name	

	#loop through term sets
	foreach ($termSet in $termSets)
	{
		# Get term sets
		Write-Host ("            Get Term Counts: {0}" -f $termSet.Name) -ForegroundColor Blue
		$terms = Get-PnPTerm -TermSet $termSet.Name -TermGroup $termGroup.Name
		$termCount = $termCount + $terms.Count
	}
	
	# Output to CSV
	Write-Host ("         Creating Term Group Record: {0}" -f $termGroup.Name) -ForegroundColor DarkGray
	Out-File -FilePath $fileManagedMetaDataTermGroups -Append -InputObject ("{0}|{1}|{2}|{3}|{4}|{5}|" -f 
	$pnpSite.RootWeb.Title.Replace("|","\"), 
	$pnpSite.Url, 
	$termGroup.Name,
	$termGroup.IsSiteCollectionGroup,
	$termSets.Count,
	$termCount)
} #foreach($termGroup in $mmdTermGroups)
#endregion

#region 6.	Content Type Hub
Write-Host
Write-Host ("   Scanning for Content Type Hub") -ForegroundColor Magenta
Write-Host ("      Connecting to SPO root site: {0}" -f $rootSiteUrl)
# TODO: Use this for app registration
Connect-PnPOnline -ClientId $appID -Url $rootSiteUrl -Tenant $tenant -CertificatePath $certPath -CertificatePassword $certPass #>
<# TODO: Use this for basic Authtentication	
$pnpSiteConnect = Connect-PnPOnline -Url $rootSiteUrl -Credentials $spCred #>

# Get CT Hub URL
Write-Host "      Retrieve Content Type Hub URL"
$contentTypeHubUrl = Get-PnPContentTypePublishingHubUrl

# Report on custom content types
if ($contentTypeHubUrl)
{
	Write-Host "      Retrieving Content Types"
	foreach ($contentType in Get-PnPContentType | Where{($skipCTGroups -notcontains $_.Group)})
	{
		# Output custom Content types
		Write-Host ("         Creating Content Type Record: {0}" -f $contentType.Name) -ForegroundColor DarkGray
		Out-File -FilePath $fileContentTypeHub -Append -InputObject ("{0}|{1}|{2}|{3}" -f 
		$contentTypeHubUrl, 
		$contentType.Name,
		$contentType.Group,
		$contentType.Parent.Name)
	} #foreach ($contentType in Get-PnPContentType)
} #if ($contentTypeHubUrl)
#endregion

#region 7.	External Users
Write-Host
Write-Host "   Scanning for external users" -ForegroundColor Magenta
foreach ($site in $sites) {
	Write-Host ("      Scanning Site: {0}" -f $site.Url)
	$extUsers = Get-PnPExternalUser -SiteUrl $site.url
	foreach ($user in $extUsers) {
		Write-Host ("         Creating External user Record: {0}" -f $user.AcceptedAs) -ForegroundColor DarkGray
		Out-File -FilePath $fileExternalUsers -Append -InputObject ("{0}|{1}|{2}|{3}|{4}|{5}|{6}|{7}|{8}" -f
			$site.Title,
			$site.Url,
			$user.DisplayName,
			$user.Email,
			$user.AcceptedAs,
			$user.WhenCreated,
			$user.InvitedBy,
			$user.UniqueId,
			$user.UserId)
	}
}
#endregion

# Loop through site collections for web application
ForEach ($site in $sites)
{
	# Screen Output for Site
	Write-Host
	Write-Host ("   Scanning site collection: {0}" -f $site.URL) -ForegroundColor Magenta

	<# TODO: If not using app registration:
	# Adding Administrator
    Write-Host ("      Adding Permissions: {0}" -f $spCred.UserName)
    $tempUser = Set-SPOUser -Site $site.Url -LoginName $spCred.UserName -IsSiteCollectionAdmin $true #>

	#region 8.	Collect Site Collection Info
	Write-Host ("      Creating Site Record: {0}" -f $site.Url) 
	Write-Host "         Connecting to site" -ForegroundColor DarkGray
	# TODO: Use this for app registration
	Connect-PnPOnline -ClientId $appID -Url $site.Url -Tenant $tenant -CertificatePath $certPath -CertificatePassword $certPass #>
	<# TODO: Use this for basic Authtentication	
	$pnpSiteConnect = Connect-PnPOnline -Url $site.Url -Credentials $spCred #>

	# Retrieve Site object if not a redirction site.
	If ($site.Template -ne "RedirectSite#0") {
		$pnpSite = Get-PnPSite -Includes RootWeb, Owner, SecondaryContact, RootWeb.LastItemUserModifiedDate, Usage, ID, IsHubSite, HubSiteID, GroupID, RelatedGroupID -Connection $pnpSiteConnect
	}

	#Create Site Record
	Write-Host "         Creating site record" -ForegroundColor DarkGray
	if ($site.Title -eq $null){
		$siteTitle = ""
	}
	else {
		$siteTitle = $site.Title.Replace("|","\")
	}
	try {
		$tempTeam = Get-PnPTeamsTeam -Identity $pnpSite.GroupId.ToString()
		$isTeam = $true
	}
	catch {
		$isTeam = $false
	}
	Out-File -FilePath $fileSites -Append -InputObject ("{0}|{1}|{2}|{3}|{4}|{5}|{6}|{7}|{8}|{9}|{10}|{11}|{12}|{13}" -f 
		$siteTitle, 
		$site.Url, 
		$site.LastContentModifiedDate, 
		$site.StorageUsage,
		$pnpSite.ID,
		$pnpSite.Owner.Email,
		$pnpSite.SecondaryContact.Email,
		$site.Template,
		$pnpSite.IsHubSite,
		$pnpSite.HubSiteID.ToString(),
		$pnpSite.GroupID.ToString(),
		$pnpSite.RelatedGroupID.ToString(),
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
		# TODO: Use this for app registration
		Connect-PnPOnline -ClientId $appID -Url $web.Url -Tenant $tenant -CertificatePath $certPath -CertificatePassword $certPass #>
		<# TODO: Use this for basic Authtentication	
		$pnpSiteConnect = Connect-PnPOnline -Url $web.Url -Credentials $spCred #>
		Write-Host ("      Connected and Scanning Web: {0}" -f $web.Title)

		#region 9.	Collect Web Info		
		# Get Modified Dates
		Write-Host "         Creating Web Record" -ForegroundColor DarkGray
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
			
			#region 10.	Check for Large Lists
			Write-Host "            Scanning for Large Lists & Unfiltered Views" -ForegroundColor Blue
			If ($list.ItemCount -gt $maxItemCount)
			{
				#get list views
				$listViews = Get-PnPView -List $listTitle
				#get unfiltered list views on LARGE LISTS
				$unfilteredListViews = ""
				$unfilteredListViewCount = "0"
				If ($list.ItemCount -gt 5000)
				{					
					$listViews = Get-PnPView -List $list.Title
					Foreach ($listView in $listViews)
					{
						if ($listView.ViewQuery.Contains("Where")) 
						{
							$unfilteredListViews += $listView.Title + ","
							$unfilteredListViewCount +=1
						}
					}
				}
				if ($unfilteredListViews.Length -gt 0) {$unfilteredListViews = $unfilteredListViews.Substring(0,$unfilteredListViews.Length-1)}

				#get custom content types list
				Write-Host "            Scanning for custom content types" -ForegroundColor Blue
				$contentTypes = Get-PnPContentType -List $listTitle | Where{($skipContentTypes -notcontains $_.Name)}
				$customContentTypes = ""
				foreach ($contentType in $contentTypes) {
					$customContentTypes += $contentType.Name + ","
				}
				if ($customContentTypes.Length -gt 0) {
					$customContentTypes = $customContentTypes.Substring(0,$customContentTypes.Length-1)
				}

				# Creating Large List Record
				Write-Host ("            Creating List record: {0}" -f $list.Title) -ForegroundColor Blue
				Out-File -FilePath $fileLists -Append -InputObject ("{0}|{1}|{2}|{3}|{4}|{5}|{6}|{7}|{8}|{9}|{10}|{11}|{12}|{13}|{14}|{15}|{16}|{17}|{18}|{19}" -f 
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
					$list.NoCrawl,
					$listViews.Count,
					$unfilteredListViews,
					$unfilteredListViewCount,
					$customContentTypes,
					$contentTypes.Count)
			} # If ($list.Items.Count -gt $MaxItems)
			#endregion

			#region 11.	Check for InfoPath Forms Libraries
			# Scanning for InfoPath based on Library type
			If (($list.BaseType -eq "DocumentLibrary") -and ($list.BaseTemplate -eq "115")) {
				# Creating InfoPath Record
				Write-Host ("            Creating InfoPath Library record: {0}" -f $list.Title) -ForegroundColor Blue
				Out-File -FilePath $fileInfoPath -Append -InputObject ("{0}|{1}|{2}|{3}|{4}|{5}" -f 
					$siteTitle, 
					$pnpSite.Url, 
					$webTitle,
					$web.Url, 
					$listTitle, 
					$list.RootFolder.ServerRelativePath.DecodedUrl)
			} 
			# Scanning for InfoPath based on the List's forms
			try {
				$listInfoPath = Get-PnPList -Identity $list -Includes Forms -Connection $pnpSiteConnect -ErrorAction Stop
				If ($listInfoPath.Forms.ServerRelativeURL -match "ifs.aspx")
				{
					# Creating InfoPath Record
					Write-Host ("            Creating InfoPath List Form record: {0}" -f $list.Title) -ForegroundColor Blue
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
				Write-Host ("            List doesn't support Infopath edited forms: {0}" -f $list.Title) -ForegroundColor Blue
				<#Write-Host "!! EXCEPTION OCCURRED !!" -ForegroundColor Red
				Write-Host "!! Exception Type: $($_.Exception.GetType().FullName)" -ForegroundColor Red
				Write-Host "!! Exception Message: $($_.Exception.Message)" -ForegroundColor Red #>
			}
			#endregion

			#region 12.	Check for SP Workflows
			ForEach ($workflow in $list.WorkflowAssociations)
			{
				If ($workflow -ne $null)
				{
					Write-Host ("            Creating Workflow record: {0}" -f $workflow.Name) -ForegroundColor Blue
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

			#region 13.	Checked Out Documents
			If (($list.BaseType -eq "DocumentLibrary") -and ($list.ItemCount -gt 0))
			{
				# Loop through Items in List
				Write-Host "            Scanning for checked out documents" -ForegroundColor Blue
				$items = Get-PnPListItem -Web $web -List $list -Query $queryCheckedOut -Connection $pnpSiteConnect
				ForEach ($item in $items)
				{
					Write-Host ("               Creating Checked Out record: {0}" -f $item["FileLeafRef"]) -ForegroundColor DarkYellow
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
				} #ForEach ($item in $items)
			}
			#endregion

			#region 14.	Complex Columns
			Write-Host "            Scanning for complex columns" -ForegroundColor Blue
			ForEach ($column in $colComplexTypes)
			{
				$fields = Get-PnPField -List $listTitle  | ?{($_.TypeAsString -like $column) -and ($_.Hidden -eq $false) -and ($_.Title -notin $skipColumns)}
				foreach($field in $fields)
				{
					Write-Host ("               Creating Complex Columns record: {0}" -f $field.Title) -ForegroundColor DarkYellow
					Out-File -FilePath $fileComplexColumns -Append -InputObject ("{0}|{1}|{2}|{3}|{4}|{5}|{6}|{7}" -f 
					$siteTitle, 
					$site.Url, 
					$webTitle,
					$web.Url, 
					$listTitle, 
					$list.RootFolder.ServerRelativePath.DecodedUrl, 
					$field.Title,
					$column)
				} # foreach($field in $fields)
			} # ForEach ($column in $colTypes)
			#endregion

		} # ForEach ($list in $web.Lists)
	
		# Screen Output for Web
		Write-Host ("      Web Scan Complete" -f $web.Title) -ForegroundColor Green
	} # ForEach ($web in $webs)

	# Dispose of site object
	Write-Host ("   Site Scan Complete" -f $web.Title) -ForegroundColor Green
	#Disconnect-PnPOnline -Connection $pnpSiteConnect
} # ForEach ($site in $sites)

# Screen Output for Farm
Write-Host ("Scan Complete" -f $web.Title) -ForegroundColor Green
#endregion

#region Completed Screen Output
Complete-CatScript $scriptStart[0]
Stop-Transcript
#endregion