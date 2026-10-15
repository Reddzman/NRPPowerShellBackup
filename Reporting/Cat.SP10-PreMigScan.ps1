<# Cat.SP10_PreMigScanWebApp.ps1

This script generates reports for across all site collections to identify the following information:
	-	Solution Deployment
	-	Site Collection Power Users, Modified Dates, Storage
	-	Web Power Users, Modified Dates, Storage
	-	Large Lists
	-	InfoPath Forms
	-	Workflows
	-	Complex Columns

Potential Improvements
	-	Checked Out Documents - Add Item URL to list
	-	InfoPath & Workflow - What other details can be pulled w/ PowerShell
	-	What Usage (Page load, Last Accessed, etc.) data can be collected w/ PS?

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

#region Start Date / Time
# Set PS location
$PSScriptRoot = Split-Path $MyInvocation.MyCommand.Path -Parent
Set-Location -Path $PSScriptRoot

# Get Script Parent Location
$LocSplit = $PSScriptRoot.Split("\")
$LocReplace = $LocSplit[($LocSplit.Count - 1)]
$Location = $PSScriptRoot.Replace($LocReplace,"")

# Remove Existing Catapult Modules to ensure PS session has most up to date functions
Get-Module Cat.* | Remove-Module

# Load General Functions
Import-Module ("{0}General\Cat.GeneralFunctions.psm1" -f $Location)

# Get Start Date / Time
$Start = Start-CatScript

# Start Transcript
Start-Transcript -Path ("..\_Log Files\PreMigScan-{0}.log" -f $Start[1])
#endregion

#region 1.	Input & Output Files
# File Paths
$SolutionFile = ("{0}_Output Files\Client-Solutions-{1}.txt" -f $Location, $Start[1])
$SitesFile = ("{0}_Output Files\Client-SiteCollections-{1}.txt" -f $Location, $Start[1])
$WebsFile = ("{0}_Output Files\Client-Webs-{1}.txt" -f $Location, $Start[1])
$LargeListsFile = ("{0}_Output Files\Client-LargeLists-{1}.txt" -f $Location, $Start[1])
$InfoPathFile = ("{0}_Output Files\Client-InfoPathFile-{1}.txt" -f $Location, $Start[1])
$ComplexFile = ("{0}_Output Files\Client-ComplexColumns-{1}.txt" -f $Location, $Start[1])
$WorkflowsFile = ("{0}_Output Files\Client-Workflows-{1}.txt" -f $Location, $Start[1])
$CheckedOutFile = ("{0}_Output Files\Client-CheckedOut-{1}.txt" -f $Location, $Start[1])

# Create Output File(s)
Out-File -FilePath $SolutionFile -InputObject "WebApp|WebAppURL|Site|SiteURL|Solution|CAS|GAC|WebAppResource|Deployed|ID|WebApps|Filename|Status"
Out-File -FilePath $SitesFile -InputObject "WebApp|WebAppURL|Site|SiteURL|PowerUsers|LastContentModified|Storage|ContentDB|ID|LastSecurityModified|Owner|SecondaryOwner"
Out-File -FilePath $WebsFile -InputObject "WebApp|WebAppURL|Site|SiteURL|Web|WebURL|PowerUsers|LastModified|MasterPage|CustomUpload|NoCrawl|Template"
Out-File -FilePath $LargeListsFile -InputObject "WebApp|WebAppURL|Site|SiteURL|Web|WebURL|List|ListURL|ItemCount|Type|EmailEnabled|EmailAlias|Created|ID|LastItemDeleted|LastItemModifed|MajorVersionLimit|MinorVersionLimit|NoCrawl"
Out-File -FilePath $InfoPathFile -InputObject "WebApp|WebAppURL|Site|SiteURL|Web|WebURL|List|ListURL"
Out-File -FilePath $WorkflowsFile -InputObject "WebApp|WebAppURL|Site|SiteURL|Web|WebURL|List|ListURL|Workflow"
Out-File -FilePath $CheckedOutFile -InputObject "WebApp|WebAppURL|Site|SiteURL|Web|WebURL|List|ListURL|Document|CheckedOutBy|CheckedOutSince|ContentType|SPID|UniqueID|CreatedBy|MajorVersion|MinorVersion|ModifiedBy|Created|Modified|ItemURL"
#endregion

#region 2.	Variables
# Script Variables
$Claim = "i:0#.w|"
$PowerUserRole = "Full Control"
$MaxItemCount = 0 # 5000 = SPO List view limit
$SkipLists = @("Content and Structure Reports";"Customized Reports";"Form Templates";"Master Page Gallery";"Reusable Content";"Style Library";"Workflow Tasks";)
$ComplexColumns = @("Calculated";"TaxonomyFieldType";"Lookup")
$SkipColumns = @("Item Child Count";"Folder Child Count";"App Created By";"App Modified By";"Check In Comment";"Source Version (Converted Document)";
			"Source Name (Converted Document)";"Check In Comment";"DisplayVersion")
#endregion

#region 3.	Snapin(s), Modules & Environment Connection(s)
# Load SharePoint PowerShell Snap-In(s)
Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue

# Load Custom Modules
Import-Module ("{0}Reporting\Cat.ReportingFunctions.psm1" -f $Location)
#endregion

#region 4.	Main Script
# Screen Output
Write-Host
Write-Host "Starting Farm Scan" -ForegroundColor DarkCyan

# Retrieving Farm Solutions
Write-Host "   Retrieving Farm Solutions" -ForegroundColor Magenta
$Solutions = Get-SPSolution 
ForEach ($Solution in $Solutions)
{
	# Write solution to SP reporting site
	Write-Host ("      Creating Solution record: {0}" -f $Solution.Name) -ForegroundColor DarkGreen
	$DeployedTo = ""
	ForEach ($WA in $Solution.DeployedWebApplications){
		If ($DeployedTo -eq ""){
			$DeployedTo = $WA.URL
		}
		else {
			$DeployedTo += (",{0}" -f $WA.URL)	
		}
	}
	Out-File -FilePath $SolutionFile -Append -InputObject ("{0}|{1}|{2}|{3}|{4}|{5}|{6}|{7}|{8}|{9}|{10}|{11}|{12}" -f 
		"Farm",
		"None",
		"Farm Solution", 
		"None", 
		$Solution.Name.Replace("|","\"), 
		$Solution.ContainsCASPolicy,
		$Solution.ContainsGlobalAssembly,
		$Solution.ContainsWebApplicationResource,
		$Solution.Deployed,
		$Solution.ID,
		$DeployedTo,
		$Solution.SolutionFile.Name.Replace("|","\"),
		$Solution.Status)
} # ForEach ($Solution in $Solutions)

# Get Web Application Objects
Write-Host "   Retrieving Web Applciations" -ForegroundColor Magenta	
$WebApps = Get-SPWebApplication

# Loop through all web applications in the farm
ForEach ($WebApp in $WebApps)
{
	# Screen Output for Web Application
	Write-Host ("      Scanning Web Application: {0}" -f $WebApp.URL) -ForegroundColor DarkGreen
	if ($WebApp.DisplayName -eq $null){
		$waTitle = ""
	}
	else {
		$waTitle = $WebApp.DisplayName.Replace("|","\")
	}
	
	# Get site collection object
	Write-Host "         Retrieving Site Collections"
	$Sites = Get-SPSite -WebApplication $WebApp.URL -Limit all

	# Loop through site collections for web application
	ForEach ($Site in $Sites)
	{
		# Screen Output for Site
		Write-Host ("         Scanning site collection: {0}" -f $Site.URL)

		#region 5.	Collect Site Collection Info
		# Get Power Users
		Write-Host "            Retrieving Site Power Users" -ForegroundColor DarkGray
		$PowerUsers = Get-CatSP10PowerUsers -Assignments $Site.RootWeb.RoleAssignments -Level $PowerUserRole
		
		# Output Site Collection Information
		Write-Host "            Creating Site Record" -ForegroundColor DarkGray
		if ($Site.RootWeb.Title -eq $null){
			$siteTitle = ""
		}
		else {
			$siteTitle = $Site.RootWeb.Title.Replace("|","\")
		}
		if ($PowerUsers -eq $null){
			$sitePowerUsers = ""
		}
		else {
			$sitePowerUsers = $PowerUsers.Replace("|","\")
		}
		if ($Site.ContentDatabase.Name -eq $null){
			$siteContentDatabase = ""
		}
		else {
			$siteContentDatabase = $Site.ContentDatabase.Name.Replace("|","\")
		}
		if ($Site.Owner.UserLogin -eq $null){
			$siteOwner = ""
		}
		else {
			$siteOwner = $Site.Owner.UserLogin.Replace($Claim,"").Replace("|","\")
		}
		if ($Site.SecondaryContact.UserLogin -eq $null){
			$SecondaryOwner = ""
		}
		else {
			$SecondaryOwner = $Site.SecondaryContact.UserLogin.Replace($Claim,"").Replace("|","\")
		}
		Out-File -FilePath $SitesFile -Append -InputObject ("{0}|{1}|{2}|{3}|{4}|{5}|{6}|{7}|{8}|{9}|{10}|{11}" -f 
			$waTitle,
			$WebApp.Url,
			$siteTitle, 
			$Site.Url, 
			$sitePowerUsers, 
			$Site.LastContentModifiedDate, 
			$Site.Usage.Storage,
			$siteContentDatabase,
			$Site.ID,
			$Site.LastSecurityModifiedDate,
			$siteOwner,
			$SecondaryOwner)
		
		# Get Sandbox Solutions
		Write-Host "            Retrieving Sandbox Solutions" -ForegroundColor DarkGray
		$SandboxSolutions = Get-SPUserSolution -Site $Site
		ForEach ($Solution in $SandboxSolutions)
		{
			# Write solution to SP reporting site
			If ($Solution -ne $null)
			{
				Write-Host ("               Creating Sandbox Solution record: {0}" -f $Solution.Name) -ForegroundColor Yellow
				Out-File -FilePath $SolutionFile -Append -InputObject ("{0}|{1}|{2}|{3}|{4}|{5}|{6}|{7}|{8}|{9}|{10}|{11}|{12}" -f 
					$waTitle,
					$WebApp.Url,
					$siteTitle, 
					$Site.Url, 
					$Solution.Name.Replace("|","\"), 
					$Solution.ContainsCASPolicy,
					$Solution.ContainsGlobalAssembly,
					$Solution.ContainsWebApplicationResource,
					$Solution.Deployed,
					$Solution.ID,
					"Sandbox Solution",
					$Solution.SolutionFile.Name.Replace("|","\"),
					$Solution.Status)
			}
		} # ForEach ($Solution in $SandboxSolutions)
		#endregion

		# Loop through sub sites
		ForEach ($Web in $Site.AllWebs)
		{
			# Screen output for web
			Write-Host ("               Scanning Web: {0}" -f $Web.Title) -ForegroundColor Yellow

			#region 6.	Collect Web Info
			# Get Power Users & AD groups for web
			If ($Web.HasUniqueRoleAssignments)
			{
				Write-Host "                  Retrieving Web Power Users" -ForegroundColor DarkRed
				$PowerUsers = Get-CatSP10PowerUsers -Assignments $Web.RoleAssignments -Level $PowerUserRole
			} # If ($Web.HasUniqueRoleAssignments)
			
			# Get Modified Dates
			Write-Host "                  Creating Web Record" -ForegroundColor DarkRed
			if ($Web.Title -eq $null){
				$webTitle = ""
			}
			else {
				$webTitle = $Web.Title.Replace("|","\")
			}
			if ($PowerUsers -eq $null){
				$webPowerUsers = ""
			}
			else {
				$webPowerUsers = $PowerUsers.Replace("|","\")
			}
			Out-File -FilePath $WebsFile -Append -InputObject ("{0}|{1}|{2}|{3}|{4}|{5}|{6}|{7}|{8}|{9}|{10}|{11}" -f 
				$waTitle,
				$WebApp.Url,
				$siteTitle, 
				$Site.Url, 
				$webTitle,
				$Web.Url, 
				$webPowerUsers, 
				$Web.LastItemModifiedDate,
				$Web.MasterURL,
				$web.CustomUploadPage,
				$Web.NoCrawl,
				$Web.WebTemplate)
			#endregion
			
			# Loop through all lists and libraries in a sub site
			$Lists = $Web.Lists | ?{($_.Hidden -eq $false) -and ($SkipLists -notcontains $_.Title)}
			ForEach ($List in $Lists)
			{
				# Screen output for list
				Write-Host ("                  Scanning List: {0}" -f $List.Title) -ForegroundColor DarkRed
				$ListURL = ("{0}/{1}" -f $Site.Url, $List.RootFolder.Url)
				
				# 7.	Check for Large Lists
				If ($List.ItemCount -gt $MaxItemCount)
				{
					# Creating Large List Record
					Write-Host ("                     Creating List record: {0}" -f $List.Title) -ForegroundColor DarkRed
					if ($List.Title -eq $null){
						$listTitle = ""
					}
					else {
						$listTitle = $List.Title.Replace("|","\")
					}
					Out-File -FilePath $LargeListsFile -Append -InputObject ("{0}|{1}|{2}|{3}|{4}|{5}|{6}|{7}|{8}|{9}|{10}|{11}|{12}|{13}|{14}|{15}|{16}|{17}|{18}" -f 
						$waTitle,
						$WebApp.Url,
						$siteTitle, 
						$Site.Url, 
						$webTitle,
						$Web.Url, 
						$listTitle, 
						$ListURL, 
						$List.ItemCount,
						$List.BaseType,
						$List.CanReceiveEmail,
						$List.EmailAlias,
						$List.Created,
						$List.ID,
						$List.LastItemDeletedDate,
						$List.LastItemModifiedDate,
						$List.MajorVersionLimit,
						$List.MajorWithMinorVersionsLimit,
						$List.NoCrawl)
				} # If ($List.Items.Count -gt $MaxItems)

				# 8.	Check for InfoPath Forms Libraries
				If ((($List.ContentTypes -ne $null) -and ($List.ContentTypes[0].ResourceFolder.Properties["_ipfs_infopathenabled"])) -or (($List.BaseType -eq "DocumentLibrary") -and ($List.BaseTemplate -eq "XMLForm")))
				{
					# Creating InfoPath Record
					Write-Host ("                     Creating InfoPath record: {0}" -f $List.Title) -ForegroundColor DarkRed
					Out-File -FilePath $InfoPathFile -Append -InputObject ("{0}|{1}|{2}|{3}|{4}|{5}|{6}|{7}" -f 
						$waTitle,
						$WebApp.Url,
						$siteTitle, 
						$Site.Url, 
						$webTitle,
						$Web.Url, 
						$listTitle, 
						$ListURL)
				} # If (InfoPath)

				# 9.	Check for Workflows
				ForEach ($Workflow in $List.WorkflowAssociations)
				{
					If ($Workflow -ne $null)
					{
						Write-Host ("                     Creating Workflow record: {0}" -f $Workflow.Name) -ForegroundColor DarkRed
						if ($Workflow.Name -eq $null){
							$wfName = ""
						}
						else {
							$wfName = $Workflow.Name.Replace("|","\")
						}
						Out-File -FilePath $WorkflowsFile -Append -InputObject ("{0}|{1}|{2}|{3}|{4}|{5}|{6}|{7}|{8}" -f 
							$waTitle,
							$WebApp.Url,
							$siteTitle, 
							$Site.Url, 
							$webTitle,
							$Web.Url, 
							$listTitle, 
							$ListURL, 
							$wfName)
					}
				} # ForEach ($Workflow in $List.WorkflowAssociations)

				# 10.	Checked Out Documents
				If ($List.BaseType -eq "DocumentLibrary")
				{
					# Loop through Items in List
					Write-Host "                     Scanning for checked out documents" -ForegroundColor Black
					ForEach ($Item in $List.Items)
					{
						# Get Checked Out Documents
						If ($Item.File.CheckOutStatus -ne "None")
						{
							Write-Host ("                     Creating Checked Out record: {0}" -f $Item.Name) -ForegroundColor Gray
							if ($Item.File.Author.UserLogin -eq $null){
								$Author = ""
							}
							else {
								$Author = $Item.File.Author.UserLogin.Replace($Claim,"").Replace("|","\")
							}
							if ($Item.File.ModifiedBy.UserLogin -eq $null){
								$modified = ""
							}
							else {
								$modified = $Item.File.ModifiedBy.UserLogin.Replace($Claim,"").Replace("|","\")
							}
							Out-File -FilePath $CheckedOutFile -Append -InputObject ("{0}|{1}|{2}|{3}|{4}|{5}|{6}|{7}|{8}|{9}|{10}|{11}|{12}|{13}|{14}|{15}|{16}|{17}|{18}|{19}|{20}" -f 
								$waTitle,
								$WebApp.Url,	
								$siteTitle, 
								$Site.Url, 
								$webTitle,
								$Web.Url, 
								$listTitle, 
								$ListURL, 
								$Item.Name.Replace("|","\"), 
								$Item.File.CheckedOutBy.UserLogin.Replace($Claim,"").Replace("|","\"), 
								$Item.File.CheckedOutDate,
								$Item.ContentType.Name.Replace("|","\"),
								$Item.ID,
								$Item.UniqueId,
								$Author,
								$Item.File.MajorVersion,
								$Item.File.MinorVersion,
								$modified,
								$Item.File.TimeCreated,
								$Item.File.TimeLastModified,
								$Item.File.Url)
						} # If (($List.BaseType -eq "DocumentLibrary") -and $Item.File.CheckOutStatus -ne "None")
					}
				}
			} # ForEach ($List in $Web.Lists)
				
			# Screen Output for Web
			Write-Host ("               Web Scan Complete" -f $Web.Title) -ForegroundColor Green
			$Web.Dispose()
		} # ForEach ($Web in $Webs)

		# Dispose of site object
		Write-Host ("            Site Scan Complete" -f $Web.Title) -ForegroundColor Green
		$Site.Dispose()
	} # ForEach ($Site in $Sites)

	# Screen Output for Web Application
	Write-Host ("         Web Application Scan Complete" -f $Web.Title) -ForegroundColor Green

} # ForEach ($WebApp in $WebApps)

# Screen Output for Farm
Write-Host ("      Farm Scan Complete" -f $Web.Title) -ForegroundColor Green
#endregion

#region Completed Screen Output
Complete-CatScript $Start[0]
Stop-Transcript
#endregion