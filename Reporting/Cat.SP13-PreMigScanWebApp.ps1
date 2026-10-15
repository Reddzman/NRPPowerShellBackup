<# Cat.SP13_PreMigraitonReport.ps1

This script generates reports for across all site collections to identify the following information
	-	Solution Deployment
	-	Site Collection Power Users, Modified Dates, Storage
	-	Web Power Users, Modified Dates, Storage
	-	AD Group utilization
	-	Large Lists
	-	InfoPath Forms
	-	Workflows
	-	Complex Columns

Potential Enhancements:
	-	Seperate Script for AD Group utilization / write to CSV / upload to doc lib
	-	Have Item output distiguish between docs and items: Title blank on docs
	-	Required Columns / Items missing required columns
	-	Skip Lists (Workflow history)
	-	Web counts in site collection record
	-	List Counts in Web List
	-	Output all lists & item counts	

Table of Contents
	1.	Variables
	3.	Snapin(s), Modules & Environment Connection(s)
	4.	Main Script
	5.	Collect Site Collection Info
	6.	Collect Web Info
	7.	Check for Large Lists
	8.	Check for InfoPath Forms Libraries & Lists
	9.	Check for Complex Columns
	10.	Check for Workflows
#>

#region Script Parameters
Param(
	[string] $client = "PR",
	[String] $spURL = "https://sharepoint.client.com",
	[String] $spFarmSQLInstance = "SQLAlias",
	[String] $nintexSQLInstance = ""
)
#endregion

#region Start Date / Time
# Remove Existing Catapult Modules to ensure PS session has most up to date functions
Get-Module Cat.* | Remove-Module

$global:currDir = (Split-Path $MyInvocation.MyCommand.Definition)
$currentLocation = $currDir.Split('\')
$scriptName = $MyInvocation.MyCommand.Path
Push-Location $global:currDir

# Load General Functions
Import-Module "..\General\Cat.GeneralFunctions.psm1"

# Get Start Date / Time
$scriptStart = Start-CatScript

# Start Transcript
Start-Transcript -Path ("..\_Log FIles\{0}-PreMigration-{1}.log" -f $client, $scriptStart[1])
#endregion

#region 1.	Variables
# User Input Variables
Write-Host
Write-Host "User Input Variables" -ForegroundColor Cyan
Write-Host "   Enter credentials for SharePoint:"
if(!($global:spCred)){$global:spCred = Get-Credential -Message "Enter your credentials for SharePoint site:" }
if(!($ctxPassword)){$ctxPassword = Read-Host -Prompt "Enter password again for CTX interface to large lists" }

# Script Variables
$spEnvironment = "SP13"
$powerRole = "Full Control"
$itemCount = 4000 # 5000 = SPO List view limit
$skipLists = @("Content and Structure Reports";"Customized Reports";"Workflow Tasks";"Style Library";)
$colComplexTypes = @("Calculated";"TaxonomyFieldType";"Lookup")
$skipColumns = @("Item Child Count";"Folder Child Count";"App Created By";"App Modified By";"Check In Comment";"Source Version (Converted Document)";
			"Source Name (Converted Document)";"Check In Comment";"DisplayVersion")

# Define path to Output files
$clientFolder = ("..\_Output Files\{0}" -f $client);
$runPathVar = $scriptStart[1];
$outputFolder = ("$clientFolder\{0}" -f $runPathVar);
$lastOutputFolder = $outputFolder;

if(Test-Path $clientFolder){ 
    if( ($null -ne (gci $clientFolder)) -and 
        ((gci $clientFolder).count -gt 0) 
    ){
        $lastRunPathVar = (gci $clientFolder | ? { $_.PSIsContainer } | sort CreationTime)[-1].Name;
        $lastOutputFolder = ("$clientFolder\{0}" -f $lastRunPathVar);
    }
}else{
    mkdir $clientFolder | Out-Null;
}
mkdir $outputFolder | Out-Null;

$adPath = ("$outputFolder\ADReport.csv")
$dbPath = ("$outputFolder\Databases.csv")
$coPath = ("$outputFolder\CheckedOut.csv")
$colPath = ("$outputFolder\Columns.csv")
$ctTypePath = ("$outputFolder\ContentTypes.csv")
$ipPath = ("$outputFolder\InfoPath.csv")
$itemPath = ("$outputFolder\Items.csv")
$listPath = ("$outputFolder\Lists.csv")
$sitePath = ("$outputFolder\SiteCollections.csv")
$solPath = ("$outputFolder\Solutions.csv")
$webPath = ("$outputFolder\Webs.csv")
$wfPath = ("$outputFolder\Workflows.csv")


# Output Parameters & Variables
Write-Host
Write-Host "Script Parameters" -ForegroundColor Cyan
Write-Host ("   Web Applciation: {0}" -f $spURL)
Write-Host
Write-Host "User Inputs" -ForegroundColor Cyan
Write-Host ("   Credentials: {0}" -f $global:spCred.UserName)
Write-Host
Write-Host "Script Variables" -ForegroundColor Cyan
Write-Host ("   Power User Permission: {0}" -f $powerRole)
Write-Host ("   Item Count: {0}" -f $itemCount)
Write-Host ("   Complex Column Types: {0}" -f $colComplexTypes)
Write-Host ("   Skip Lists: {0}" -f $skipLists)
Write-Host ("   Skip Columns: {0}" -f $skipColumns)
Write-Host
Write-Host "Output Files" -ForegroundColor Cyan
Write-Host ("   Solutions: {0}" -f $solPath)
Write-Host ("   Sites: {0}" -f $sitePath)
Write-Host ("   Webs: {0}" -f $webPath)
Write-Host ("   AD Groups: {0}" -f $adPath)
Write-Host ("   Lists: {0}" -f $listPath)
Write-Host ("   InfoPath: {0}" -f $ipPath)
Write-Host ("   Workflow: {0}" -f $wfPath)
Write-Host ("   Complex Columns: {0}" -f $colPath)
#endregion

#region 2.	Input & Output Files

	# Create Output File(s)
	if(Test-Path -Path "$lastOutputFolder\ADReport.csv"){Copy-Item -Path "$lastOutputFolder\ADReport.csv" -Destination $outputFolder}else{
		Out-File -FilePath $adPath 		-InputObject "Site|SiteURL|Web|WebURL|List|ListURL|Item|Group" -Encoding unicode
    }
	if(Test-Path -Path "$lastOutputFolder\CheckedOut.csv"){Copy-Item -Path "$lastOutputFolder\CheckedOut.csv" -Destination $outputFolder}else{
		Out-File -FilePath $coPath 		-InputObject "Site|SiteURL|Web|WebURL|List|ListURL|Document|ContentType|SPID|UniqueID|CreatedBy|MajorVersion|MinorVersion|ModifiedBy|Created|Modified|ItemURL|ListGUID|CheckedOutBy|CheckedOutSince" -Encoding unicode
    }
	if(Test-Path -Path "$lastOutputFolder\Columns.csv"){Copy-Item -Path "$lastOutputFolder\Columns.csv" -Destination $outputFolder}else{
		Out-File -FilePath $colPath 	-InputObject "Site|SiteURL|Web|WebURL|List|ListURL|ContentType|Column|Type|Scope|ComplexType" -Encoding unicode
    }
	 if(Test-Path -Path "$lastOutputFolder\ContentTypes.csv"){Copy-Item -Path "$lastOutputFolder\ContentTypes.csv" -Destination $outputFolder}else{
         Out-File -FilePath $ctTypePath 	-InputObject "SiteGUID|Site|SiteURL|WebGUID|Web|WebURL|ListGUID|List|ListURL|ContentTypeID|ContentType|Parent|Scope" -Encoding unicode
     }
	if(Test-Path -Path "$lastOutputFolder\Databases.csv"){Copy-Item -Path "$lastOutputFolder\Databases.csv" -Destination $outputFolder}else{
	    Out-File -FilePath $dbPath 		-InputObject "Site|SiteURL|Web|WebURL|List|ListURL" -Encoding unicode
    }
	if(Test-Path -Path "$lastOutputFolder\InfoPath.csv"){Copy-Item -Path "$lastOutputFolder\InfoPath.csv" -Destination $outputFolder}else{
	    Out-File -FilePath $ipPath 		-InputObject "Site|SiteURL|Web|WebURL|List|ListURL" -Encoding unicode
    }
	 if(Test-Path -Path "$lastOutputFolder\Items.csv"){Copy-Item -Path "$lastOutputFolder\Items.csv" -Destination $outputFolder}else{
	 	Out-File -FilePath $itemPath	-InputObject "Site|SiteURL|Web|WebURL|ListGUID|List|ListURL|ItemName|ContentType|SPID|UniqueID|CreatedBy|MajorVersion|MinorVersion|ModifiedBy|Created|Modified|ItemURL" -Encoding unicode
     }
	 if(Test-Path -Path "$lastOutputFolder\Lists.csv"){Copy-Item -Path "$lastOutputFolder\Lists.csv" -Destination $outputFolder}else{
	     Out-File -FilePath $listPath 	-InputObject "Site|SiteURL|WebGUID|Web|WebURL|ListGUID|List|ListURL|ItemCount|Type|Created|LastItemDeleted|LastItemModifed|MajorVersionLimit|MinorVersionLimit|NoCrawl|EmailEnabled|EmailAlias" -Encoding unicode
     }
	 if(Test-Path -Path "$lastOutputFolder\SiteCollections.csv"){Copy-Item -Path "$lastOutputFolder\SiteCollections.csv" -Destination $outputFolder}else{
	 	Out-File -FilePath $sitePath 	-InputObject "SiteGUID|Site|SiteURL|LastContentModified|Storage|ID|Owner|SecondaryOwner|Template|ContentDB" -Encoding unicode
     }
	if(Test-Path -Path "$lastOutputFolder\Solutions.csv"){Copy-Item -Path "$lastOutputFolder\Solutions.csv" -Destination $outputFolder}else{
		Out-File -FilePath $solPath 	-InputObject "Site|SiteURL|Solution|CAS|GAC|WebAppResource|Deployed|ID|WebApps|Filename|Status" -Encoding unicode
    }
	 if(Test-Path -Path "$lastOutputFolder\Webs.csv"){Copy-Item -Path "$lastOutputFolder\Webs.csv" -Destination $outputFolder}else{
	 	Out-File -FilePath $webPath 	-InputObject "SiteGUID|Site|SiteURL|WebGUID|Web|WebURL|LastModified|MasterPage|NoCrawl|Template|CustomUpload|PowerUsers" -Encoding unicode
     }
	if(Test-Path -Path "$lastOutputFolder\Workflows.csv"){Copy-Item -Path "$lastOutputFolder\Workflows.csv" -Destination $outputFolder}else{
		Out-File -FilePath $wfPath 		-InputObject "Site|SiteURL|Web|WebURL|ListGUID|List|ListURL|Workflow|TaskList|HistoryList|Created|Modified" -Encoding unicode
    }

    $qCheckedOut = New-Object System.Collections.Queue; Import-Csv -Delimiter "|" -Path $coPath | %{ $qCheckedOut.Enqueue($_) };
    $qColumns = New-Object System.Collections.Queue; Import-Csv -Delimiter "|" -Path $colPath | %{ $qColumns.Enqueue($_) };
    $qContentTypes = New-Object System.Collections.Queue; Import-Csv -Delimiter "|" -Path $ctTypePath | %{ $qContentTypes.Enqueue($_) };
    $qInfoPath = New-Object System.Collections.Queue; Import-Csv -Delimiter "|" -Path $ipPath | %{ $qInfoPath.Enqueue($_) };
    $qItems = New-Object System.Collections.Queue; Import-Csv -Delimiter "|" -Path $itemPath | %{ $qItems.Enqueue($_) };
    $qLists = New-Object System.Collections.Queue; Import-Csv -Delimiter "|" -Path $listPath | %{ $qLists.Enqueue($_) };
    $qSiteCollections = New-Object System.Collections.Queue; Import-Csv -Delimiter "|" -Path $sitePath | %{ $qSiteCollections.Enqueue($_) };
    $qSolutions = New-Object System.Collections.Queue; Import-Csv -Delimiter "|" -Path $solPath | %{ $qSolutions.Enqueue($_) };
    $qWebs = New-Object System.Collections.Queue; Import-Csv -Delimiter "|" -Path $webPath | %{ $qWebs.Enqueue($_) };
    $qWorkflows = New-Object System.Collections.Queue; Import-Csv -Delimiter "|" -Path $wfPath | %{ $qWorkflows.Enqueue($_) };



#endregion

#region 3.	Snapin(s), Modules & Environment Connection(s)
# Load SharePoint PowerShell Snap-In(s)
Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue

# Load Custom Modules
Import-Module ".\Cat.ReportingFunctions.psm1"

#Load SharePoint CSOM Assemblies
# SP13 Assemblies
Add-Type -Path "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll"
Add-Type -Path "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Runtime.dll" #>
<# SP16 Assemblies
Add-Type -Path "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll"
Add-Type -Path "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Runtime.dll" #>
#endregion

#region 4.	Main Script
# Screen Output
Write-Host
Write-Host "Starting Environment Scan" -ForegroundColor Cyan

	#region Gather Farm Database information from SQL
		# Retrieving Farm databases
		Write-Host
	    Write-Host "   Retrieving Farm Database info from SQL" -ForegroundColor Magenta
	    $databases = Get-SPDatabase

	    #Set the parameters for the environment
	    $srcSQLSrvInst = $spFarmSQLInstance
	    $sqlConnTimeout = 10800

	    [Void][System.Reflection.Assembly]::LoadWithPartialName("System.Data")
	    $srcConn = new-object System.Data.SqlClient.SQLConnection
	    $srcConnString = "Integrated Security=True;Server={0};Connection Timeout={1}" -f $srcSQLSrvInst, $sqlConnTimeout
	    $srcConn.ConnectionString = $srcConnString

	    #Connect to the SQL Server and get the databases
	    $srcConn.Open()

	    #Process all found databases
	    $srcQuery = ("
		    WITH fs
		    AS
		    (
			    SELECT database_id, TYPE, SIZE * 8.0 / 1024 SIZE
			    FROM sys.master_files
		    )
		    SELECT 
			    name,
			    (SELECT SUM(SIZE) FROM fs WHERE TYPE = 0 AND fs.database_id = db.database_id) DataFileSizeMB,
			    (SELECT SUM(SIZE) FROM fs WHERE TYPE = 1 AND fs.database_id = db.database_id) LogFileSizeMB
		    FROM sys.databases db
	    ")

	    $srcCmd = New-Object System.Data.SqlClient.SqlCommand($srcQuery, $srcConn)
	    $srcCmd.CommandTimeout=$sqlConnTimeout

	    $ds= New-Object System.Data.Dataset
	    $da= New-Object System.Data.SqlClient.SqlDataAdapter($srcCmd)
	    [void]$da.fill($ds)
	    $filteredRows = $ds.Tables[0] | ?{
		    ($_.name -ne "master") -and 
		    ($_.name -ne "tempdb") -and 
		    ($_.name -ne "model") -and 
		    ($_.name -ne "msdb")
	    } | Select-Object @('name','DataFileSizeMB','LogFileSizeMB') 


	    #Disconnect from the SQL Server database
	    $srcConn.Close()

	    ForEach ($db in $databases)
	    {
		    if($db.Type -match 'Content Database'){
			    $contentDBCheck = "Yes"
		    } else{
			    $contentDBCheck = "No"
		    }
		    # Write db to SP reporting site
		    Write-Host ("      Creating Database record: {0}" -f $db.Name)
		    #"Database Name|DB Size (MB)|Log Size (MB)|Environment|Content DB"
		    Out-File -FilePath $dbPath -Append -InputObject ("{0}|{1}|{2}|{3}|{4}" -f $db.Name, ($filteredRows |?{$_.name -eq $db.Name} | Select DataFileSizeMB).DataFileSizeMB, ($filteredRows |?{$_.name -eq $db.Name} | Select LogFileSizeMB).LogFileSizeMB, $spEnvironment, $contentDBCheck)
	    } 
	#endregion

	#region Retrieving Farm Solutions
		Write-Host 
		Write-Host "   Retrieving Farm Solutions" -ForegroundColor Magenta
		$solutions = Get-SPSolution 
		ForEach ($solution in $solutions)
		{
			# Write solution to SP reporting site
			Write-Host ("      Creating Solution record: {0}" -f $solution.Name)
			#"Site|SiteURL|Solution|CAS|GAC|WebAppResource|Deployed|ID|WebApps|Filename|Status"
			Out-File -FilePath $solPath -Append -InputObject ("{0}|{1}|{2}|{3}|{4}|{5}|{6}|{7}|{8}|{9}|{10}" -f "Farm Solution", "None", $solution.Name, $solution.ContainsCasPolicy, $solution.ContainsGlobalAssembly, , $solution.ContainsWebApplicationResource, $solution.Deployed, $solution.ID.GUID, $solution.DeployedWebApplications, $solution.SolutionFile.Name, $solution.Status)
		}
	#endregion

	#region Gather Nintex workflows if exist
		if( ($nintexSQLInstance) -and (null -ne $nintexSQLInstance) -and ($nintexSQLInstance -ne '') ){
			Write-Host 
			Write-Host "   Retrieving Nintex Workflows" -ForegroundColor Magenta

			# Loading SharePoint and Nintex Objects into the PS session
			[void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint")
			[void][System.Reflection.Assembly]::LoadWithPartialName("Nintex.Workflow")
			[void][System.Reflection.Assembly]::LoadWithPartialName("Nintex.Workflow.SupportConsole")
			[void][System.Reflection.Assembly]::LoadWithPartialName("Nintex.Workflow.Administration")
			[void][System.Reflection.Assembly]::LoadWithPartialName("Nintex.Forms.SharePoint.Administration")

			# Grab Nintex Config database name

			$CFGDB = [Nintex.Workflow.Administration.ConfigurationDatabase]::OpenConfigDataBase().Database

			# Creating instance of .NET SQL client
			$cmd = New-Object -TypeName System.Data.SqlClient.SqlCommand

			$cmd.CommandType = [System.Data.CommandType]::Text

			# Begin SQL Query
			$cmd.CommandText = "SELECT
							i.WorkflowName,
							i.SiteID,
							i.WebID,
							i.listid,
							pw.Author
			FROM dbo.WorkflowInstance I
			inner join WorkflowProgress P
							ON I.InstanceID = P.InstanceID
			Inner join [$CFGDB].dbo.publishedworkflows pw
			on i.WorkflowID = pw.WorkflowId
			GROUP BY GROUPING SETS((i.siteid, i.webid, i.listid, i.workflowname, pw.Author), ());"

			$indexes = @()

			# Call to find all Nintex Content Databases in the Nintex Configuration Database, then execute the above query against each. 
			foreach ($database in [Nintex.Workflow.Administration.ConfigurationDatabase]::GetConfigurationDatabase().ContentDatabases)
			{
				$reader = $database.ExecuteReader($cmd)

				# Creating a table
				while($reader.Read())
				{
					$row = New-Object System.Object

					if(![string]::IsNullOrEmpty($reader["SiteID"])){
						$Site = $(Get-SPSite -identity $reader["SiteID"] -ErrorAction SilentlyContinue)
					}

					if(![string]::IsNullOrEmpty($reader["WebID"])){
						$SubSite = $Site.Allwebs[[Guid]"$($reader["WebID"])"]
					}
					if(![string]::IsNullOrEmpty($reader["ListID"])){
						$List = $SubSite.Lists[[Guid]"$($reader["ListID"])"]
					}

					#Adding Query results to table object
					$row | Add-Member -MemberType NoteProperty -Name "Site" -Value $Site.RootWeb.Title -ErrorAction SilentlyContinue
					$row | Add-Member -MemberType NoteProperty -Name "SiteURL" -Value $Site.Url -ErrorAction SilentlyContinue
					$row | Add-Member -MemberType NoteProperty -Name "Web" -Value $SubSite -ErrorAction SilentlyContinue
					$row | Add-Member -MemberType NoteProperty -Name "WebURL" -Value $SubSite.URL -ErrorAction SilentlyContinue
					$row | Add-Member -MemberType NoteProperty -Name "List" -Value $List.Title -ErrorAction SilentlyContinue
					$row | Add-Member -MemberType NoteProperty -Name "ListURL" -Value $List.DefaultViewUrl -ErrorAction SilentlyContinue
					$row | Add-Member -MemberType NoteProperty -Name "Workflow" -Value $reader["WorkflowName"] -ErrorAction SilentlyContinue
					$row | Add-Member -MemberType NoteProperty -Name "Author" -Value $reader["Author"] -ErrorAction SilentlyContinue
					$row | Add-Member -MemberType NoteProperty -Name "ListID" -Value $reader["ListID"] -ErrorAction SilentlyContinue
					$row | add-member -MemberType NoteProperty -Name "Database" -value $Site.ContentDatabase.Name -ErrorAction SilentlyContinue
					$row | add-member -MemberType NoteProperty -Name "DatabaseID" -value $Site.ContentDatabase.ID.Guid -ErrorAction SilentlyContinue

					$indexes += $row
				}
			}

			#Capture results to Reporting site
			ForEach ($workflow in $indexes)
			{
				Write-Host ("      Creating Workflow record: {0}" -f $workflow.Workflow)
				#"Site|SiteURL|Web|WebURL|ListGUID|List|ListURL|Workflow|TaskList|HistoryList|Created|Modified"
				# Out-File -FilePath $wfPath -Append -InputObject ("{0}|{1}|{2}|{3}|{4}|{5}|{6}|{7}|{8}|{9}|{10}" -f 
				Out-File -FilePath $wfPath -Append -InputObject ("{0}|{1}|{2}|{3}|{4}|{5}|{6}|{7}" -f 
					$workflow.Site, 
					$workflow.SiteURL, 
					$workflow.Web, 
					$workflow.WebURL, 
					$workflow.ListID, 
					$workflow.List, 
					$workflow.ListURL, 
					$workflow.Workflow)
					# $workflow.ListID,
					# $workflow.TaskListTitle,
					# $workflow.HistoryListTitle,
					# $workflow.Created,
					# $workflow.Modified)

			} # ForEach ($workflow in $indexes)
		}
	#endregion

	#region Get site collection object
		Write-Host "   Retrieving Site Collections"
		$sites = Get-SPSite -WebApplication $webAppURL -Limit All

		#region Loop through site collections for web application
			ForEach ($site in $sites)
			{
			# if($site.Url -match '/sites/finance'){
			#     #$true | Out-Null;
			# }	
				# Screen Output for Site
				Write-Host
				Write-Host ("   Scanning all sites for site collection: {0}" -f $site.RootWeb.Title) -ForegroundColor Magenta

				#region 5.	Collect Site Collection Info
					#region Get Sandbox Solutions
						Write-Host "      Retrieving Sandbox Solutions"
						$sandboxSolutions = Get-SPUserSolution -Site $site
						ForEach ($solution in $sandboxSolutions)
						{
							# Write solution to SP reporting site
							Write-Host ("         Creating Sandbox Solution record in Reporting Site: {0}" -f $solution.Name)
							#"Site|SiteURL|Solution|CAS|GAC|WebAppResource|Deployed|ID|WebApps|Filename|Status"
							Out-File -FilePath $solPath -Append -InputObject ("{0}|{1}|{2}|{3}|{4}|{5}|{6}|{7}|{8}|{9}|{10}" -f 
								$site.RootWeb.Title, 
								$site.Url, 
								$solution.Name, 
								"N/A", 
								"N/A", 
								"N/A", 
								"N/A", 
								$solution.SolutionID.GUID, 
								$site.URL, 
								$solution.Name, 
								$solution.Status)
						} # ForEach ($solution in $sandboxSolutions)
					#endregion

				#region Get Power Users
					Write-Host "      Retrieving Site Power Users"
					$powerUsers = Get-CatPowerUsers -Assignments $site.RootWeb.RoleAssignments -Level $powerRole
				#endregion
				#region Get Modified Dates, Storage
				$siteGUID = $site.Id.Guid;
				$siteUrl = $site.Url
				if ($site.RootWeb.Title -eq $null){
					$siteTitle = ""
				}
				else {
					$siteTitle = $site.RootWeb.Title.Replace("|","\")
				}
				Write-Host "      Creating Site Record in Reporting Site"
				if(!($qSiteCollections) -or ($qSiteCollections.count -eq 0) -or !($qSiteCollections.SiteGUID.Contains($siteGUID))){
					#"SiteGUID|Site|SiteURL|LastContentModified|Storage|ID|Owner|SecondaryOwner|Template|ContentDB" 
					Out-File -FilePath $sitePath -Append -InputObject ("{0}|{1}|{2}|{3}|{4}|{5}|{6}|{7}|{8}|{9}" -f 
						$siteGUID,
						$siteTitle, 
						$siteUrl, 
						$site.LastContentModifiedDate, 
						$site.Usage.Storage, 
						$siteGUID,
						$site.Owner.DisplayName, 
						$site.SecondaryOwner.DisplayName, 
						($site.RootWeb.WebTemplate + '#' + $site.RootWeb.Configuration),
						# $site.LastSecurityModifiedDate, 
						# $powerUsers, 
						$site.ContentDatabase.Name)
				}

			#endregion

			#region Loop through sub sites
				ForEach ($web in $site.AllWebs)
				{
					$webGUID = $web.Id.Guid;
					$webUrl = $web.Url;
							if ($web.Title -eq $null){
					$webTitle = ""
				}
							else {
					$webTitle = $web.Title.Replace("|","\")
				}
					# Screen output for web
					Write-Host ("      Scanning Web: {0}" -f $web.Title) -ForegroundColor DarkYellow

					#region 6.	Collect Web Info
						#region Get Power Users & AD groups for web
							If ($web.HasUniqueRoleAssignments)
							{
								Write-Host "         Retrieving Web Power Users"
								$powerUsers = Get-CatPowerUsers -Assignments $web.RoleAssignments -Level $powerRole

								Write-Host "         Retrieving AD groups"
								Get-CatSP10ADGroups -File $adPath -Assignments $Web.RoleAssignments -SiteName $Site.RootWeb.Title -SiteUrl $Site.Url -WebName $Web.Title -WebUrl $Web.Url
							} 
						#endregion
					
						#region Get Modified Dates, Storage
							Write-Host "         Creating Web Record in Reporting Site"
							if(!($qWebs) -or ($qWebs.count -eq 0) -or !($qWebs.WebGUID.Contains($webGUID))){
								#"SiteGUID|Site|SiteURL|WebGUID|Web|WebURL|LastModified|MasterPage|NoCrawl|Template|CustomUpload"
								Out-File -FilePath $webPath -Append -InputObject ("{0}|{1}|{2}|{3}|{4}|{5}|{6}|{7}|{8}|{9}|{10}|{11}" -f 
									$siteGUID,
									$siteTitle, 
									$siteUrl, 
									$webGUID,
									$webTitle, 
									$webUrl, 
									$powerUsers, 
									$web.LastItemModifiedDate, 
									$web.MasterPage, 
									$web.NoCrawl, 
									$web.Template,
									$web.CustomUpload) 
							}
						#endregion
				
					#region Loop through all lists and libraries in a sub site
						$lists = $web.Lists | ?{($_.Hidden -eq $false) -and ($_.Title -notin $skipLists)}
						ForEach ($list in $lists)
						{
							$listGUID = $list.Id.Guid;
							$listURL = $list.DefaultViewUrl
							if ($list.Title -eq $null){
								$listTitle = ""
							}
							else {
								$listTitle = $list.Title.Replace("|","\")
							}
							# Creating List Record
							Write-Host ("            Creating List record in Reporting Site: {0}" -f $list.Title)
							if(!($qLists) -or ($qLists.count -eq 0) -or !($qLists.ListGUID.Contains($list.Id.GUID))){
								#"Site|SiteURL|WebGUID|Web|WebURL|ListGUID|List|ListURL|ItemCount|Type|Created|LastItemDeleted|LastItemModifed|MajorVersionLimit|MinorVersionLimit|NoCrawl|EmailEnabled|EmailAlias"
								Out-File -FilePath $listPath -Append -InputObject ("{0}|{1}|{2}|{3}|{4}|{5}|{6}|{7}|{8}|{9}|{10}|{11}|{12}|{13}|{14}|{15}|{16}|{17}" -f 
									$siteTitle, 
									$siteUrl, 
									$webGUID,
									$webTitle, 
									$webUrl, 
									$listGUID, 
									$listTitle, 
									$listURL, 
									$list.Items.Count, 
									$list.BaseType, 
									$list.Created, 
									$list.LastItemDeletedDate, 
									$list.LastItemModifiedDate, 
									$list.MajorVersionLimit, 
									$list.MajorWithMinorVersionsLimit, 
									$list.NoCrawl, 
									$list.EmailEnabled, 
									$list.EmailAlias)
							}
							# Screen output for list
							Write-Host ("         Scanning List: {0}" -f $list.Title) -ForegroundColor DarkGray

						#region 7.	Check for Large Lists
							If ($list.ItemCount -gt $itemCount)
							{
								$qLargeListItems = New-Object System.Collections.Queue;
								#$BatchSize = $itemCount;

								Try {
										#Setup the context
		                                 $Ctx = New-Object Microsoft.SharePoint.Client.ClientContext($SiteURL);
	                                    #$Ctx.Credentials = $spCred;
										$Ctx.Credentials = New-Object System.Net.NetworkCredential($spCred.UserName , $ctxPassword)

										#Get the List
										$largeList = $Ctx.Web.Lists.GetByTitle($listTitle)
	                                    $Ctx.Load($largeList)
	                                    $Ctx.ExecuteQuery()

	                                    #Define Query to get List Items in batch
	                                     $Query = New-Object Microsoft.SharePoint.Client.CamlQuery
	                                     $Query.ViewXml = @"
											<View Scope='RecursiveAll'>
			                                    <Query>
				                                    <OrderBy><FieldRef Name='ID' Ascending='TRUE'/></OrderBy>
												</Query>
			                                    <RowLimit Paged="TRUE">$itemCount</RowLimit>
											</View>
"@

										#Get List Items in Batch
	                                    Do 
	                                    {
		                                    $ListItems = $largeList.GetItems($Query)
		                                    $Ctx.Load($ListItems)
		                                    $Ctx.ExecuteQuery()
											if($ListItems.count -gt 0) {
												$ListItems | %{ $qLargeListItems.Enqueue($_) };
											}
		                                    $Query.ListItemCollectionPosition = $ListItems.ListItemCollectionPosition
											if($qLargeListItems.count -gt (($itemCount)-1))
											{
												while($qLargeListItems.count -gt 0) 
												{
													$largeListItem = $qLargeListItems.Dequeue();

													if(!($qItems) -or ($qItems.count -eq 0) -or !($qItems.UniqueId.Contains($largeListItem['UniqueId'].guid))){
														if(($largeListItem.FieldValues['ContentTypeId']) -and ($null -ne $largeListItem.FieldValues['ContentTypeId']) -and ($largeListItem.FieldValues['ContentTypeId'].StringValue -ne '')){
															$liCTTypeId = ($largeListItem.FieldValues['ContentTypeId'].StringValue);
															if(!($qContentTypes) -or ($qContentTypes.count -eq 0) -or !($qContentTypes.ContentTypeID.Contains($liCTTypeId))){
																#$liCTType = (Get-PnPContentType -Web $web -list $list -Identity $liCTTypeId);
																$liCTType = $largeListItem.ContentType;
																if(($liCTType) -and ($null -ne $liCTType) -and ($liCTType.Name -ne '')){$liCTTypeName = $liCTType.Name.Replace("|","\")}else{$liCTTypeName = ""};
																if(($liCTType) -and ($null -ne $liCTType.Parent) -and ($liCTType.Parent -ne '')){$liCTTypeParent = $liCTType.Parent.Name}else{$liCTTypeParent = ""};
																if(($liCTType) -and ($null -ne $liCTType.Scope) -and ($liCTType.Scope -ne '')){$liCTTypeScope = $liCTType.Scope.Replace("|","\")}else{$liCTTypeScope = ""};
																if(($liCTType) -and ($null -ne $liCTType.Group) -and ($liCTType.Group -ne '')){$liCTTypeGroup = $liCTType.Group.Replace("|","\")}else{$liCTTypeGroup = ""};

																#"SiteGUID|Site|SiteURL|WebGUID|Web|WebURL|ListGUID|List|ListURL|ContentTypeID|ContentType|Parent|Scope"
																Out-File -FilePath $ctTypePath -Append -InputObject ("{0}|{1}|{2}|{3}|{4}|{5}|{6}|{7}|{8}|{9}|{10}|{11}|{12}" -f 
																	$siteGUID,
																	$siteTitle, 
																	$siteUrl, 
																	$webGUID,
																	$webTitle,
																	$webUrl, 
																	$listGUID,
																	$listTitle, 
																	$listURL, 
																	$liCTTypeId,
																	$liCTTypeName, 
																	$liCTTypeParent, 
																	$liCTTypeScope)
															}
															else{
																$prevLiCTType = $qContentTypes | ?{ ($_.ContentTypeID -eq $liCTTypeId) -and ($_.WebGUID -eq $webGUID) -and ($_.ListURL -eq $listURL) }
																$liCTTypeName =  $prevLiCTType.ContentType;
																$liCTTypeParent = $prevLiCTType.Scope;
																$liCTTypeScope = $prevLiCTType.Group;
															}

														}
														else{
															$liCTTypeId = ""
														};
						
														if(($largeListItem["FileLeafRef"]) -and ($null -ne $largeListItem["FileLeafRef"]) -and ($largeListItem["FileLeafRef"] -ne '')){$itemFileLeafRef = $largeListItem["FileLeafRef"].Replace("|","\")}else{$itemFileLeafRef = ""};
														if(($largeListItem["Author"].Email) -and ($null -ne  $largeListItem["Author"].Email) -and ( $largeListItem["Author"].Email -ne '')){$itemAuthor = $largeListItem["Author"].Email.Replace("|","\")}else{if(($largeListItem["Author"].LookupValue) -and ($null -ne  $largeListItem["Author"].LookupValue) -and ( $largeListItem["Author"].LookupValue -ne '')){$itemAuthor = $largeListItem["Author"].LookupValue.Replace("|","\")}else{$itemAuthor = ""};};
														if(($largeListItem["Editor"].Email) -and ($null -ne $largeListItem["Editor"].Email) -and ($largeListItem["Editor"].Email -ne '')){$itemEditor = $largeListItem["Editor"].Email.Replace("|","\")}else{if(($largeListItem["Editor"].LookupValue) -and ($null -ne  $largeListItem["Editor"].LookupValue) -and ( $largeListItem["Editor"].LookupValue -ne '')){$itemEditor = $largeListItem["Editor"].LookupValue.Replace("|","\")}else{$itemEditor = ""};};
														if(($largeListItem["CheckoutUser"].Email) -and ($null -ne $largeListItem["CheckoutUser"].Email) -and ($largeListItem["CheckoutUser"].Email -ne '')){$itemCheckedOutUser = $largeListItem["CheckoutUser"].Email.Replace("|","\")}else{$itemCheckedOutUser = ""};
														if(($largeListItem["FileRef"]) -and ($null -ne $largeListItem["FileRef"]) -and ($largeListItem["FileRef"] -ne '')){$itemFileRef = $largeListItem["FileRef"].Replace("|","\")}else{$itemFileRef = ""};

														#region 10.	Checked Out Documents
															If (($list.BaseType -ne "") -and 
																($list.CheckedOutFiles.Count -ne 0) -and 
																($null -ne $largeListItem.File) -and
																($largeListItem.File.CheckOutStatus -ne "None")
															){
																Write-Host ("               Creating Checked Out record: {0}" -f $largeListItem["FileLeafRef"])
																If ($null -eq $largeListItem["Modified_x0020_By"]){
																	$userModified = ""
																}
																else {
																	$userModified = $largeListItem["Modified_x0020_By"].Replace("|","/")
																}
																#Write to CSV if not already logged
																if(!($qCheckedOut) -or ($qCheckedOut.count -eq 0) -or !($qCheckedOut.UniqueId.Contains($largeListItem['UniqueId'].guid))){
																	#"Site|SiteURL|Web|WebURL|ListGUID|List|ListURL|Document|ContentType|SPID|UniqueID|CreatedBy|MajorVersion|MinorVersion|ModifiedBy|Created|Modified|ItemURL|CheckedOutBy|CheckedOutSince"
																	Out-File -FilePath $coPath -Append -InputObject ("{0}|{1}|{2}|{3}|{4}|{5}|{6}|{7}|{8}|{9}|{10}|{11}|{12}|{13}|{14}|{15}|{16}|{17}|{18}|{19}" -f 
																		$siteTitle, 
																		$site.Url, 
																		$webTitle,
																		$web.Url,
																		$listGUID, 
																		$listTitle, 
																		$listURL, 
																		$itemFileLeafRef, 
																		$liCTTypeName,
																		$largeListItem.ID,
																		$largeListItem['UniqueId'].guid,
																		$itemAuthor,
																		$largeListItem["_UIVersionString"].split('.')[0],
																		$largeListItem["_UIVersionString"].split('.')[1],
																		$itemEditor,
																		$largeListItem["Created"],
																		$largeListItem["Modified"],
																		$itemFileRef,
																		$itemCheckedOutUser,
																		$largeListItem["CheckedOutSince"])
																}
															}
														#endregion

														#region Capture List Items Information
															# Creating item Record
															#"Site|SiteURL|Web|WebURL|ListGUID|List|ListURL|ItemName|ContentType|SPID|UniqueID|CreatedBy|MajorVersion|MinorVersion|ModifiedBy|Created|Modified|ItemURL"
															Out-File -FilePath $itemPath -Append -InputObject ("{0}|{1}|{2}|{3}|{4}|{5}|{6}|{7}|{8}|{9}|{10}|{11}|{12}|{13}|{14}|{15}|{16}|{17}" -f 
																$siteTitle, 
																$site.Url, 
																$webTitle,
																$web.Url,
																$listGUID, 
																$listTitle, 
																$list.RootFolder.ServerRelativePath.DecodedUrl, 
																$itemFileLeafRef, 
																$liCTTypeName,
																$largeListItem.ID,
																$largeListItem['UniqueId'].guid,
																$itemAuthor,
																$largeListItem["_UIVersionString"].split('.')[0],
																$largeListItem["_UIVersionString"].split('.')[1],
																$itemEditor,
																$largeListItem["Created"],
																$largeListItem["Modified"],
																$itemFileRef)
														#endregion
													}	
												}
											}
	                                   	}
	                                	While($Query.ListItemCollectionPosition -ne $null)

									}
									Catch {
	                                     write-host -f Red "Error Getting List Items:" $_.Exception.Message
									}
		 
								} # If ($list.Items.Count -gt $MaxItems)
								else
                                {  
							        #region Loop through Items in List
							            Write-Host "            Scanning Items"
							            ForEach ($item in $list.Items)
							            {
					                        if(!($qItems) -or ($qItems.count -eq 0) -or !($qItems.UniqueId.Contains($item['UniqueId'].guid))){

                                                #region Parent Content Type Handlers
								                    switch -Wildcard ($item.ContentTypeId.ToString()) {
								                    #"0x0120*"{#Folder
								                    #        Write-Host ("ContentType" + $item.ContentType.Name);
								                    #        Write-Host ("ContentTypeID" + $item.ContentType | select Id);
								                    #        Write-Host ("ContentTypeParent" + $item.ContentType.Id.Parent);
								                    #        break}
                                                    "0x0101*"{#Document
										                    $liCTTypeName = $item.ContentType.Name;
										                    $liCreated = $item.File.TimeCreated;
										                    $liCreatedBy = $item.File.Properties.vti_author;
										                    $liModified = $item.File.TimeLastModified;
										                    $liModifiedBy = $item.File.Properties.vti_modifiedby;
										                    $liMajorVersion = $item.File.MajorVersion;
										                    $liMinorVersion = $item.File.MinorVersion;
										                    $liURL = $item.File.URL;
										                    break}
								                    "0x01*"{#Item
										                    $liCTTypeName = $item.ContentType.Name;
										                    $liCreated = ([XML]$item.XML).row.ows_Created;
										                    $liCreatedBy = ([XML]$item.XML).row.ows_Author;
										                    $liModified = ([XML]$item.XML).row.ows_Modified;
										                    $liModifiedBy = ([XML]$item.XML).row.ows_Editor;
										                    $liMajorVersion = ([XML]$item.XML).row.ows_owshiddenversion;
										                    $liMinorVersion = "0";
										                    $liURL = ([XML]$item.XML).row.ows_ServerURL;
										                    break}
								                    default {
										                    Write-Host ("ContentType" + $item.ContentType.Name);
										                    Write-Host ("ContentTypeID" + $item.ContentType | select Id);
										                    Write-Host ("ContentTypeParent" + $item.ContentType.Id.Parent);
										                    break}
								                    }
								                    #"0x0102*"{#Event
								                    #        Write-Host ("ContentType" + $item.ContentType.Name)
								                    #        Write-Host ("ContentTypeID" + $item.ContentType | select Id)
								                    #        break}
								                    #"0x0103*"{#Issue
								                    #        Write-Host ("ContentType" + $item.ContentType.Name)
								                    #        Write-Host ("ContentTypeID" + $item.ContentType | select Id)
								                    #        break}
								                    #"0x0104*"{#Announcement
								                    #        Write-Host ("ContentType" + $item.ContentType.Name)
								                    #        Write-Host ("ContentTypeID" + $item.ContentType | select Id)
								                    #        break}
								                    #"0x0105*"{#Link
								                    #        Write-Host ("ContentType" + $item.ContentType.Name)
								                    #        Write-Host ("ContentTypeID" + $item.ContentType | select Id)
								                    #        break}
								                    #"0x0106*"{#Contact
								                    #        Write-Host ("ContentType" + $item.ContentType.Name)
								                    #        Write-Host ("ContentTypeID" + $item.ContentType | select Id)
								                    #        break}
								                    #"0x0107*"{#Message
								                    #        Write-Host ("ContentType" + $item.ContentType.Name)
								                    #        Write-Host ("ContentTypeID" + $item.ContentType | select Id)
								                    #        break}
								                    #"0x0108*"{#Task
								                    #        Write-Host ("ContentType" + $item.ContentType.Name)
								                    #        Write-Host ("ContentTypeID" + $item.ContentType | select Id)
								                    #        break}
								                    #"0x0109*"{#Workflow History
								                    #        Write-Host ("ContentType" + $item.ContentType.Name)
								                    #        Write-Host ("ContentTypeID" + $item.ContentType | select Id)
								                    #        break}
								                    #"0x0110*"{#Post
								                    #        Write-Host ("ContentType" + $item.ContentType.Name)
								                    #        Write-Host ("ContentTypeID" + $item.ContentType | select Id)
								                    #        break}
								                    #"0x0111*"{#Comment
								                    #        Write-Host ("ContentType" + $item.ContentType.Name)
								                    #        Write-Host ("ContentTypeID" + $item.ContentType | select Id)
								                    #        break}
								                    #"0x0116*"{#East Asia Contact
								                    #        Write-Host ("ContentType" + $item.ContentType.Name)
								                    #        Write-Host ("ContentTypeID" + $item.ContentType | select Id)
								                    #        break}
                                
                                                #endregion

								                #region Get AD groups for List
								                    If ($item.HasUniqueRoleAssignments)
								                    {
									                    Write-Host ("               Retrieving AD groups: {0}" -f $item.Title)
									                    Get-CatSP10ADGroups -File $adPath -Assignments $Web.RoleAssignments -SiteName $Site.RootWeb.Title -SiteUrl $Site.Url -WebName $Web.Title -WebUrl $Web.Url -ListName $List.Title -LURL $ListURL -ItemTitle $Item.Name
								                    } # If ($web.HasUniqueRoleAssignments)
								                #endregion

								                #region Get Checked Out Documents
								                    #If (($list.BaseType -eq "DocumentLibrary") -and 
								                    If (($list.BaseType -ne "") -and 
                                                        ($list.CheckedOutFiles.Count -ne 0) -and 
                                                        ($null -ne $item.File) -and
                                                        ($item.File.CheckOutStatus -ne "None")
                                                    ){
									                    Write-Host ("               Creating Checked Out record in Reporting Site: {0}" -f $item.Name)
									                    #"Site|SiteURL|Web|WebURL|ListGUID|List|ListURL|Document|ContentType|SPID|UniqueID|CreatedBy|MajorVersion|MinorVersion|ModifiedBy|Created|Modified|ItemURL|CheckedOutBy|CheckedOutSince"
									                    Out-File -FilePath $coPath -Append -InputObject ("{0}|{1}|{2}|{3}|{4}|{5}|{6}|{7}|{8}|{9}|{10}|{11}|{12}|{13}|{14}|{15}|{16}|{17}|{18}|{19}" -f 
										                    $site.RootWeb.Title, 
										                    $site.Url, 
										                    $web.Title, 
										                    $web.Url, 
										                    $listGUID, 
										                    $list.Title, 
										                    $listURL, 
										                    $item.Name, 
										                    $item.ContentType.Name, 
										                    $item.ID, 
										                    $item.UniqueID.GUID, 
										                    $liCreatedBy, 
										                    $liMajorVersion, 
										                    $liMinorVersion, 
										                    $liModifiedBy, 
										                    $liCreated, 
										                    $liModified, 
										                    $liURL, 
										                    $item.file.CheckedOutByUser.UserLogin, 
										                    $item.File.CheckedOutDate)
								                    } 
								                #endregion

								                #region Capture List Items Information
									                # Write-Host ("           	 Capturing List Item information in Reporting Site: {0}" -f $list.Title)
									
									                # Creating item Record
									                if( ($liCreated) -and
										                ($null -ne $liCreated) -and
										                ("" -ne $liCreated)
									                ){
										                #"Site|SiteURL|Web|WebURL|ListGUID|List|ListURL|ItemName|ContentType|SPID|UniqueID|CreatedBy|MajorVersion|MinorVersion|ModifiedBy|Created|Modified|ItemURL"
										                Out-File -FilePath $itemPath -Append -InputObject ("{0}|{1}|{2}|{3}|{4}|{5}|{6}|{7}|{8}|{9}|{10}|{11}|{12}|{13}|{14}|{15}|{16}|{17}" -f 
											                $siteTitle, 
											                $siteUrl, 
											                $webTitle, 
											                $webUrl, 
											                $listGUID, 
											                $listTitle, 
											                $listURL, 
											                $item.Name, 
											                $liCTTypeName, 
											                $item.ID, 
											                $item.UniqueID.GUID, 
											                $liCreatedBy, 
											                $liMajorVersion, 
											                $liMinorVersion, 
											                $liModifiedBy, 
											                $liCreated, 
											                $liModified, 
											                $liURL)
									                }else{
										                Write-Host ("ContentType" + $item.ContentType.Name)
									                }
								                #endregion
							                }
                                        }
							        #endregion
                              
							    } # If ($list.Items.Count -gt $MaxItems)
							#endregion

							# 8.	Check for InfoPath Forms Libraries
							If (($list.ContentTypes[0].ResourceFolder.Properties["_ipfs_infopathenabled"]) -or (($list.BaseType -eq "DocumentLibrary") -and ($list.BaseTemplate -eq "XMLForm")))
							{
								# Creating InfoPath Record
								Write-Host ("            Creating InfoPath record in Reporting Site: {0}" -f $list.Title)
								Out-File -FilePath $ipPath -Append -InputObject ("{0}|{1}|{2}|{3}|{4}|{5}" -f $site.RootWeb.Title, $site.Url, $web.Title, $web.Url, $list.Title, $listURL)
							} # If (InfoPath)

							# 9.	Check for Complex Columns
							ForEach ($column in $colTypes)
							{
								$fields = $list.Fields | ?{($_.TypeAsString -like $column) -and ($_.Hidden -eq $false) -and ($_.Title -notin $skipColumns)}
								foreach($field in $fields)
								{
									Write-Host ("            Creating {0} column record in Reporting Site: {1}" -f $column, $field.Title)
									Out-File -FilePath $colPath -Append -InputObject ("{0}|{1}|{2}|{3}|{4}|{5}|{6}|{7}" -f $site.RootWeb.Title, $site.Url, $web.Title, $web.Url, $list.Title, $listURL, $field.Title, $column)
								} # foreach($field in $fields)
							} # ForEach ($column in $colTypes)

							# 10.	Check for Workflows
							ForEach ($workflow in $list.WorkflowAssociations)
							{
								Write-Host ("            Creating Workflow record in Reporting Site: {0}" -f $workflow.Name)
								Out-File -FilePath $wfPath -Append -InputObject ("{0}|{1}|{2}|{3}|{4}|{5}|{6}" -f $site.RootWeb.Title, $site.Url, $web.Title, $web.Url, $list.Title, $listURL, $workflow.Name)
							} # ForEach ($workflow in $list.WorkflowAssociations)


						} # ForEach ($list in $web.Lists)
						#endregion
					
					# Screen Output for Web
					Write-Host ("         {0} complete" -f $web.Title) -ForegroundColor Green
					$web.Dispose()
				} # ForEach ($web in $webs)
				#endregion
			#endregion
			# Dispose of site object
			$site.Dispose()
			#endregion

		} # ForEach ($site in $sites)
		#endregion
	#endregion
#endregion

#region Completed Screen Output
Complete-CatScript $Start[0]
#endregion

$outputFolderNewName = 'OnPrem';
$outputFolder = ("..\_Output Files\{0}\{1}" -f $client, $scriptStart[1])
$outputFolderNewPath = ("..\_Output Files\{0}\{1}" -f $client,$outputFolderNewName)
Rename-Item -Path $outputFolder -NewName $outputFolderNewName -Verbose
Compress-Archive -Path $outputFolderNewPath -DestinationPath ("..\_Output Files\{0}\{1}.ZIP" -f $client, $scriptStart[1]) -CompressionLevel Optimal -Verbose -Force
Invoke-Item ("..\_Output Files\{0}" -f $client)