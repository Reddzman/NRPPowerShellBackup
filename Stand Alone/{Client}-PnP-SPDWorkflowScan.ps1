#region Documentation
<#  {Client}-PnP-SPDWorkflowScan.ps1

This script leverages the workflow output from the Microsoft Modernization Scanner to collect additional information about each workflow. 
Depending on your authentication method, the script will create and enterprise application in Azure for PnP Office 365 Management Shell. To create
the enterprise application, Global Administrator permissions are required.  If the enterprise application already exists than you can sign in
with your sharepoint admin permissions.

Dependencies
	-	PnP PowerShell https://docs.microsoft.com/en-us/powershell/sharepoint/sharepoint-pnp/sharepoint-pnp-cmdlets?view=sharepoint-ps
	-	Modernization Scanner CSV File - Example: ModernizationWorkflowScanResults.csv

Call script with parameters
	.\{Client}-PnP-SPDWorkflowScan.ps1  -inputFile .\ModernizationWorkflowScanResults.csv -tenantName "M365TenantName"

	If tenant URL is "contoso.sharepoint.com", tenant name is "contoso"

Table of Contents
	1.	Input & Output Files
	2.	Variables
	3.	Snapin(s), Modules & Environment Connection(s)
	4.	Main Script
	5.	Build Base Workflow Record
	6.	Collect List Related Information
	7.	Workflow instance Information
	8.	Latest Workflow History Information
	9.	Latest Workflow Task Information
	10.	Writing to CSV
#>
#endregion

#region Script Parameters
Param(	[string] $inputFile = "ModernizationWorkflowScanResults.csv",
		[string] $tenantName = "{client}",
		[string] $spoAdminURL = "https://$tenantName-admin.sharepoint.com"
)
#endregion

#region Functions
function Start-CatScript{
	# Collecting Start Time
	$SDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
	$FDate = Get-Date -Format "yyMMddHHmm"
	
	# Return Start Time
	$DateArray = @()
	$DateArray += $SDate
	$DateArray += $FDate
	return $DateArray
} # Start-CatScript

function Complete-CatScript{
	param(
        [parameter(Mandatory=$true, Position=0, ValueFromPipeline=$true)][datetime]$STime
    )
	
	# Finish Time & Time Taken
	$FTime = Get-Date -Format "MM/dd/yyyy HH:mm:ss"
	$TTime = New-TimeSpan -Start $STime -End $FTime

	# Screen Output
	Write-Host
	Write-Host "-----------------------------------------------------------------------------------------------------" -ForegroundColor Cyan
	Write-Host ("Started: {0}   Finished: {1}   Time Taken: {2}" -f $STime, $FTime, $TTime) -ForegroundColor Cyan
	Write-Host "-----------------------------------------------------------------------------------------------------" -ForegroundColor Cyan
	Write-Host

	# Stopping Transcript
	#Stop-Transcript
} # Complete-CatScript
#endregion

#region Start Date / Time
# Get Start Date / Time
$scriptStart = Start-CatScript

# Start Transcript
Start-Transcript -Path ("..\_Log FIles\{0}.{1}-{2}.log" -f $tenantName, $inputFile.Replace(".csv",""), $scriptStart[1])
#endregion

#region 1.	Input & Output Files
# File Paths
$inPath = ("{0}" -f $inputFile)
$outPath = ("{0}{1}.csv" -f $inputFile.Replace(".csv","-").Replace("Modernization","Catapult"), $scriptStart[1])

# Create Output File(s)
Out-File -FilePath $outPath -InputObject "Site Url,Site Collection Url,Definition Name,Definition ID,Subscription Name,Subscription Id,List Url,List Id,List Item Count,List Last Item Modified Date,List Last Item User Modified Date,List Last item Deleted Date,Most Recent Instance Created Date,Most Recent Instance Modified Date,Completed Instances,ErrorOccurred Instances,ErrorOccurredRetrying Instances,FailedOnStart Instances,FailedOnStartRetrying Instances,InProgress Instances,Max Instances,NotStarted Instances,StoppedByUser Instances,ViewQueryOverflow Instances,Last History Item Created,Last Task Item Created"

# Import Input File(s)
# Load CSV file
If (Test-Path ($inPath)){
	$inFile = Import-Csv $inPath #-Encoding Unicode # Modernization scanner doesn't create the CSV files with Unicode
}
else {
	Write-Host
	throw ("The input file does not exist: {0}" -f $inPath)	
}
#endregion

#region 2.	Variables
# Script Variables
$listViewThreshold = 5000
$sp10Version = "2010"
$sp13Version = "2013"
$wfHistoryName = "Workflow History"
$wfStatusColumnType = "WorkflowStatus"
$sp10WFTaskName = "Tasks"
$sp13WFTaskName = "Workflow Tasks"
$queryHistoryList = "@<View Scope='RecursiveAll'>
<Query>
	<Where>
		<And>
			<Eq>
				<FieldRef Name='FSObjType' />
				<Value Type='Integer'>0</Value>
			</Eq>
			<In>
				<FieldRef Name='WorkflowAssociation' />
				<Values>
					<Value Type='Text'>{[WFAssociationID]}</Value>
				</Values>
			</In>
		</And>
	</Where>
	<OrderBy>
		<FieldRef Name='Created' Ascending='False'/>
	</OrderBy>
</Query>
<RowLimit>1</RowLimit>
</View>"
$querysp13TaskList = "@<View Scope='RecursiveAll'>
<Query>
	<Where>
		<And>
			<Eq>
				<FieldRef Name='FSObjType' />
				<Value Type='Integer'>0</Value>
			</Eq>
			<In>
				<FieldRef Name='WF4InstanceId' />
				<Values>
					<Value Type='Text'>[WFInstanceID]</Value>
				</Values>
			</In>
		</And>
	</Where>
	<OrderBy>
		<FieldRef Name='Created' Ascending='False'/>
	</OrderBy>
</Query>
<RowLimit>1</RowLimit>
</View>"
$querysp10TaskList = "@<View Scope='RecursiveAll'>
<Query>
	<Where>
		<And>
			<Eq>
				<FieldRef Name='FSObjType' />
				<Value Type='Integer'>0</Value>
			</Eq>
			<Eq>
				<FieldRef Name='WorkflowName' />
				<Value Type='Text'>[WorkflowName]</Value>
			</Eq>
		</And>
	</Where>
	<OrderBy>
		<FieldRef Name='Created' Ascending='False'/>
	</OrderBy>
</Query>
<RowLimit>1</RowLimit>
</View>"
$querysp10Instances = "@<View Scope='RecursiveAll'>
<Query>
	<Where>
		<And>
			<Eq>
				<FieldRef Name='FSObjType' />
				<Value Type='Integer'>0</Value>
			</Eq>
			<IsNotNull>
				<FieldRef Name='[WFStatusColumnName]' />
			</IsNotNull>
		</And>
	</Where>
	<OrderBy>
		<FieldRef Name='Created' Ascending='False'/>
	</OrderBy>
</Query>
</View>"
<# These variables are used for connecting via an app registration.  Recommended if we can get the customer to set it up.
$certPass = ConvertTo-SecureString "pwd" -AsPlainText -Force
$appId = "00000000-0000-0000-0000-000000000000";
$certPath = "path\to\certificate.pfx";
$aadUrl = "$client.com";
#>

# Output Parameters & Variables
Write-Host
Write-Host "Script Parameters" -ForegroundColor Cyan
Write-Host ("   Input File: {0}" -f $inputFile)
Write-Host ("   SPO Root URL: {0}" -f $spoAdminURL)
Write-Host ("   Client Abreviation: {0}" -f $tenantName)
Write-Host
Write-Host "Script Variables" -ForegroundColor Cyan
Write-Host ("   List Veiw Threshold: {0}" -f $listViewThreshold)
Write-Host ("   SP 10 WF Version Option: {0}" -f $sp10Version)
Write-Host ("   SP 13 WF Version Option: {0}" -f $sp13Version)
Write-Host ("   WF History List Name: {0}" -f $wfHistoryName)
Write-Host ("   WF Status Column Type: {0}" -f $wfStatusColumnType)
Write-Host ("   SP 10 WF Tasks List Name: {0}" -f $sp10WFTaskName)
Write-Host ("   SP 13 WF Tasks List Name: {0}" -f $sp13WFTaskName)
<#Write-Host
Write-Host "SP 10 Instances CAML Query" -ForegroundColor Cyan
Write-Host $querysp10Instances #>
<#Write-Host
Write-Host "Workflow History CAML Query" -ForegroundColor Cyan
Write-Host $queryHistoryList #>
<#Write-Host
Write-Host "SP 13 Workflow Task CAML Query" -ForegroundColor Cyan
Write-Host $querysp13TaskList #>
<#Write-Host
Write-Host "SP 10 Workflow Task CAML Query" -ForegroundColor Cyan
Write-Host $querysp10TaskList #>
#endregion

#region 3.	Snapin(s), Modules & Environment Connection(s)
# Load MS / PnP PowerShell Modules
Import-Module SharePointPnPPowershellOnline -DisableNameChecking <# Choose appropriate PnP Module for your use case
Import-Module SharePointPnPPowershell2013 -DisableNameChecking 
Import-Module SharePointPnPPowershell2016 -DisableNameChecking 
Import-Module SharePointPnPPowershell2019 -DisableNameChecking  #>
#endregion
    
#region 4.	Main Script
# Connecting To SP Online
Write-Host
Write-Host "Connecting to Services:" -ForegroundColor Cyan
Write-Host "   Connecting to SharePoint Online: " -nonewline
# Below are many to authenticate to SPO (Web login, PnPManagementShell and App registration) you may need to update parameters/variables to use the appropriate mechanism.
#$pnpConnection = Connect-PnPOnline -url $spoAdminURL.Replace("-admin", "").Replace("-Admin", "") -UseWebLogin
$pnpConnection = Connect-PnPOnline -Url $spoAdminURL.Replace("-admin", "").Replace("-Admin", "") -PnPManagementShell -LaunchBrowser
<# Use this for connecting via an App registration
	$pnpConnectionWeb = Connect-PnPOnline `
		-CertificatePath $certPath `
		-Tenant $aadUrl `
		-ClientId $appId `
		-CertificatePassword $certPass `
		-IgnoreSslErrors `
		-NoTelemetry `
		-Url $spoAdminURL.Replace("-admin", "").Replace("-Admin", "") `
		-ReturnConnection 	#>
Write-Host ("{0}" -f $spoAdminURL.Replace("-admin", "").Replace("-Admin", ""))

# Screen output
Write-Host
Write-Host ("Gathering information on Workflows: {0}" -f $inputFile) -ForegroundColor Cyan

# Loop through modernization scanner workflow results
foreach ($workflow in $inFile) {
	# Starting Output for specific workflows
	Write-Host
	Write-Host ("   Site Collection: {0} | {2} Workflow: {1}" -f $workflow."Site Collection URL", $workflow."Definition Name", $workflow.Version) -ForegroundColor Magenta
	
	# Connecting to site collection for specific workflow
	Write-Host "      Connecting to site for specific workflow: " -nonewline
	# Below are 2 ways to authenticate to SPO (Web login and App registration) you may need to update parameters/variables to use the appropriate mechanism.
	#Connect-PnPOnline -Url $workflow."Site URL" -UseWebLogin # Comment this out if you are using an app registration
	Connect-PnPOnline -Url $workflow."Site URL" -PnPManagementShell
	<# Use this for connecting via an App registration
	$pnpConnectionWeb = Connect-PnPOnline `
		-CertificatePath $certPath `
		-Tenant $aadUrl `
		-ClientId $appId `
		-CertificatePassword $certPass `
		-IgnoreSslErrors `
		-NoTelemetry `
		-Url $workflow."Site URL" `
		-ReturnConnection 	#>
	Write-Host ("{0}" -f $workflow."Site URL")

	#region 5.	Build Base Workflow Record
	Write-Host "      Creating Initial Workflow Record: " -nonewline
	$wfRecord = New-Object PSObject
	$wfRecord | Add-Member NoteProperty "Site URL"($workflow."Site URL")
	$wfRecord | Add-Member NoteProperty "Site Collection URL"($workflow."Site Collection URL")
	$wfRecord | Add-Member NoteProperty "Definition Name"($workflow."Definition Name")
	$wfRecord | Add-Member NoteProperty "Definition ID"($workflow."Definition Id")
	$wfRecord | Add-Member NoteProperty "List Url"($workflow."List Url")
	$wfRecord | Add-Member NoteProperty "List Id"($workflow."List Id")
	$wfRecord | Add-Member NoteProperty "Subscription Name"($workflow."Subscription Name")
	$wfRecord | Add-Member NoteProperty "Subscription Id"($workflow."Subscription Id")
	Write-Host ("{0}" -f $wfRecord."Definition Name")
	#endregion

	#region 6.	Collect List Related Information
	# If list information provided
	if (-not [String]::IsNullOrWhitespace($workflow."List Url")) {
		# Retrieving List
		Write-Host "      Retrieving List Information: " -nonewline
		$wfList = (Get-PnPList -Identity $workflow."List ID" -Includes WorkflowAssociations,Fields | Select -First 1)

		# Adding list information to Workflow Record
		$wfRecord | Add-Member NoteProperty "List Item Count"($wfList.ItemCount)
		$wfRecord | Add-Member NoteProperty "List Last Item Modified Date"($wfList.LastItemModifiedDate)
		$wfRecord | Add-Member NoteProperty "List Last Item User Modified Date"($wfList.LastItemUserModifiedDate)
		$wfRecord | Add-Member NoteProperty "List Last Item Deleted Date"($wfList.LastItemDeletedDate)

		# List Information Successfull
		Write-Host ("{0}" -f $workflow."List Url")
	}
	else {
		# Adding placeholder list information to Workflow Record
		$wfRecord | Add-Member NoteProperty "List Item Count"(-1)
		$wfRecord | Add-Member NoteProperty "List Last Item Modified Date"("1/1/1111")
		$wfRecord | Add-Member NoteProperty "List Last Item User Modified Date"("1/1/1111")
		$wfRecord | Add-Member NoteProperty "List Last Item Deleted Date"("1/1/1111")
	}
	#endregion

	#region 7.	Workflow instance Information
	# Different logic based on version
	switch ($workflow.Version) {
		$sp13Version {
			# Retrieving Instance Information
			Write-Host "      Retrieving Instance Information: " -nonewline
			$wfSubscription = Get-PnPWorkflowSubscription -Name $workflow."Subscription Name"
			$wfInstances = @()
			$wfInstances += Get-PnPWorkflowInstance -WorkflowSubscription $wfSubscription
			$wfInstanceCreated = $wfInstances | Sort-Object -Property InstanceCreated -Descending | Select-Object -First 1
			$wfInstanceLastUpdated = $wfInstances | Sort-Object -Property LastUpdated -Descending | Select-Object -First 1
			$wfInstanceGroupBy = $wfInstances | Group-Object -Property Status

			# Adding instance information to Workflow Record
			$wfRecord | Add-Member NoteProperty "Most Recent Instance Created Date"($wfInstanceCreated.InstanceCreated)
			$wfRecord | Add-Member NoteProperty "Most Recent Instance Modified Date"($wfInstanceLastUpdated.LastUpdated)
			$wfRecord | Add-Member NoteProperty "Completed Instances"(($wfInstanceGroupBy | Where {$_.Name -eq "Completed"}).Count)
			$wfRecord | Add-Member NoteProperty "ErrorOccurred Instances"(($wfInstanceGroupBy | Where {$_.Name -eq "ErrorOccurred"}).Count)
			$wfRecord | Add-Member NoteProperty "ErrorOccurredRetrying Instances"(($wfInstanceGroupBy | Where {$_.Name -eq "ErrorOccurredRetrying"}).Count)
			$wfRecord | Add-Member NoteProperty "FailedOnStart Instances"(($wfInstanceGroupBy | Where {$_.Name -eq "FailedOnStart"}).Count)
			$wfRecord | Add-Member NoteProperty "FailedOnStartRetrying Instances"(($wfInstanceGroupBy | Where {$_.Name -eq "FailedOnStartRetrying"}).Count)
			$wfRecord | Add-Member NoteProperty "InProgress Instances"(($wfInstanceGroupBy | Where {$_.Name -eq "InProgress"}).Count)
			$wfRecord | Add-Member NoteProperty "Max Instances"(($wfInstanceGroupBy | Where {$_.Name -eq "Max"}).Count)
			$wfRecord | Add-Member NoteProperty "NotStarted Instances"(($wfInstanceGroupBy | Where {$_.Name -eq "NotStarted"}).Count)
			$wfRecord | Add-Member NoteProperty "StoppedByUser Instances"(($wfInstanceGroupBy | Where {$_.Name -eq "StoppedByUser"}).Count)
			$wfRecord | Add-Member NoteProperty "ViewQueryOverflow Instances"(($wfInstanceGroupBy | Where {$_.Name -eq "ViewQueryOverflow"}).Count)			

			# Instance Information Successfull
			Write-Host ("{0}" -f $workflow."Definition Name")
		}
		$sp10Version {
			if (-not [String]::IsNullOrWhiteSpace($workflow."List Url")) {
				# Retrieve Status Column
				$wfStatusColumn = $wfList.Fields | Where{($_.TypeAsString -eq $wfStatusColumnType) -and ($_.Title -eq $workflow."Definition Name")} | Sort-Object -Property InternalName -Descending | Select-Object -First 1

				# Query Workflow Task List
				$queryUpdated = $querysp10Instances.Replace("[WFStatusColumnName]",$wfStatusColumn.InternalName)
				$wfInstances = @()
				$wfInstances += Get-PnPListItem -list $wfList -Query $queryUpdated
				$wfInstanceCreated = $wfInstances | Sort-Object -Property Created -Descending | Select-Object -First 1
				$wfInstanceLastUpdated = $wfInstances | Sort-Object -Property Modified -Descending | Select-Object -First 1
				#$wfInstanceGroupBy = $wfInstances | Group-Object -Property ("FieldValues.{0}" -f $wfStatusColumn.InternalName)

				# Adding instance information to Workflow Record
				# TODO: Get proper instance counts via the list status
				$wfRecord | Add-Member NoteProperty "Most Recent Instance Created Date"($wfInstanceCreated.FieldValues.Created)
				$wfRecord | Add-Member NoteProperty "Most Recent Instance Modified Date"($wfInstanceLastUpdated.FieldValues.Modified)
				$wfRecord | Add-Member NoteProperty "Completed Instances"(-10)
				$wfRecord | Add-Member NoteProperty "ErrorOccurred Instances"(-10)
				$wfRecord | Add-Member NoteProperty "ErrorOccurredRetrying Instances"(-10)
				$wfRecord | Add-Member NoteProperty "FailedOnStart Instances"(-10)
				$wfRecord | Add-Member NoteProperty "FailedOnStartRetrying Instances"(-10)
				$wfRecord | Add-Member NoteProperty "InProgress Instances"(-10)
				$wfRecord | Add-Member NoteProperty "Max Instances"(-10)
				$wfRecord | Add-Member NoteProperty "NotStarted Instances"(-10)
				$wfRecord | Add-Member NoteProperty "StoppedByUser Instances"(-10)
				$wfRecord | Add-Member NoteProperty "ViewQueryOverflow Instances"(-10)
			}
			else {
				# Adding placeholder list information to Workflow Record
				$wfRecord | Add-Member NoteProperty "Most Recent Instance Created Date"("10/10/1010")
				$wfRecord | Add-Member NoteProperty "Most Recent Instance Modified Date"("10/10/1010")
				$wfRecord | Add-Member NoteProperty "Completed Instances"(-10)
				$wfRecord | Add-Member NoteProperty "ErrorOccurred Instances"(-10)
				$wfRecord | Add-Member NoteProperty "ErrorOccurredRetrying Instances"(-10)
				$wfRecord | Add-Member NoteProperty "FailedOnStart Instances"(-10)
				$wfRecord | Add-Member NoteProperty "FailedOnStartRetrying Instances"(-10)
				$wfRecord | Add-Member NoteProperty "InProgress Instances"(-10)
				$wfRecord | Add-Member NoteProperty "Max Instances"(-10)
				$wfRecord | Add-Member NoteProperty "NotStarted Instances"(-10)
				$wfRecord | Add-Member NoteProperty "StoppedByUser Instances"(-10)
				$wfRecord | Add-Member NoteProperty "ViewQueryOverflow Instances"(-10)
			}
		}
		Default {
			# Adding placeholder instance information to Workflow Record
			$wfRecord | Add-Member NoteProperty "Most Recent Instance Created Date"("1/1/1111")
			$wfRecord | Add-Member NoteProperty "Most Recent Instance Modified Date"("1/1/1111")
			$wfRecord | Add-Member NoteProperty "Completed Instances"(-1)
			$wfRecord | Add-Member NoteProperty "ErrorOccurred Instances"(-1)
			$wfRecord | Add-Member NoteProperty "ErrorOccurredRetrying Instances"(-1)
			$wfRecord | Add-Member NoteProperty "FailedOnStart Instances"(-1)
			$wfRecord | Add-Member NoteProperty "FailedOnStartRetrying Instances"(-1)
			$wfRecord | Add-Member NoteProperty "InProgress Instances"(-1)
			$wfRecord | Add-Member NoteProperty "Max Instances"(-1)
			$wfRecord | Add-Member NoteProperty "NotStarted Instances"(-1)
			$wfRecord | Add-Member NoteProperty "StoppedByUser Instances"(-1)
			$wfRecord | Add-Member NoteProperty "ViewQueryOverflow Instances"(-1)
		}
	}
	#endregion

	#region 8.	Latest Workflow History Information

	# Logic based on Version
	switch ($workflow.Version) {
		$sp13Version {
			# Using $wfSubscription variable from 7.	Workflow instance Information
			$wfHistoryListId = $null;
			$wfHistoryListId = ($wfSubscription | Select -First 1 | Select @{ Expression={$_.PropertyDefinitions.HistoryListId}; Name="HistoryListId" }).HistoryListId;

			$wfHistoryList = $null;
			if (-not [String]::IsNullOrWhiteSpace($wfHistoryListId)) {
				$wfHistoryList = (Get-PnPList -Identity $wfHistoryListId | Select -First 1);
			
				# Check View Threshold before quering.
				If ($wfHistoryList.ItemCount -gt 0 -and $wfHistoryList.ItemCount -le $listViewThreshold){
					Write-Host "      Retrieving Workflow History Information (Total: $($wfHistoryList.ItemCount)): " -nonewline

					# Query Workflow History List
					$wfHistoryCreated = $null
					$queryUpdated = $queryHistoryList.Replace("{[WFAssociationID]}", $workflow."Subscription Name")

					$wfHistoryResults = Get-PnPListItem -list $wfHistoryList -Query $queryUpdated
					Write-Host "($($wfHistoryResults.Count) matches) $wfHistoryListId";

					$wfHistoryCreated = $wfHistoryResults | Sort-Object -Property Created -Descending | Select-Object -First 1

					# Adding workflow History information to Workflow Record
					$wfRecord | Add-Member NoteProperty "Last History Item Created"($wfHistoryCreated.FieldValues.Created)
				}
				else {
					# Screen output.
					Write-Host -F YELLOW ("      Zero or Too Many Items in WF History List: {0}" -f $wfHistoryList.ItemCount)
					$wfRecord | Add-Member NoteProperty "Last History Item Created"("1/1/1010")
				}
			}
			else {
				# Adding placeholder workflow History information to Workflow Record
				$wfRecord | Add-Member NoteProperty "Last History Item Created"("1/1/1010")
			}
		}
		$sp10Version {
			if (-not [String]::IsNullOrWhiteSpace($workflow."List URL")) {
				# Using $wfList variable from 6.	Collect List Related Information
				$wfAssociations = $wfList.WorkflowAssociations | Where {$_.InternalName -Like ("{0}*" -f $workflow."Definition Name")}
				
				# Retrieving Workflow History
				$wfHistoryListTitle = $null;
				$wfHistoryListTitle = ($wfAssociations | Select -First 1 | Select @{ Expression={$_.HistoryListTitle}; Name="HistoryListTitle" }).HistoryListTitle;
				
				$wfHistoryList = $null;
				if (-not [String]::IsNullOrWhiteSpace($wfHistoryListTitle)) {
					$wfHistoryList = (Get-PnPList -Identity $wfHistoryListTitle | Select -First 1);

					# Check View Threshold before quering.
					If ($wfHistoryList.ItemCount -gt 0 -and $wfHistoryList.ItemCount -le $listViewThreshold){
						Write-Host "      Retrieving Workflow History Information (Total: $($wfHistoryList.ItemCount)): " -nonewline
						$WFAssociationIDs = [string]($wfAssociations | Select @{ Expression={$_.ID.GUID}; Name="WFAssociationID" }).WFAssociationID

						# Query Workflow History List
						$wfHistoryCreated = $null
						$queryUpdated = $queryHistoryList.Replace("[WFAssociationID]", $WFAssociationIDs.Replace(" ", "}</Value><Value Type='Text'>{"))

						$wfHistoryResults = Get-PnPListItem -list $wfHistoryList -Query $queryUpdated
						Write-Host "($($wfHistoryResults.Count) matches) $wfHistoryTitle";

						$wfHistoryCreated = $wfHistoryResults | Sort-Object -Property Created -Descending | Select-Object -First 1
			
						# Adding workflow History information to Workflow Record
						$wfRecord | Add-Member NoteProperty "Last History Item Created"($wfHistoryCreated.FieldValues.Created)
					}
					else {
						# Screen output.
						Write-Host -F YELLOW ("      Zero or Too Many Items in WF History List: {0}" -f $wfHistoryList.ItemCount)
						$wfRecord | Add-Member NoteProperty "Last History Item Created"("1/1/1010")
					}
				}
				else {
					# Adding placeholder workflow History information to Workflow Record
					$wfRecord | Add-Member NoteProperty "Last History Item Created"("1/1/1010")
				}
			}
			else {
				# Adding placeholder workflow History information to Workflow Record
				$wfRecord | Add-Member NoteProperty "Last History Item Created"("1/1/1010")
			}
		}
		Default {
			# Adding placeholder workflow History information to Workflow Record
			Write-Host ("Unrecognized Workflow Version: {0} - " -f $workflow.Version) -nonewline
			$wfRecord | Add-Member NoteProperty "Last History Item Created"("1/1/1111")
		}
	}
	#endregion #>

	#region 9.	Latest Workflow Task Information
	# Retrieving Workflow Tasks List
	$wfTaskList = $null;

	switch ($workflow.Version) {
		$sp13Version {
			# Retrieving Workflow Tasks list
			Write-Host "      Retrieving Workflow Task List: " -nonewline

			# Using $wfSubscription variable from 7.	Workflow instance Information
			$wfTaskListId = $null;
			$wfTaskListId = ($wfSubscription | Select -First 1 | Select @{ Expression={$_.PropertyDefinitions.TaskListId}; Name="TaskListId" }).TaskListId;
			Write-Host $wfTaskListId

			$wfTaskList = $null;
			if (-not [String]::IsNullOrWhiteSpace($wfTaskListId)) {
				$wfTaskList = (Get-PnPList -Identity $wfTaskListId | Select -First 1);
			}
		}
		$sp10Version {
			# Retrieving Workflow Tasks list
			Write-Host "      Retrieving Workflow Task List: " -nonewline

			# Using $wfAssociations variable from 8.	Latest Workflow History Information
			$wfTaskListTitle = $null;
			$wfTaskListTitle = ($wfAssociations | Select -First 1 | Select @{ Expression={$_.TaskListTitle}; Name="TaskListTitle" }).TaskListTitle
			Write-Host $wfTaskListTitle

			$wfTaskList = $null;
			if (-not [String]::IsNullOrWhiteSpace($wfTaskListTitle)) {
				$wfTaskList = (Get-PnPList -Identity $wfTaskListTitle | Select -First 1);
			}
		}
		Default {
			# Screen Ouptut
			Write-Host ("      Unrecognized Workflow Version: {0}" -f $workflow.Version)
		}
	}

	# Check View Threshold before quering.
	If ($null -ne $wfTaskListId) {
		if ($wfTaskList.ItemCount -gt 0 -and $wfTaskList.ItemCount -le $listViewThreshold){
			# Logic based on Version
			switch ($workflow.Version) {
				$sp13Version {
					# Retrieving Workflow Tasks
					Write-Host "      Retrieving Workflow Task Information (Total: $($wfTaskList.ItemCount)): " -nonewline
					
					# Using $wfInstances variable from 7.	Workflow instance Information
					$WFInstanceIDs = [string]($wfInstances | Select @{ Expression={$_.Id}; Name="WFInstanceID" }).WFInstanceID

					# Query Workflow Task List
					$wfTaskCreated = $null
					$queryUpdated = $querysp13TaskList.Replace("[WFInstanceID]", $WFInstanceIDs.Replace(" ", "</Value><Value Type='Text'>"))

					$wfTaskResults = Get-PnPListItem -list $wfTaskList -Query $queryUpdated
					Write-Host "$($wfTaskResults.Count) matches";

					$wfTaskCreated = $wfTaskResults | Sort-Object -Property Created -Descending | Select-Object -First 1	

					# Adding workflow Task information to Workflow Record
					$wfRecord | Add-Member NoteProperty "Last Task Item Created"($wfTaskCreated.FieldValues.Created)
				}
				$sp10Version {
					# Retrieving Workflow Tasks
					Write-Host "      Retrieving Workflow Task Information: " -nonewline

					# Query Workflow Task List
					$wfTaskCreated = $null
					$queryUpdated = $querysp10TaskList.Replace("[WorkflowName]", $workflow."Definition Name")

					$wfTaskResults = Get-PnPListItem -list $wfTaskList -Query $queryUpdated
					Write-Host "$($wfTaskResults.Count) matches";

					$wfTaskCreated = $wfTaskResults | Sort-Object -Property Created -Descending | Select-Object -First 1

					# Adding placeholder workflow History information to Workflow Record
					$wfRecord | Add-Member NoteProperty "Last Task Item Created"($wfTaskCreated.FieldValues.Created)
				}
				Default {
					# Adding placeholder workflow Task information to Workflow Record
					Write-Host ("Unrecognized Workflow Version: {0} - " -f $workflow.Version) -nonewline
					$wfRecord | Add-Member NoteProperty "Last Task Item Created"("1/1/1111")
				}
			}
		}
		else {
			# Screen output.
			Write-Host -F YELLOW ("      Zero or Too Many Items in Tasks List: {0}" -f $wfTaskList.ItemCount)
			$wfRecord | Add-Member NoteProperty "Last Task Item Created"("1/1/1111")
		}	
	}
	else {
		# Screen output.
		$wfRecord | Add-Member NoteProperty "Last Task Item Created"("1/1/1111")
	}
	#endregion

	#region 10.	Writing to CSV
	Write-Host "      Writing to CSV: " -nonewline
	Out-File -FilePath $outPath -Append -InputObject ('"{0}","{1}","{2}",{3},"{4}",{5},{6},{7},{8},{9},{10},{11},{12},{13},{14},{15},{16},{17},{18},{19},{20},{21},{22},{23},{24},{25}' -f
	$wfRecord."Site Url",
	$wfRecord."Site Collection Url",
	$wfRecord."Definition Name",
	$wfRecord."Definition ID",
	$wfRecord."Subscription Name",
	$wfRecord."Subscription Id",
	$wfRecord."List Url",
	$wfRecord."List Id",
	$wfRecord."List Item Count",
	$wfRecord."List Last Item Modified Date",
	$wfRecord."List Last Item User Modified Date",
	$wfRecord."List Last item Deleted Date",
	$wfRecord."Most Recent Instance Created Date",
	$wfRecord."Most Recent Instance Modified Date",
	$wfRecord."Completed Instances",
	$wfRecord."ErrorOccurred Instances",
	$wfRecord."ErrorOccurredRetrying Instances",
	$wfRecord."FailedOnStart Instances",
	$wfRecord."FailedOnStartRetrying Instances",
	$wfRecord."InProgress Instances",
	$wfRecord."Max Instances",
	$wfRecord."NotStarted Instances",
	$wfRecord."StoppedByUser Instances",
	$wfRecord."ViewQueryOverflow Instances",
	$wfRecord."Last History Item Created",
	$wfRecord."Last Task Item Created")
	
	# Instance Information Successfull
	Write-Host ("{0}" -f $outPath)
	#endregion
}
#endregion

#region Completed Script
# Disconnect from Services
Write-Host
Write-Host "Disconnecting from Services:" -ForegroundColor Cyan
Write-Host "   Disconnecting from SharePoint Online"
Disconnect-PnPOnline

# Output Script Info
Complete-CatScript $scriptStart[0]

# Stop logging
Stop-Transcript
#endregion