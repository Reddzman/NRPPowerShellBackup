<#  Cat.PnP-ArvhiveItemsToSQL.ps1

This script script is designed to archive items from the AZDPS PAR system to a SQL Database to avoid threshold limits.  It looks up settings in 
a settings list on the PAR site to determine the Database/Table information and the active time period for list items.  Archived items will
be in a Complete status and will not have been modified in the past X months based on the setting. The decision was made at AZDPS that the requst and
attachments will be printed to PDF and stored in another system to meet retention requirements.  Therefore, Archived requests will be copied to SQL for 
reporting purposes and all attachments will be deleted. Currently built for SP13 PnP PowerShell.

Dependencies
	-	SharePointPnPPowerShell2013 - https://docs.microsoft.com/en-us/powershell/sharepoint/sharepoint-pnp/sharepoint-pnp-cmdlets?view=sharepoint-ps
	-	SQLServer - https://docs.microsoft.com/en-us/powershell/module/sqlserver/?view=sqlserver-ps
	-	SP Columns CSV file: ..\_Input Files\Cat.PnP-ArchiveItemsToSQL-Columns.csv
	-	SP Account JSON file: ..\_Input Files\Cat.PnP-ArchiveItemsToSQL-Acct.json

Call script with parameters
	.\Cat.PnP-ArchivePARItems.ps1  -siteURL https://devportal.azdps.gov/sites/PAR -settingsList Lists/PARSettings -requestList Lists/PARs -documentsList PARDocs -secureDocumentsList SecureDocs -memosList Lists/memos -notificationsList Lists/PARNotifications -spColumnInput Cat.PnP-ArchiveItemsToSQL-Columns.csv -client Testing -recycleBin $true -skipConfirmation $true

Table of Contents
	1.	Input & Output Files
	2.	Snapin(s), Modules & Environment Connection(s)
	3.	Variables
	4.	Main Script
	5.	Getting Settings from SharePoint
	6.	Process List Items
	7.	Writing To SQL
	8.	Confirm Successfull Copy
	9.	Deleting Documents
	10.	Deleting Secure Documents
	11.	Deleting Memos
	12.	Deleting Notifications
	13. Deleting PAR Record
#>

#region Script Parameters
Param(	[string] $siteURL = "https://devportal.azdps.gov/sites/PAR",
        [string] $settingsList = "Lists/PARSettings",
		[string] $requestList = "Lists/PARs",
		[string] $documentsList = "PARDocs",
		[string] $secureDocumentsList = "SecureDocs",
		[string] $memosList = "Lists/memos",
		[string] $notificationsList = "Lists/PARNotifications",
		[string] $spColumnInput = "Samples\Archival\Cat.PnP-ArchiveItemsToSQL-Columns.csv",
		[string] $spAcctInput = "Samples\Archival\Cat.PnP-ArchiveItemsToSQL-Acct.json",
		[string] $client = "Testing",
		[boolean] $recycleBin = $true,
		[boolean] $skipConfirmation = $true
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
Start-Transcript -Path ("..\_Log FIles\{0}.PnP-ArchivePARItems-{1}.log" -f $client, $scriptStart[1])
#endregion

#region 1.	Input & Output Files
# File Paths
$spColumnsPath = ("..\_Input Files\{0}" -f $spColumnInput)
$spAcctPath = ("..\_Input Files\{0}" -f $spAcctInput)

# Import Input File(s)
If (Test-Path ($spColumnsPath)){
	$spColumnsList = Import-Csv $spColumnsPath -Encoding Unicode
}
else {
	Write-Host
	throw ("The SharePoint Columns input file does not exist: {0}" -f $spColumnsPath)	
}
If (Test-Path ($spAcctPath)){
	$spAcctJSON = Get-Content $spAcctPath -Raw | ConvertFrom-Json
	[securestring] $secureWord = ConvertTo-SecureString $spAcctJSON.word -AsPlainText -Force
	[pscredential] $spAcct = New-Object System.Management.Automation.PSCredential ($spAcctJSON.name,$secureWord)
}
else {
	Write-Host
	throw ("The SharePoint Acct input file does not exist: {0}" -f $spAcctPath)	
}
#endregion

#region 2.	Snapin(s), Modules & Environment Connection(s)
# Load MS / PnP PowerShell Modules
Get-Module -ListAvailable | Where{$_.Name -like "SharePoint*"} | Remove-Module
Import-Module SharePointPnPPowershell2013 -DisableNameChecking # Add appropriate PnP Module for your use case
Import-Module SqlServer
#endregion

#region 3.	Variables
# Script Variables
$settingsQuery = "<View><Query><Where><And><Eq><FieldRef Name='PARSettingType' /><Value Type='Text'>Archival</Value></Eq><Eq><FieldRef Name='Title' />
	<Value Type='Text'>[Setting]</Value></Eq></And></Where></Query></View>"
$requestQuery = "<View><Query><Where><And><Eq><FieldRef Name='PARPARStatus' /><Value Type='Text'>Complete</Value></Eq><Leq><FieldRef Name='Modified' />
	<Value IncludeTimeValue='TRUE' Type='DateTime'>[ModifiedDate]</Value></Leq></And></Where></Query></View>"
$confirmQuery = "SELECT * FROM [TableName] WHERE [ID] = '[SearchID]'"
$docsQuery = "<View><Query><Where><Eq><FieldRef Name='PARRequestSPID' /><Value Type='Text'>[SearchID]</Value></Eq></Where></Query></View>"
$notificationsQuery = "<View><Query><Where><Eq><FieldRef Name='PARNotificationRequestSharePoint' /><Value Type='Text'>[SearchID]</Value></Eq></Where></Query></View>"
# Array of values to be replaced in SP column names to match the column names in the SQL table for the variables used in the write to SQL Command.
$replaceColumnNames = @("_x002f_","_x0028_","_x0023_","_x0029_","_x002d_","_x2019_","_x002e_","_x0020_","_x005F_","_x005F_","_x002","_x")
# Hash table of special character and ASCII tSQL CHAR commands to be replaced in the value for the variables used in the write to SQL Command.
$replaceValues = @{"="="'+CHAR(61)+'";"'"="'+CHAR(39)+'"}

# Output Parameters & Variables
Write-Host
Write-Host "Script Parameters" -ForegroundColor Cyan
Write-Host ("   Site URL: {0}" -f $siteURL)
Write-Host ("   Settings List: {0}" -f $settingsList)
Write-Host ("   Request List: {0}" -f $requestList)
Write-Host ("   Documents List: {0}" -f $documentsList)
Write-Host ("   Secure Documents List: {0}" -f $secureDocumentsList)
Write-Host ("   Memos List: {0}" -f $memosList)
Write-Host ("   Notifications List: {0}" -f $notificationsList)
Write-Host ("   SP Columns Input: {0}" -f $spColumnInput)
Write-Host ("   SP Acct Input: {0}" -f $spAcctInput)
Write-Host ("   Client: {0}" -f $client)
Write-Host ("   Use Recycle Bin: {0}" -f $recycleBin)
Write-Host
Write-Host "Script Variables" -ForegroundColor Cyan
Write-Host ("   Service Account: {0}" -f $spAcct.UserName)
<#Write-Host
Write-Host "Base Settings Query" -ForegroundColor Cyan
$settingsQuery
Write-Host
Write-Host "Base Request Query" -ForegroundColor Cyan
$requestQuery #>
<#Write-Host
Write-Host "Base Confirm Query" -ForegroundColor Cyan
$confirmQuery #>
<#Write-Host
Write-Host "Base Documents Query" -ForegroundColor Cyan
$docsQuery #>
<#Write-Host
Write-Host "Base Notifications Query" -ForegroundColor Cyan
$docsQuery #>
<#Write-Host
Write-Host "Values to be replaced in Column Names" -ForegroundColor Cyan
$replaceColumnNames #>
<#Write-Host
Write-Host "Values to be replaced in values" -ForegroundColor Cyan
$replaceValues #>
#endregion
    
#region 4.	Main Script
# Connecting to Site
Write-Host
Write-Host "Connecting to Services" -ForegroundColor Cyan
Write-Host ("   Connecting to SharePoint: {0}" -f $siteURL)
Connect-PnPOnline -Url $siteURL -UseAdfs -Credentials $spAcct

#region 5.	Getting Settings from SharePoint
Write-Host
Write-Host ("Retrieving Settings List Information: {0}" -f $settingsList) -ForegroundColor Cyan
$spSettingsList = Get-PnPList -Identity $settingsList
$sqlServer = (Get-PnPListItem -list $spSettingsList -query $settingsQuery.Replace("[Setting]","SQL Server"))["SettingValue"]
$sqlDB = (Get-PnPListItem -list $spSettingsList -query $settingsQuery.Replace("[Setting]","SQL Database"))["SettingValue"]
$sqlTable = (Get-PnPListItem -list $spSettingsList -query $settingsQuery.Replace("[Setting]","SQL Table"))["SettingValue"]
$activeDuration = (Get-PnPListItem -list $spSettingsList -query $settingsQuery.Replace("[Setting]","Active Duration"))["SettingValue"]

# Output Settings values
Write-Host ("   SQL Server: {0}" -f $sqlServer)
Write-Host ("   SQL DB: {0}" -f $sqlDB)
Write-Host ("   SQL Table: {0}" -f $sqlTable)
Write-Host ("   Active Duration (Days): {0}" -f $activeDuration)
#endregion

#region 6.	Process List Items
# Screen Output
Write-Host
Write-Host ("Archiving Request List Items: {0}" -f $requestList) -ForegroundColor Magenta

# Get List items that meet criteria
Write-Host ("   Retrieving Items to be archived")
$filterDate = (Get-Date).AddDays(-($activeDuration))
[string] $filterDateString = Get-Date -Date $filterDate -Format "yyyy-MM-ddTHH:mm:ssZ"
$spRequestList = Get-PnPList -Identity $requestList
$spRequestItems = Get-PnPListItem -list $spRequestList -query $RequestQuery.Replace("[ModifiedDate]",$filterDateString)

# Process items ready for archival
if ($spRequestItems.Count -gt 0){
	Write-Host ("   Processing {0} items ready for archival" -f $spRequestItems.Count)

	# Updating Table Query Information
	[string] $confirmQuery = $confirmQuery.Replace("[TableName]",$sqlTable)

	# Loop through items
	foreach ($spItem in $spRequestItems){
		#region 7.	Writing To SQL
		Write-Host
		Write-Host ("      Writing record to SQL: {0}" -f $spItem["PARPARRequestID"]) -ForegroundColor DarkGray -NoNewline

		# Updating New Query Record
		[string] $sqlNewQuery = ("INSERT INTO {0} (" -f $sqlTable)
		[string] $sqlNewQueryValues = ""
		$spItemSQLVariables = @()
		foreach ($column in $spColumnsList){
			$columnName = $column.columnInternalName
			foreach ($replace in $replaceColumnNames) {
				$columnName = $columnName.Replace($replace,"")
			}
			if ($null -ne $spItem[$column.columnInternalName]) {
				switch ($column.columnType) {
					User {[string] $value = $spItem[$column.columnInternalName].LookupValue}
					DateTime {[string] $value = Get-Date -Date $spItem[$column.columnInternalName]}
					URL {[string] $value = $spItem[$column.columnInternalName].Url}
					Default {[string] $value = $spItem[$column.columnInternalName]}
				}
				$value = $value.Trim()
				foreach ($replace in $replaceValues.GetEnumerator()) {
					$value = $value.Replace($replace.Name,$replace.Value)
				}
				$value = ("'{0}'" -f $value)
				$spItemSQLVariables += ("{0}={1}" -f $columnName, $value)
				$sqlNewQuery = $sqlNewQuery + ("{0}," -f $columnName)
				$sqlNewQueryValues = $sqlNewQueryValues + ("`$({0})," -f $columnName)
			}
		}
		$sqlNewQuery = $sqlNewQuery.Substring(0,$sqlNewQuery.Length-1)
		$sqlNewQueryValues = $sqlNewQueryValues.Substring(0,$sqlNewQueryValues.Length-1)
		$sqlNewQuery = ("{0}) VALUES ({1})" -f $sqlNewQuery, $sqlNewQueryValues)		

		<#region Use for testing Query updates
		Write-Host
		Write-Host
		Write-Host ("* SQL Insert Into Variables: {0}" -f $spItem["PARPARRequestID"]) -ForegroundColor Yellow
		$spItemSQLVariables
		Write-Host
		Write-Host ("* SQL Insert Into Query: {0}" -f $spItem["PARPARRequestID"]) -ForegroundColor Yellow
		$sqlNewQuery
		Write-Host 
		#endregion #>

		# Write new recored to SQL
		try {
			Invoke-Sqlcmd -ServerInstance $sqlServer -Database $sqlDB -Query $sqlNewQuery -Variable $spItemSQLVariables -Verbose
			Write-Host " - Success" -ForegroundColor Green
		}
		catch {
			Write-Host " - Failure" -ForegroundColor Red
			Write-Host "         Exception Message: $($_.Exception.Message)"  -ForegroundColor Red
		}
		#endregion #>

		#region 8.	Confirm Successfull Copy
		Write-Host "      Confirming record written to SQL successfully: " -ForegroundColor DarkGray -NoNewline
		$confirmed = @()
		try {
			$confirmed += Invoke-Sqlcmd -ServerInstance $sqlServer -Database $sqlDB -Query $confirmQuery.Replace("[SearchID]",$spItem["ID"])
		}
		catch {
			Write-Host " - Failure" -ForegroundColor Red
			Write-Host ". Exception Message: $($_.Exception.Message)"
		}

		If ($confirmed.Count -eq 1){
			Write-Host ("{0}" -f $confirmed[0].PARPARRequestID) -ForegroundColor DarkGray -NoNewline
			Write-Host " - Success" -ForegroundColor Green

			#region 9.	Deleting Documents
			# Get documents to delete
			$deleteList = Get-PnPList -Identity $documentsList
			$deleteItems = Get-PnPListItem -List $deleteList -Query $docsQuery.Replace("[SearchID]",$spItem["ID"])

			# Deleting Documents
			If ($deleteItems.Count -gt 0){
				Write-Host ("      Deleting record documents: {0}" -f $spItem["PARPARRequestID"]) -ForegroundColor DarkGray
				foreach ($deleteItem in $deleteItems) {
					Write-Host ("         Deleting document: {0}" -f $deleteItem["FileLeafRef"]) -ForegroundColor Blue -NoNewline
					try {
						Remove-PnPListItem -List $deleteList -Identity $deleteItem -Recycle:$recycleBin -Force:$skipConfirmation
						Write-Host " - Success" -ForegroundColor Green
					}
					catch {
						Write-Host " - Failure" -ForegroundColor Red
						Write-Host ". Exception Message: $($_.Exception.Message)" 
					}
				}
			}
			else {
				Write-Host ("      No documents to be deleted for {0}" -f $spItem["PARPARRequestID"]) -ForegroundColor DarkGray
			}
			#endregion #>

			#region 10.	Deleting Secure Documents
			# Get documents to delete
			$deleteList = Get-PnPList -Identity $secureDocumentsList
			$deleteItems = Get-PnPListItem -List $deleteList -Query $docsQuery.Replace("[SearchID]",$spItem["ID"])

			# Deleting Documents
			If ($deleteItems.Count -gt 0){
				Write-Host ("      Deleting record secure documents: {0}" -f $spItem["PARPARRequestID"]) -ForegroundColor DarkGray
				foreach ($deleteItem in $deleteItems) {
					Write-Host ("         Deleting document: {0}" -f $deleteItem["FileLeafRef"]) -ForegroundColor Blue -NoNewline
					try {
						Remove-PnPListItem -List $deleteList -Identity $deleteItem -Recycle:$recycleBin -Force:$skipConfirmation
						Write-Host " - Success" -ForegroundColor Green
					}
					catch {
						Write-Host " - Failure" -ForegroundColor Red
						Write-Host ". Exception Message: $($_.Exception.Message)" 
					}
				}
			}
			else {
				Write-Host ("      No secure documents to be deleted for {0}" -f $spItem["PARPARRequestID"]) -ForegroundColor DarkGray
			}
			#endregion #>

			#region 11.	Deleting Memos
			# Get memos to delete
			$deleteList = Get-PnPList -Identity $memosList
			$deleteItems = Get-PnPListItem -List $deleteList -Query $docsQuery.Replace("[SearchID]",$spItem["ID"])

			# Deleting Documents
			If ($deleteItems.Count -gt 0){
				Write-Host ("      Deleting record memos: {0}" -f $spItem["PARPARRequestID"]) -ForegroundColor DarkGray
				foreach ($deleteItem in $deleteItems) {
					Write-Host ("         Deleting memo: {0}" -f $deleteItem["PARMemoSubject"]) -ForegroundColor Blue -NoNewline
					try {
						Remove-PnPListItem -List $deleteList -Identity $deleteItem -Recycle:$recycleBin -Force:$skipConfirmation
						Write-Host " - Success" -ForegroundColor Green
					}
					catch {
						Write-Host " - Failure" -ForegroundColor Red
						Write-Host ". Exception Message: $($_.Exception.Message)" 
					}
				}
			}
			else {
				Write-Host ("      No memos to be deleted for {0}" -f $spItem["PARPARRequestID"]) -ForegroundColor DarkGray
			}
			#endregion #>

			#region 12.	Deleting Notifications
			# Get notifications to delete
			$deleteList = Get-PnPList -Identity $notificationsList
			$deleteItems = Get-PnPListItem -List $deleteList -Query $notificationsQuery.Replace("[SearchID]",$spItem["ID"])

			# Deleting Notifications
			If ($deleteItems.Count -gt 0){
				Write-Host ("      Deleting record notifications: {0}" -f $spItem["PARPARRequestID"]) -ForegroundColor DarkGray
				Write-Host "         Warning: Notifications tied to multiple PARs will not be deleted!" -ForegroundColor Yellow
				foreach ($deleteItem in $deleteItems) {
					Write-Host ("         Deleting notification: {0}" -f $deleteItem["Title"]) -ForegroundColor Blue -NoNewline
					try {
						Remove-PnPListItem -List $deleteList -Identity $deleteItem -Recycle:$recycleBin -Force:$skipConfirmation
						Write-Host " - Success" -ForegroundColor Green
					}
					catch {
						Write-Host " - Failure" -ForegroundColor Red
						Write-Host ". Exception Message: $($_.Exception.Message)" 
					}
				}
			}
			else {
				Write-Host ("      No notifications to be deleted for {0}" -f $spItem["PARPARRequestID"]) -ForegroundColor DarkGray
				Write-Host "         Warning: Notifications tied to multiple PARs will not be deleted!" -ForegroundColor Yellow
			}
			#endregion #>

			#region 13. Deleting PAR Record
			Write-Host ("      Deleting PAR record: {0}" -f $spItem["PARPARRequestID"]) -ForegroundColor DarkGray -NoNewline
			try {
				Remove-PnPListItem -List $spRequestList -Identity $spItem -Recycle:$recycleBin -Force:$skipConfirmation
				Write-Host " - Success" -ForegroundColor Green
			}
			catch {
				Write-Host " - Failure" -ForegroundColor Red
				Write-Host ". Exception Message: $($_.Exception.Message)" 
			}
			#endregion #>
		} 
		elseif ($confirmed -eq 0) {
			Write-Host ("         Failure to create SQL record: {0}" -f $spItem["PARPARRequestID"]) -ForegroundColor Red
		} 
		else {
			Write-Host " - Multiple Records" -ForegroundColor Yellow
			Write-Host "         Nothing will be Deleted due to mutliple records with same SharePoint ID. Please remediate manually!" -ForegroundColor Blue
		}
	}
}
else {
	Write-Host "   No items are ready to be archvied at this time"
}
#endregion

# Screen Output
Write-Host
Write-Host ("Archiving requests complete: {0}" -f $requestList) -ForegroundColor Green
#endregion

#region Completed Script
# Disconnect from Services
Write-Host
Write-Host "Disconnecting from Services:" -ForegroundColor Cyan
Write-Host ("   Disconnecting from SharePoint: {0}" -f $siteURL)
Disconnect-PnPOnline

# Output Script Info
Complete-CatScript $scriptStart[0]

# Stop logging
Stop-Transcript
#endregion