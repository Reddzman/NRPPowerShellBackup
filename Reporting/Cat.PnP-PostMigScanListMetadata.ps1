<#  Cat.PnP-PostMigScanListMetadata.ps1

This script leverage PnP PowerShell to compare item level metadata between lists from a CSV file.

Dependencies
	-	PnP PowerShell https://docs.microsoft.com/en-us/powershell/sharepoint/sharepoint-pnp/sharepoint-pnp-cmdlets?view=sharepoint-ps
	-	Input CSV File in Input Files folder - Script Parameter - "Cat.PerformAction.Test.csv"

Call script with parameters
	.\Cat.PnP-PostMigScanListMetadata.ps1 -inputFile Clients\Cat.PnP-ListMetadataCompare-Test.csv -clientName "Client"

CSV Example
	Site,siteURL,SPOUrl
	SiteName,https://sharepoint.client.com/sitedirectory/law,https://client.sharepoint.com/sites/law

Table of Contents
	1.	Input & Output Files
	2.	Variables
	3.	Snapin(s), Modules & Environment Connection(s)
	4.	Main Script
	5.	Scanning SPO List
	6.	Scanning SP On Premises List
#>

#region Script Parameters
Param(	
		[string] $inputFile = "Testing\Cat.PnP-ListMetadataCompare-Test.csv",
		[string] $columnFile = "Testing\Cat.PnP-ListMetadataColumns-Test.csv",
		[string] $clientName = "Testing",
		[int] $itemPageCount = 5000
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
Start-Transcript -Path ("..\_Log FIles\{0}-PostMigScanListMetadata-{1}.log" -f $clientName, $scriptStart[1])
#endregion

#region 1.	Input & Output Files
# File Paths
$inPath = ("..\_Input Files\{0}" -f $inputFile)
$columnPath = ("..\_Input Files\{0}" -f $columnFile)
$outPath = ("..\_Output Files\{0}{1}.csv" -f $inputFile.Replace(".csv","-"), $scriptStart[1])

# Import Input File(s)
# Load CSV file
If (Test-Path ($inPath)){
	$inFile = Import-Csv $inPath -Encoding Unicode
}
else {
	Write-Host
	throw ("The input file does not exist: {0}" -f $inPath)	
}
If (Test-Path ($columnPath)){
	$colFile = Import-Csv $columnPath -Encoding Unicode
}
else {
	Write-Host
	throw ("The column file does not exist: {0}" -f $columnPath)	
}

# Create Output File(s)
$outHeader = "environment|siteURL|listName|itemName|itemTitle|itemID|itemPath|itemCreated|itemCreatedBy|itemModified|itemModifiedBy"
ForEach ($column in $colFile){
	$outHeader += ("|{0}" -f $column.columnName)
}
Out-File -FilePath $outPath -InputObject $outHeader -Encoding unicode
#endregion

#region 2.	Variables
# User Inputs
Write-Host
Write-Host "User Input Variables" -ForegroundColor Cyan
Write-Host "   Enter credentials for SharePoint Online:"
$spoCred = Get-Credential -Message "Enter your credentials for SharePoint Online:" 
Write-Host "   Enter credentials for SharePoint On Premises:"
$spCred = Get-Credential -Message "Enter your credentials for SharePoint On Premises:"

# Script Variables
$itemsQuery = "<View Scope='RecursiveAll'>
	<Query>
		<Where>
			<And>
			<Geq>
				<FieldRef Name='ID' />
				<Value Type='Number'>[StartID]</Value>
			</Geq>
			<Leq>
				<FieldRef Name='ID' />
				<Value Type='Number'>[EndID]</Value>
			</Leq>
			</And>
		</Where>
	</Query>
</View>"

# Output Parameters & Variables
Write-Host
Write-Host "Script Parameters" -ForegroundColor Cyan
Write-Host ("   Input File: {0}" -f $inputFile)
Write-Host ("   Client Name: {0}" -f $clientName)
Write-Host ("   Item Page Count: {0}" -f $itemPageCount)
Write-Host
Write-Host "User Inputs" -ForegroundColor Cyan
Write-Host ("   SP Online Credentials: {0}" -f $spoCred.UserName)
Write-Host ("   SP on Premises Credentials: {0}" -f $spCred.UserName)
<#Write-Host
Write-Host "CSV Header" -ForegroundColor Cyan
Write-Host $outHeader #>
<#Write-Host
Write-Host "Items Query Template" -ForegroundColor Cyan
Write-Host $itemsQuery #>
#endregion

#region 3.	Snapin(s), Modules & Environment Connection(s)
# Load Custom Modules
Import-Module "..\Reporting\Cat.ReportingFunctions.psm1"
#endregion
    
#region 4.	Main Script
# Screen Output
Write-Host
Write-Host ("Scanning Lists from CSV File: {0}" -f $inputFile) -ForegroundColor Cyan

# Loop through Lists
ForEach ($list in $inFile){
	Write-Host
	Write-Host ("   Scanning List: {0}/{1}" -f $list.dSiteURL, $list.listPath) -ForegroundColor Magenta

	#region 5.	Scanning SPO List
	# Connecting to Site
	Write-Host ("      Connecting to SPO Site: {0}" -f $list.dSiteURL)
	Connect-PnPOnline -Url $list.dSiteURL -Credentials $spoCred
	$pnpContext = Get-PnPContext

	# Get List
	Write-Host ("      Retrieving SPO List: ") -NoNewline
	$pnpList = Get-PnPList -Identity $list.listPath
	Write-Host $pnpList.Title

	# Get the last ID used in the list
	Write-Host ("      Retrieving SPO Last ID: ") -NoNewline
	$lastID = Get-LastID -context $pnpContext -listID $pnpList.ID
	Write-Host $lastID

	# Iterate through batchs of items
	$startID = 1
	$endID = $itemPageCount
	Write-Host ("      Retrieving All SPO Items")
	while ($startID -le $lastID) {
		$query = $itemsQuery
		If($endID -gt $lastID){
			$e = $lastID
		} 
		Else{
			$e = $endID
		}
		$query = $query.Replace("[StartID]",$startID).Replace("[EndID]",$e)
		
		# Screen Output for Troubleshooting
		#$query # This will output the CAML QUery prior to doing the query

		# Get All Items
		Write-Host ("         Writing SPO Items: {0} through {1}" -f $startID, $e) -ForegroundColor DarkGray
		$items = Get-PnPListItem -List $pnpList -Query $query

		# Loop Through All Items and Add a metadata record
		ForEach ($item in $items){
			# Screen Output for Troubleshooting
			#Write-Host ("            Writing Item: {0} - {1}" -f $item.ID, $item.FieldValues["FileLeafRef"]) -ForegroundColor Blue # This will output each item to the screen.

			# Create base record entry
			$record = ("{0}|{1}|{2}|{3}|{4}|{5}|{6}|{7}|{8}|{9}|{10}" -f
				"SPO",
				$list.dSiteURL,
				$pnpList.Title,
				$item.FieldValues["FileLeafRef"],
				$item.Title,
				$item.ID,
				$item.FieldValues["FileRef"],
				$item.FieldValues["Created"],
				$item.FieldValues["Created_x0020_By"],
				$item.FieldValues["Modified"],
				$item.FieldValues["Modified_x0020_By"]
			)
			
			# Add Dynamic column information to each record
			ForEach ($column in $colFile){
				Try{
					Switch -Wildcard ($item.FieldValues[$column.columnName]){
						"*TaxonomyFieldValue" {$record += ("|{0}" -f $item.FieldValues[$column.columnName].Label)}
						"*FieldUserValue" {$record += ("|{0}" -f $item.FieldValues[$column.columnName].Email)}
						Default {$record += ("|{0}" -f $item.FieldValues[$column.columnName])}
					}
				}
				Catch{
					$record += ("|")
				}
			}

			# Write record to CSV
			Out-File -FilePath $outPath -Append -InputObject $record
		}
		
		# Incrment Start & End ID number
		$startID += $itemPageCount
		$endID += $itemPageCount
	}

	# Screen Output
	Write-Host ("      SPO List Complete") -ForegroundColor Green
	#endregion #>

	#region 6.	Scanning SP On Premises List
	# Connecting to Site
	Write-Host ("      Connecting to SP On Premises Site: {0}" -f $list.sSiteURL)
	Connect-PnPOnline -Url $list.sSiteURL -Credentials $spCred
	$pnpContext = Get-PnPContext

	# Get List
	Write-Host ("      Retrieving SP On Premises List: ") -NoNewline
	$pnpList = Get-PnPList -Identity $list.listPath
	Write-Host $pnpList.Title

	# Get the last ID used in the list
	Write-Host ("      Retrieving SP On Premises Last ID: ") -NoNewline
	$lastID = Get-LastID -context $pnpContext -listID $pnpList.ID
	Write-Host $lastID

	# Iterate through batchs of items
	$startID = 1
	$endID = $itemPageCount
	Write-Host ("      Retrieving All SP On Premises Items")
	while ($startID -le $lastID) {
		$query = $itemsQuery
		If($endID -gt $lastID){
			$e = $lastID
		} 
		Else{
			$e = $endID
		}
		$query = $query.Replace("[StartID]",$startID).Replace("[EndID]",$e)
		
		# Screen Output for Troubleshooting
		#$query # This will output the CAML QUery prior to doing the query

		# Get All Items
		Write-Host ("         Writing SP On Premises Items: {0} through {1}" -f $startID, $e) -ForegroundColor DarkGray
		$items = Get-PnPListItem -List $pnpList -Query $query

		# Loop Through All Items and Add a metadata record
		ForEach ($item in $items){
			# Screen Output for Troubleshooting
			#Write-Host ("            Writing Item: {0} - {1}" -f $item.ID, $item.FieldValues["FileLeafRef"]) -ForegroundColor Blue # This will output each item to the screen.

			# Create base record entry
			$record = ("{0}|{1}|{2}|{3}|{4}|{5}|{6}|{7}|{8}|{9}|{10}" -f
				"SP On Premises",
				$list.sSiteURL,
				$pnpList.Title,
				$item.FieldValues["FileLeafRef"],
				$item.Title,
				$item.ID,
				$item.FieldValues["FileRef"],
				$item.FieldValues["Created"],
				$item.FieldValues["Created_x0020_By"],
				$item.FieldValues["Modified"],
				$item.FieldValues["Modified_x0020_By"]
			)
			
			# Add Dynamic column information to each record
			ForEach ($column in $colFile){
				Try{
					Switch -Wildcard ($item.FieldValues[$column.columnName]){
						"*TaxonomyFieldValue" {$record += ("|{0}" -f $item.FieldValues[$column.columnName].Label)}
						"*FieldUserValue" {$record += ("|{0}" -f $item.FieldValues[$column.columnName].Email)}
						Default {$record += ("|{0}" -f $item.FieldValues[$column.columnName])}
					}
				}
				Catch{
					$record += ("|")
				}
			}

			# Write record to CSV
			Out-File -FilePath $outPath -Append -InputObject $record
		}
		
		# Incrment Start & End ID number
		$startID += $itemPageCount
		$endID += $itemPageCount
	}

	# Screen Output
	Write-Host ("      SP On Premises List Complete") -ForegroundColor Green
	#endregion #>
}

# Screen Output
Write-Host
Write-Host ("Scanning Complete") -ForegroundColor Green
#endregion

#region Completed Script
# Disconnect from SharePoint
Disconnect-PnPOnline

# Output Script Info
Complete-CatScript $scriptStart[0]

# Stop logging
Stop-Transcript
#endregion