<# Cat.CM-NewListJobs.ps1

This script creates a migration job for the lists and Libraries listed in a CSV file.  It is designed to break up large lists into multiple jobs
by looping through month ranges based on jobFilter, filterDateField, filterStartDate, filterEndDate & filterNumMoBetween.

Dependencies
	-	Metalogix PowerShell Snapins
	-	Cat.GeneralFunctions.psm1 in General folder
	-	Cat.CMFunctions.psm1 - should be updated with data from "create powershell" feature in Content Matrix
	-	Input CSV File in _Input Files\Clients folder for job list
	-	Input XML File in _Input Files\Clients folder for filter xml - should be updated with data from "create powershell" feature in Content Matrix 

Call Script with Input File
	.\BCT.CM-NewListJobs.ps1 -jobFile CM-ListMigrations-Client.csv

CSV Example
	jobName,sSiteURL,dSiteURL,listPath,jobfilter,filterDateField,filterStartDate,filterEndDate,filterNumMoBetween
	sourcesite-list1,https://sharepoint.client.com/sites/sourcesite,https://client.sharepoint.com/sites/destinationsite,list1,CMj-ListMigrations-CircleK-Filter.xml,Modified,1/1/2019,10/15/2019,1
	sourcesite-list2,https://sharepoint.client.com/sites/sourcesite,https://client.sharepoint.com/sites/destinationsite,list2,CMj-ListMigrations-CircleK-Filter.xml,Modified,1/2/2019,10/15/2019,1

Table of Contents
	1.	Input & Output Files
	2.	Variables
	3.	Snapin(s), Modules & Environment Connection(s)
	4.	Main Script
	5.	Run Metalogix Job
#>

#region Script Parameters
Param(
		[string]$jobFile = "CM-ListMigrations-CircleK-test.csv",
		[int]$maxItems = 50000,
		$newJobWait = 60
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
Start-Transcript -Path ("..\_Log FIles\CMNewListJobs-{0}.log" -f $scriptStart[1])
#endregion

#region 1.	Input & Output Files
# File Paths
$jobInPath = ("..\_Input Files\Clients\{0}" -f $jobFile)

# Import Input File(s)
If (Test-Path ($jobInPath)){
	$jobInFile = import-csv $jobInPath -Encoding Unicode
}
else {
	Write-Host
	throw ("The jobs file does not exist: {0}" -f $jobInPath)	
}
#endregion

#region 2.	Variables
# User Inputs
Write-Host
Write-Host "User Input Variables" -ForegroundColor Cyan
Write-Host "   Enter credentials for the source SharePoint environment:"
$sCred = Get-Credential -Message "Enter your credentials for the source SharePoint environment:" 

# Script Variables
$FilterTemplate = "<And IsImplicitGroup=`"True`"><FilterExpression><GreaterThanOrEqualTo Property=`"[-ColumnName-]`" CaseSensitive=`"True`" BaseFilter=`"False`" Pattern=`"[-StartDate-]`" CultureInfoName=`"en-US`"><AppliesToTypes><Type>Metalogix.SharePoint.SPListItem</Type></AppliesToTypes></GreaterThanOrEqualTo></FilterExpression><FilterExpression><LessThanOrEqualTo Property=`"[-ColumnName-]`" CaseSensitive=`"True`" BaseFilter=`"False`" Pattern=`"[-EndDate-]`" CultureInfoName=`"en-US`"><AppliesToTypes><Type>Metalogix.SharePoint.SPListItem</Type></AppliesToTypes></LessThanOrEqualTo></FilterExpression></And>"

# Output Parameters & Variables
Write-Host
Write-Host "User Inputs" -ForegroundColor Cyan
Write-Host ("   Source Credentials: {0}" -f $sCred.UserName)
Write-Host
Write-Host "Script Parameters" -ForegroundColor Cyan
Write-Host ("   Input File: {0}" -f $jobFile)
Write-Host ("   New Job Wait Time (Sec): {0}" -f $newJobWait)
#endregion

#region 3.	Snapin(s), Modules & Environment Connection(s)
# Load Custom Modules
Import-Module "..\Content Matrix\Cat.CMFunctions.psm1"
#endregion
    
#region 4.	Main Script
Write-Host
Write-Host "Creating list migration jobs" -ForegroundColor Cyan

# Looping through Sites in CSV Input to create initial migration jobs for all rows 
ForEach ($job in $jobInFile){
	# Screen Output
	Write-Host
	Write-Host ("   Job Name: {0}" -f $job.jobName) -ForegroundColor Magenta
	
	# Getting current filter information
	$curFilterColumn = $job.filterDateField
	[int]$curFilterNumBetween = $job.filterNumMoBetween
	$curFilterStart = Get-Date($job.filterStartDate)
	$curEnd = (Get-Date($job.filterEndDate)).AddMonths($curFilterNumBetween)
	
	# Retrieving List Information
	Write-Host ("      Retrieving List Info: {0}" -f $job.listPath)
	Connect-PnPOnline -Url $job.sSiteURL -Credentials $sCred
	$slist = Get-PnPList $job.listPath

	# Determine if List needs to be broken up
	If ($slist.ItemCount -le $maxItems){
		# Screen Output
		Write-Host ("      Creating Single Job. Item Count: {0}" -f $slist.ItemCount)

		# Output Site to Screen
		Write-Host ("      Initial Job Name: {0}" -f $job.jobName) -ForegroundColor DarkYellow

		# Create Current Filter
		$curFilter = $FilterTemplate.Replace("[-ColumnName-]",$curFilterColumn).Replace("[-StartDate-]",$curFilterStart.ToString("yyyy-MM-ddTHH:mm:ss.fffffffZ")).Replace("[-EndDate-]",$curEnd.ToString("yyyy-MM-ddTHH:mm:ss.fffffffZ"))

		# 5.	Run Migration Jobs
		$jobTemp = New-CMListJob -jobName $job.jobName -sourceURL $job.sSiteURL -targetURL $job.dSiteURL -list $job.listPath -filter $curFilter
		Write-Host "         Job created." -ForegroundColor Green #>
	}
	<#else {
		# Screen Output
		Write-Host ("      Breaking list up into many jobs. Item Count: {0}" -f $sList.ItemCount)

		# Create Current Filter
		$curFilterEnd = $curFilterStart.AddMonths($curFilterNumBetween)
		$curFilterEnd = $curFilterEnd.AddMilliseconds(-1)
		$curFilter = $FilterTemplate.Replace("[-ColumnName-]",$curFilterColumn).Replace("[-StartDate-]",$curFilterStart.ToString("yyyy-MM-ddTHH:mm:ss.fffffffZ")).Replace("[-EndDate-]",$curFilterEnd.ToString("yyyy-MM-ddTHH:mm:ss.fffffffZ"))
				
		# Create Job Name
		$curJob = ("{0}-{1}-{2}" -f $job.jobName, $curFilterStart.tostring("yyMMdd"),$curFilterEnd.ToString("yyMMdd"))

		# Output Site to Screen
		Write-Host ("      Initial Job Name: {0}" -f $curJob) -ForegroundColor DarkYellow

		# 5.	Run Migration Jobs
		$jobTemp = New-CMListJob -jobName $curJob -sourceURL $job.sSiteURL -targetURL $job.dSiteURL -list $job.listPath -filter $curFilter
		Write-Host "         Job created." -ForegroundColor Green #

		<# Testing Values prior to sending to content matrix
		Write-Host ("         Source Site URL: {0}" -f $job.sSiteURL)
		Write-Host ("         Target Site URL: {0}" -f $job.dSiteURL)
		Write-Host ("         List: {0}" -f $job.listPath)
		Write-Host ("         Filter File: {0}" -f $job.jobFilter)
		Write-Host ("         Filter Column: {0}" -f $curFilterColumn)
		Write-Host ("         Filter Start: {0}" -f $curFilterStart)
		Write-Host ("         Filter End: {0}" -f $curFilterEnd)
		<#Write-Host
		Write-Host "****  Filter XML ****"
		Write-Host $curFilter #
	} #>
} # ForEach ($job in $jobInFile)

# Screen Output for Initial Job Creation
Write-Host
Write-Host "Initial migration job creation complete." -ForegroundColor Green

<# Wait before creating addtional jobs, make sure servers are picking up jobs
Write-Host
Write-Host ("Waiting for {0} seconds to continue." -f $newJobWait)
Start-Sleep -Seconds $newJobWait

# Looping through Sites in CSV Input to create additional migration jobs for list with more than Max items. 
ForEach ($job in $jobInFile){
	# Screen Output
	Write-Host
	Write-Host ("   Job Name: {0}" -f $job.jobName) -ForegroundColor Magenta
	
	# Getting current filter information
	$curFilterColumn = $job.filterDateField
	[int]$curFilterNumBetween = $job.filterNumMoBetween
	$curFilterStart = (Get-Date($job.filterStartDate)).AddMonths($curFilterNumBetween)
	$curEnd = (Get-Date($job.filterEndDate)).AddMonths($curFilterNumBetween)

	# Retrieving List Information
	Write-Host ("      Retrieving List Info: {0}" -f $job.listPath)
	Connect-PnPOnline -Url $job.sSiteURL -Credentials $sCred
	$slist = Get-PnPList $job.listPath

	# Determine if List needs to be broken up
	If ($slist.ItemCount -gt $maxItems){
		# Screen Output
		Write-Host ("      Creating Addition Jobs. Item Count: {0}" -f $sList.ItemCount)

		while ($curFilterStart -le $curEnd) {
			# Create Current Filter
			$curFilterEnd = $curFilterStart.AddMonths($curFilterNumBetween).AddMilliseconds(-1)
			$curFilter = $FilterTemplate.Replace("[-ColumnName-]",$curFilterColumn).Replace("[-StartDate-]",$curFilterStart.ToString("yyyy-MM-ddTHH:mm:ss.fffffffZ")).Replace("[-EndDate-]",$curFilterEnd.ToString("yyyy-MM-ddTHH:mm:ss.fffffffZ"))

			# Create Job Name
			$curJob = ("{0}-{1}-{2}-{3}" -f $job.jobName, "Items", $curFilterStart.tostring("yyMMdd"),$curFilterEnd.ToString("yyMMdd"))

			# Output Site to Screen
			Write-Host ("      Additional Job Name: {0}" -f $curJob) -ForegroundColor DarkYellow

			# 5.	Run Migration Jobs
			New-CMListItemsJob -jobName $curJob -sourceURL $job.sSiteURL -targetURL $job.dSiteURL -list $job.listPath -filter $curFilter
			Write-Host "         Job created." -ForegroundColor Green #

			<# Testing Values prior to sending to content matrix
			Write-Host ("         Source Site URL: {0}" -f $job.sSiteURL)
			Write-Host ("         Target Site URL: {0}" -f $job.dSiteURL)
			Write-Host ("         List: {0}" -f $job.listPath)
			Write-Host ("         Filter File: {0}" -f $job.jobFilter)
			Write-Host ("         Filter Column: {0}" -f $curFilterColumn)
			Write-Host ("         Filter Start: {0}" -f $curFilterStart)
			Write-Host ("         Filter End: {0}" -f $curFilterEnd)
			Write-Host
			<#Write-Host "****  Filter XML ****"
			Write-Host $curFilter #
					
			# Increment Filter Start Date
			$curFilterStart = $curFilterStart.AddMonths($curFilterNumBetween)
		}
	}
} # ForEach ($job in $jobInFile)

# Screen Output for Additional Job Creation
Write-Host
Write-Host "Additional migration job creation complete." -ForegroundColor Green
#endregion #>

#region Completed Screen Output
# Disconnect from SharePoint
Disconnect-PnPOnline

# Output Script Info
Complete-CatScript $scriptStart[0]

# Stop logging
Stop-Transcript
#endregion