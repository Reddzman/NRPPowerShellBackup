<#  Cat.PnP-ListModifiedDates.ps1

This script this script iterates through a csv of sites/lists and gets the date they where created and the last modified date. 
It is then outputed to another csv file.  It was written for 2013 but will likely work for over versions that support PnP. 
The import module command should reflect the version of PnP powershell that relates to your version of SharePoint.

Dependencies
	-	PnP PowerShell https://docs.microsoft.com/en-us/powershell/sharepoint/sharepoint-pnp/sharepoint-pnp-cmdlets?view=sharepoint-ps
	-	..\General\Cat.PnP-GeneralFunctionsFunctions.psm1
	-	Input CSV File in Input Files folder - Script Parameter - "Cat.ListModifiedDates-Client.csv"

Call script with parameters
	.\Cat.Tech-ScriptName.ps1 -InputFile InputFileName.csv

CSV Example
	Site,siteURL,SPOUrl
	SiteName,https://sharepoint.client.com/sitedirectory/law,https://client.sharepoint.com/sites/law

Table of Contents
	1.	Input & Output Files
	2.	Variables
	3.	Snapin(s), Modules & Environment Connection(s)
	4.	Main Script
#>

#region Script Parameters
Param(	[String] $inputFile = "CM-ListMigrations-CircleK.csv"
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
Start-Transcript -Path ("..\_Log FIles\ListModifiedDates-{0}.log" -f $scriptStart[1])
#endregion

#region 1.	Input & Output Files
# File Paths
$inPath = ("..\_Input Files\Clients\{0}" -f $inputFile)
$outPath = ("..\_Output Files\Clients\{0}{1}.csv" -f $inputFile.Replace(".csv","-"), $scriptStart[1])

# Create Output File(s)
Out-File -FilePath $outPath -InputObject "jobName,sSiteURL,dSiteURL,listPath,jobFilter,filterDateField,filterStartDate,filterEndDate,filterNumMoBetween,itemCount" -Encoding unicode

# Import Input File(s)
# Load CSV file
If (Test-Path ($inPath)){
	$inFile = Import-Csv $inPath -Encoding unicode
}
else {
	Write-Host
	throw ("The jobs file does not exist: {0}" -f $inPath)	
}
#endregion

#region 2.	Variables
# User Inputs
Write-Host
Write-Host "User Input Variables" -ForegroundColor Cyan
Write-Host "   Enter credentials for SharePoint:"
$spCred = Get-Credential -Message "Enter your credentials for SharePoint:" 

# Output Parameters & Variables
Write-Host
Write-Host "Script Parameters" -ForegroundColor Cyan
Write-Host ("   Input File: {0}" -f $inputFile)
Write-Host
Write-Host "User Inputs" -ForegroundColor Cyan
Write-Host ("   Credentials: {0}" -f $spCred.UserName)
#endregion

#region 3.	Snapin(s), Modules & Environment Connection(s)
# Load MS / PnP PowerShell Modules
Import-Module SharePointPnPPowerShell2013 -WarningAction SilentlyContinue
#endregion
    
#region 4.	Main Script
Write-Host
Write-Host "Retrieving List Modified Dates" -ForegroundColor Magenta

ForEach ($job in $inFile){
	# Connecting To SP Site
	Write-Host
	Write-Host ("   Connecting to SharePoint Site: {0}" -f $job.sSiteURL) -ForegroundColor DarkBlue
	Connect-PnPOnline -Url $job.sSiteURL -Credentials $spCred

	Write-Host ("      Retrieving List Information: {0}" -f $job.listPath)
	$curList = Get-PnPList $job.listPath

	Write-Host ("      Writing Modified Dates to CSV: {0}" -f $job.listPath)
	Out-File -FilePath $outPath -Append -InputObject ("{0},{1},{2},{3},{4},{5},{6},{7},{8},{9}" -f 
		$job.jobName,
		$job.sSiteURL,
		$job.dSiteURL,
		$job.listPath,
		$job.jobFilter,
		$jbo.filterDateField,
		(Get-Date $curList.Created -Format "MM/dd/yyyy"),
		(Get-Date $curList.LastItemModifiedDate -Format "MM/dd/yyyy"),
		$job.filterNumMoBetween,
		$curList.ItemCount
	) # Out-File -FilePath $outPath
} # ForEach ($job in $inFile)
#endregion

#region Completed Script
# Disconnect from SharePoint
Disconnect-PnPOnline

# Output Script Info
Complete-CatScript $scriptStart[0]

# Stop logging
Stop-Transcript
#endregion