<#  Cat.SPOnPremForceCheckIn.ps1
This script...

Forces a check in of files for each library in Cat.SG-SPOnPremListMigrations-Client.csv that has forceCheckIn set to "true". 
Set the $pastDate parameter to limit the files checked in to only those with a last modified date that is greater than or 
equal to the $pastDate

Dependencies
	-	PnP.PowerShell Modules
	-	SPO PowerShell Modules
	-	General Functions - ..\General\Cat.GeneralFunctions.psm1
	-	Input CSV File - Example: ..\_Input FIles\Testing\Cat.SG-SPOnPremListMigrations-Client.csv

Call script with parameters
	.\Cat.SPOnPremForceCheckIn.ps1 -jobFile "Cat.SG-SPOnPremListMigrations-Client.csv" -pastDate "2024-06-01"

Table of Contents
	1.	Input & Output Files
	2.	Variables
	3.  Custom Modules & SharePoint CSOM Assemblies
	4.  Functions
	5.	Main Script
	6.	Force Check in Job
#>

#region Script Parameters
Param(
	[String] $jobFile = "Cat.SG-SPOnPremListMigrations-Client.csv",
	[String] $pastDate = "2000-01-01" # Checked out files will be checked in if it is equal to or greater than this date. Files that have no checked in versions will be ignored.
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
Start-Transcript -Path ("..\_Log Files\{0}.SG-SPOnPremForceCheckIn-{1}.log" -f $destTenant, $scriptStart[1])
#endregion

#region 1.	Input & Output Files
# File Paths
$inPath = ("..\_Input Files\{0}" -f $jobFile)
$errorPath = ("..\_Output Files\{0}.SPOnPremForceCheckInErrors-{1}.log" -f $destTenant, $scriptStart[1])

# Import Input File(s)
# Load CSV file
If (Test-Path ($inPath)){
	$inFile = Import-Csv $inPath -Encoding Unicode
}
else {
	Write-Host
	throw ("The jobs file does not exist: {0}" -f $inPath)	
}
#endregion

#region 2.	Variables
# Script Variables
$siteWaitMaxCount = 20
$siteWaitPause = 30 # Seconds
$pastDateToCheckIn = Get-Date $pastDate # Formats the $pastDate to a comparable date
$rowLimit = 2000	# Used in query pagination. Limits the number of items return in each batch
$specialCharacters = "[~``!@#$%^&*()_-{|?/\<>:'}]"

# User Inputs
Write-Host
Write-Host "Enter Connection Information" -ForegroundColor Cyan
Write-Host "   Enter ShareGate credentials for the SOURCE SharePoint environment:"
$srcSGCreds = Get-Credential -Message "Enter ShareGate credentials for the SOURCE environment:" 

# Variable Output
Write-Host
Write-Host "Script Parameters" -ForegroundColor Cyan
Write-Host ("   Job File: {0}" -f $jobFile)
Write-Host ("   Site Creation Wait Max Iterations: {0}" -f $siteWaitMaxCount)
Write-Host ("   Site Creation Wait Pause Duration: {0}" -f $siteWaitPause)
Write-Host ("   Past Date to Check In Files: {0}" -f $pastDateToCheckIn)
Write-Host ("   Query Pagination Row Limit: {0}" -f $rowLimit)
Write-Host ("   Special Characters: {0}" -f $specialCharacters)
Write-Host
Write-Host "User Inputs" -ForegroundColor Cyan
Write-Host ("   Source Credentials: {0}" -f $srcSGCreds.UserName)
#endregion

#region 3. Load Custom Modules & SharePoint CSOM Assemblies
# Load Custom Modules
Import-Module "..\Reporting\Cat.ReportingFunctions.psm1"
#endregion

# Load SharePoint CSOM Assemblies
Add-Type -Path "c:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll"
Add-Type -Path "c:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Runtime.dll"
#endregion

#region 4. Functions
<#  Log-MigrationError
This function creates a log file for errors that happen throughout the migration process.

Command: Log-MigrationError -funcJob $job -funcAction "Creating Site" -funcException $_.Exception -funcDetails $_.ErrorDetails -funcStackTrace $_.ScriptStackTrace -funcLogPath $errorPath

Parameters
    funcJob = Job information from Migraiton CSV file
    funcAction = Action that resulted in error
    funcException = Error Exception
    funcDetails = Error Details
    funcStackTrace = Error Script Stack Trace
    funcLogPath = Path to log file
#>
function Log-MigrationError
{ 
	param ( 
        $funcJob,
        $funcAction,
        $funcException,
        $funcDetails,
        $funcStackTrace,
        $funcLogPath
    )
    Write-Host
    Write-Host ("{0} - {1}: {2}" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"),$funcAction,$_.Exception) -ForegroundColor Red
    $err = $funcJob
    $err | Add-Member -MemberType NoteProperty -Name "errorAction" -Value $funcAction
    $err | Add-Member -MemberType NoteProperty -Name "errorException" -Value ((($funcException -replace $specialCharacters,'') -replace "`n",' ') -replace "`r",' ')
    $err | Add-Member -MemberType NoteProperty -Name "errorDetails" -Value ((($funcDetails -replace $specialCharacters,'') -replace "`n",' ') -replace "`r",' ')
    $err | Add-Member -MemberType NoteProperty -Name "errorStackTrace" -Value ((($funcStackTrace -replace $specialCharacters,'') -replace "`n",' ') -replace "`r",' ')
    $err | Export-Csv -Path $funcLogPath -NoTypeInformation -Append
    Clear-Variable err
}
#endregion
    
#region 5.	Main Script
Write-Host
Write-Host ("Running Force Check In Files for job list: {0}" -f $jobFile) -ForegroundColor Cyan

# Looping through input file
:JobLoop ForEach ($job in $inFile) {

	#region 6 Force Check In Files
	try {
		if($job.forceCheckIn) {

			# Connect to SharePoint OnPrem site
			#$context = Get-PnPContext
			$context = New-Object Microsoft.SharePoint.Client.ClientContext($job.srcSiteURL)
			$context.Credentials = $srcSGCreds

			# Get the web
			$web = $context.Web
			
			$context.Load($web)
			$context.ExecuteQuery()
			
			# Get the specific list
			$list = $web.Lists.GetByTitle($job.srcListName)
			
			$context.Load($list)
			$context.ExecuteQuery()

			<#
			#Get field names 
			$fields = $list.Fields
					$context.Load($fields)
					$context.ExecuteQuery()
			foreach ($field in $fields) {
				Write-Host $field.InternalName
			}
			#>

			# Get only the following libraries: 101 = Document Library, 109 = Picture Library, 851 Asset Library
			if(($list.BaseTemplate -eq "101" -or $list.BaseTemplate -eq "109" -or $list.BaseTemplate -eq "851") -and $list.Hidden -eq $False) {

				# Get all checked-out items in the list
				$query = New-Object Microsoft.SharePoint.Client.CamlQuery
				$query.ViewXml = "<View Scope='RecursiveAll'><Query><Where><IsNotNull><FieldRef Name='CheckoutUser' /></IsNotNull></Where></Query><RowLimit Paged='TRUE'>$rowLimit</RowLimit></View>"

				# Batch list items
				$counter=1
				Do {
					$items = $list.GetItems($query)

					$context.Load($items)
					$context.ExecuteQuery()

					$query.ListItemCollectionPosition = $items.ListItemCollectionPosition

					#Loop through each file
					foreach ($item in $items | Where {$_.FileSystemObjectType -eq "File"}) {
						try {

							#Display a Progress bar
							Write-Progress -Activity "Scanning Files..." -Status "for Checked-Out Status '$($item.FieldValues.FileRef)' ($counter of $($list.ItemCount))" -PercentComplete (($counter / $list.ItemCount) * 100)

							# Force check-in the item if it meets the date requirement
							$isCheckedOut = if($item["CheckoutUser"]){"Yes"}else{"No"}
							Write-Host
							Write-Host ("File: {0} - Check Out Status: {1} - Last Modified: {2}" -f $item["FileLeafRef"], $isCheckedOut, $item["Modified"]) -ForegroundColor Cyan
							if($item["Modified"] -ge $pastDateToCheckIn) {
								$item.File.CheckIn("Checked in by migration script", [Microsoft.SharePoint.Client.CheckinType]::MajorCheckIn)
								Write-Host ("Checked in by script: {0}" -f $item["FileLeafRef"])  -ForegroundColor Green
								
							} else {
								Write-Host ("...Not Checked in by script: {0} - Reason: Last modified older than {1}" -f $item["FileLeafRef"], $pastDateToCheckIn) -ForegroundColor Yellow
							}
							$counter++
							$context.ExecuteQuery()							
							} catch {
								Write-Host "Failed to check in: $($item["FileLeafRef"]) - $($_.Exception.Message)"  -ForegroundColor Red
		
							}
						} catch {
							Log-MigrationError -funcJob $job -funcAction "Check in File" -funcException $_.Exception -funcDetails $_.ErrorDetails -funcStackTrace $_.ScriptStackTrace -funcLogPath $errorPath
						}
					}  					
				While($null -ne $query.ListItemCollectionPosition) 
			}
					
		} else {
			Write-Host "Job item {0} is not a library" -f $job.srcListName
		}    
	} catch {
		# Output Error Information
		Continue JobLoop  
	}        
	#end region #>
}

# Script complete output
Write-Host
Write-Host "Force Check in Files Job Completed. " -ForegroundColor Green
#endregion


#region Completed Script
# Output Script Info
Complete-CatScript $scriptStart[0]

# Stop logging
Stop-Transcript
#endregion