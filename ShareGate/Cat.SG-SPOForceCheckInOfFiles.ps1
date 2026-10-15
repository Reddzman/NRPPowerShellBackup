<#  Cat.SPO-ForceCheckInOfFiles.ps1

This script...

Forces a check in of files for each library in Cat.SPO-SPOnPremListMigrations-Client.csv that has forceCheckIn set to "true". 
Set the $pastDate parameter to limit the files checked in to only those with a last modified date that is greater than or 
equal to the $pastDate

Dependencies
	-	PnP.PowerShell Modules
	-	SPO PowerShell Modules
	-	General Functions - ..\General\Cat.GeneralFunctions.psm1
	-	Input CSV File - Example: ..\_Input FIles\Testing\Cat.SG-SPOnPremListMigrations-Client.csv

Call script with parameters
	.\Cat.SG-ForceCheckInOfFiles.ps1  -inputFile Clients\Cat.SPO-SPOnPremListMigrations-Client.csv -client "ClientNameNoSpacesOrAbriviation"

Table of Contents
	1.	Snapin(s), Modules & Environment Connection(s)
	2.	Input & Output Files
	3.	Variables
	4.	Main Script
	5.	Force Check in Job
#>

#region Script Parameters
Param(	[string] $jobFile = "Cat.SG-SPOnPremListMigrations-Client.csv",
        [string] $pastDate = "6/17/2024"    # Date you want to ignore files for check in. Leave Blank to check in ALL files that are currently checked out.
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
Start-Transcript -Path ("..\_Log Files\{0}.SPO-ForceCheckInOfFiles-{1}.log" -f $clientTenant, $scriptStart[1])
#endregion

#region 1.	Snapin(s), Modules & Environment Connection(s)
# Load MS / PnP PowerShell Modules
Get-Module -ListAvailable | Where-Object{($_.Name -like "*SharePoint*") -or ($_.Name -like "*PnP*") } | Remove-Module
Import-Module PnP.PowerShell
Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue
#endregion

#region 2.	Input & Output Files
# File Paths
$inPath = ("..\_Input Files\{0}" -f $jobFile )
$errorPath = ("..\_Output Files\{0}.SPO-ForceCheckInOfFilesErrors-{1}.log", $scriptStart[1])

# Import Input File(s)
# Load CSV file
If (Test-Path ($inPath)){
	$inFile = Import-Csv $inPath -Encoding Unicode
}
else {
	Write-Host
	throw ("The input file does not exist: {0}" -f $inPath)	
}
[xml]$inFile = Get-Content $inPath # Importing XML File
$inFile = Get-Content $inPath -Raw # Importing JSON File
#endregion

#region 3.	Variables

# Script Variables
$pageSize = 1000 # Number of files retrieved in batch

# User Inputs
Write-Host
Write-Host "Enter Connection Information" -ForegroundColor Cyan
Write-Host "   Enter ShareGate credentials for the SOURCE SharePoint environment:"
$srcSGCreds = Get-Credential -Message "Enter ShareGate credentials for the SOURCE environment:" 

# Output Parameters & Variables
Write-Host
Write-Host "Script Parameters" -ForegroundColor Cyan
Write-Host ("   Job File: {0}" -f $jobFile )
Write-Host
Write-Host "Script Variables" -ForegroundColor Cyan
Write-Host ("   pageSize: {0}" -f $pageSize)
Write-Host
Write-Host ("   pastDate: {0}" -f $pastDate)
Write-Host
#endregion
    
#region 4.	Main Script

#region Connecting To SP Online
Write-Host ("   {0} - Connecting to SharePoint Online Admin site: {0}" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"), $list.siteTitle)
try {
	if ($clientAppID) {
		Connect-PnPOnline $spoAdminUrl -tenant $tenantDomain -ClientId $clientAppID -Thumbprint $cert.Thumbprint -ValidateConnection -ErrorAction Stop # PnP Modules w/ App Reg
	} else {
		Connect-PnPOnline -Url $spoAdminUrl -Interactive -ValidateConnection -ErrorAction Stop # PnP Modules w/o App Reg
	} 
}
catch {
	Write-CatErrorCSV -funcCSVItem $list -funcAction "Connecting to SharePoint" -funcException $_.Exception -funcDetails $_.ErrorDetails -funcStackTrace $_.ScriptStackTrace -funcFilePath $errorPath
	Continue listLoop
}
#endregion

#region 5. Force Check in Job
:JobLoop foreach ($job in $inFile) {
    Write-Host
    Write-Host ("      Loading Source Site: {0}" -f $job.srcSiteTitle)
    try {
        Connect-Site -Url $job.srcSiteURL -Credential $srcSGCreds -ErrorAction Stop
    }
    catch {
        # Output Error Information
        Log-MigrationError -funcJob $job -funcAction "Loading Source Site" -funcException $_.Exception -funcDetails $_.ErrorDetails -funcStackTrace $_.ScriptStackTrace -funcLogPath $errorPath
        Continue JobLoop
    }

    # Perform the force check in
    if ($job.forceCheckIn.ToLower() -eq "true") {
        try {

            # Get files from the current library omitting folders
            If($pastDate -ne $Null -and $pastDate -ne '') {
                # Get all checked out files that are greater than or equal to $pastDate
                $files = Get-PnPListItem -List $job.srcListName -PageSize $pageSize | Where-Object { $_.FileSystemObjectType -eq 'File' -and $_.CheckoutUser -ne $Null -and $_.Modified -ge $pastDate }
            } else {
                # Get all checked out files 
                $files = Get-PnPListItem -List $job.srcListName -PageSize $pageSize | Where-Object { $_.FileSystemObjectType -eq 'File' -and $_.CheckoutUser -ne $Null }
            }

            # region Loop through each file in library
            :Files foreach ($item in $files) {
                $fileUrl = $item.FieldValues.FileRef # Get the relative URL of the file

                try {
                    # Check in the file
                    Set-PnPFileCheckedIn -Url $fileUrl -CheckinType Major -Comment "Checked in by System Account"                    
                    Write-Host "Checked in file: $fileUrl"
                } catch {
                    Write-CatErrorCSV -funcCSVItem $item -funcAction "Error checking in file: $fileUrl" -funcException $_.Exception -funcDetails $_.ErrorDetails -funcStackTrace $_.ScriptStackTrace -funcFilePath $errorPath
                    Continue Files
                }
            }
            #endregion
        } catch {
            Write-CatErrorCSV -funcCSVItem $item -funcAction "Error getting library: $item.srcListName" -funcException $_.Exception -funcDetails $_.ErrorDetails -funcStackTrace $_.ScriptStackTrace -funcFilePath
            $errorPath
            Continue Libraries
        }
    }      
}
#endregion

#region Completed Script
# Disconnect from Services
Write-Host
Write-Host "Disconnecting from Services:" -ForegroundColor Cyan
Write-Host ("   Disconnecting from SharePoint Online: {0}" -f $spoAdminURL)
Disconnect-PnPOnline

# Output Script Info
Complete-CatScript $scriptStart[0]

# Stop logging
Stop-Transcript
#endregion