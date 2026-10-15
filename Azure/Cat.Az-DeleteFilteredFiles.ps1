<#  Cat.Az-DeleteFilteredFiles.ps1

This script is intended to delete a subset of files in an Azure Storage Account's blob
contain based on filter criteria (Starts w/ "UserData" but not "UserDataUserData"). This script
was written specifically to clean up files in the Content Re-platforming file share discovery
dashboard. To update a portion of the dashboard, you need to clean up the desired old files and 
replace them with files generated from a new scan.

Dependencies
	-	Az Modules
	-	General Functions - ..\General\Cat.GeneralFunctions.psm1

Call script with parameters
	cd C:\GitHub\CatPowerShell\Azure\
	.\Cat.Az-DeleteFilteredFiles.ps1 -storageAccount "stdatamigrationeus2" -resourceGroup "rg-prac-prod-file-share-migration-discovery-01" -containerName "datamigration" -directoryPath "bronze/D4C/2509100802/directories" -filterMatch "UserData" -filterNotMatch "UserDataUserData"

Table of Contents
	1.	Snapin(s), Modules & Environment Connection(s)
	2.	Variables
	3.	Main Script
#>

#region Script Parameters
Param(	[string] $storageAccount = "{azure storage account name}",
		[string] $resourceGroup = "{azure resource group name}",
		[string] $containerName = "{azure blob container name}",
		[string] $directoryPath = "{folder path with /s}",
		[string] $filterMatch = "{file name starts with}",
		[string] $filterNotMatch = "{file name doesn't start with}"
)
#endregion Script Parameters

#region Start Date / Time
# Remove Existing Catapult Modules to ensure PS session has most up to date functions
Get-Module Cat.* | Remove-Module

# Load General Functions
Import-Module "..\General\Cat.GeneralFunctions.psm1"

# Get Start Date / Time
$scriptStart = Start-CatScript

# Start Transcript
Start-Transcript -Path ("..\_Log Files\Cat.Az-DeleteFilteredFiles-{1}.log" -f $clientTenant, $scriptStart[1])
#endregion Start Date / Time

#region 1.	Snapin(s), Modules & Environment Connection(s)
# Load MS / PnP PowerShell Modules
Import-Module Az
#endregion Snapin(s), Modules & Environment Connection(s)

#region 2.	Variables
Write-Host
Write-Host "*** Deleting files can be permanent based on you storage account settings ***" -ForegroundColor Red
Write-Host
$decision = Read-Host "   Would you like to report or delete the files that meet the filter criteria? (Delete or Report)"

# if user chose to proceed
switch ($decision.ToLower()) {
	delete {$delete = $true}
	report {$delete = $false}
	Default {$delete = $false}
}

# Output Parameters & Variables
Write-Host
Write-Host "Script Parameters" -ForegroundColor Cyan
Write-Host ("   Storage Account Name: {0}" -f $storageAccount)
Write-Host ("   Resource Group Name: {0}" -f $resourceGroup)
Write-Host ("   Blob Container Name: {0}" -f $containerName)
Write-Host ("   Directory Path: {0}" -f $directoryPath)
Write-Host ("   File Name Starts With: {0}" -f $filterMatch)
Write-Host ("   File Name Does Not Start With: {0}" -f $filterNotMatch)
Write-Host
Write-Host "User Inputs" -ForegroundColor Cyan
Write-Host ("   Script Behavior: {0}" -f $decision)
Write-Host ("   Delete Variable: {0}" -f $delete)
#endregion
    
#region 3.	Main Script
Write-Host
Write-Host ("{0} - Deleting blob files" -f (Get-Date -Format "yy-MM-dd HH:mm:ss")) -ForegroundColor Cyan

# Connect to Azure
Write-Host ("   {0} - Connecting to Azure Account" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"))
$temp = Connect-AzAccount

# Get the storage account context
Write-Host ("   {0} - Retrieving Storage Account Context" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"))
$ctx = (Get-AzStorageAccount -ResourceGroupName $resourceGroup -Name $storageAccount).Context

# List all blobs in the container that start with the target directory path
Write-Host ("   {0} - Retrieving files to delete" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"))
$blobs = Get-AzStorageBlob -Container $containerName -Context $ctx -Prefix $directoryPath

# Filter blobs
Write-Host ("   {0} - Filtering files" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"))
$filteredBlobs = $blobs | Where-Object {
    ($_.Name -match ("^" + $directoryPath + "/" + $filterMatch)) -and ($_.Name -notmatch ("^" + $directoryPath + "/" + $filterNotMatch))
}

# Delete the filtered blobs
Write-Host
Write-Host ("   {0} - Deleting blob files" -f (Get-Date -Format "yy-MM-dd HH:mm:ss")) -ForegroundColor Magenta
foreach ($blob in $filteredBlobs) {
	if ($delete) {
		Write-Host ("      {0} - Deleting file: {1}" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"), $blob.Name)
    	Remove-AzStorageBlob -Container $containerName -Blob $blob.Name -Context $ctx
	} else {
		Write-Host ("      {0} - File to be Deleted: {1}" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"), $blob.Name)
	}
}

Write-Host
if ($delete) {
	Write-Host ("   {0} - Total Files Deleted: {1}" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"), $filteredBlobs.Count) -ForegroundColor Green
} else {
	Write-Host ("   {0} - Total Files to be Deleted: {1}" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"), $filteredBlobs.Count) -ForegroundColor Green
}
#endregion Main Script

#region Completed Script
# Disconnect from Services
Write-Host
Write-Host "Disconnecting from Services:" -ForegroundColor Cyan
Write-Host ("   Disconnecting from Azure")
$temp = Disconnect-AzAccount

# Output Script Info
Complete-CatScript $scriptStart[0]

# Stop logging
Stop-Transcript
#endregion Completed Script