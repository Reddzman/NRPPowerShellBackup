<#  Cat.Az-CopyFilteredFiles.ps1

This script is intended to copy a subset of files in an Azure Storage Account's blob
containers based on filter criteria (Starts w/ "UserData" but not "UserDataUserData"). It copies
documents from a directory to another directory. This script was written specifically to
clean up files in the Content Re-platforming file share discovery dashboard. To update
a portion of the dashboard, you need to clean up the desired old files and replace them
with files generated from a new scan.

Dependencies
	-	Az Modules
	-	General Functions - ..\General\Cat.GeneralFunctions.psm1

Call script with parameters
	cd C:\GitHub\CatPowerShell\Azure\
	.\Cat.Az-CopyFilteredFiles.ps1 -storageAccount "stdatamigrationeus2" -resourceGroup "rg-prac-prod-file-share-migration-discovery-01" -containerName "datamigration" -srcDirectoryPath "bronze/D4C/RawData/2509081610-d4c-Discovery-UserData-250908-02/directories" -destDirectoryPath "bronze/D4C/2509100802/directories" -filterMatch "UserData" -filterNotMatch "UserDataUserData"

Table of Contents
	1.	Snapin(s), Modules & Environment Connection(s)
	2.	Variables
	3.	Main Script
#>

#region Script Parameters
Param(	[string] $storageAccount = "{azure storage account name}",
		[string] $resourceGroup = "{azure resource group name}",
		[string] $containerName = "{azure blob container name}",
		[string] $srcDirectoryPath = "{folder path with /s}",
		[string] $destDirectoryPath = "{destination folder path with /s}",
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
Start-Transcript -Path ("..\_Log Files\Cat.Az-CopyFilteredFiles-{1}.log" -f $clientTenant, $scriptStart[1])
#endregion Start Date / Time

#region 1.	Snapin(s), Modules & Environment Connection(s)
# Load MS / PnP PowerShell Modules
Import-Module Az
#endregion Snapin(s), Modules & Environment Connection(s)

#region 3.	Variables
# Output Parameters & Variables
Write-Host
Write-Host "Script Parameters" -ForegroundColor Cyan
Write-Host ("   Storage Account Name: {0}" -f $storageAccount)
Write-Host ("   Resource Group Name: {0}" -f $resourceGroup)
Write-Host ("   Blob Container Name: {0}" -f $containerName)
Write-Host ("   Source Directory Path: {0}" -f $srcDirectoryPath)
Write-Host ("   Destination Directory Path: {0}" -f $destDirectoryPath)
Write-Host ("   File Name Starts With: {0}" -f $filterMatch)
Write-Host ("   File Name Does Not Start With: {0}" -f $filterNotMatch)
#endregion
    
#region 4.	Main Script
Write-Host
Write-Host ("{0} - Coping blob files" -f (Get-Date -Format "yy-MM-dd HH:mm:ss")) -ForegroundColor Cyan

# Connect to Azure
Write-Host ("   {0} - Connecting to Azure Account" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"))
$temp = Connect-AzAccount

# Get the storage account context
Write-Host ("   {0} - Retrieving Storage Account Context" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"))
$ctx = (Get-AzStorageAccount -ResourceGroupName $resourceGroup -Name $storageAccount).Context

# List all blobs in the container that start with the target directory path
Write-Host ("   {0} - Retrieving files to copy" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"))
$blobs = Get-AzStorageBlob -Container $containerName -Context $ctx -Prefix $srcDirectoryPath

# Filter blobs
Write-Host ("   {0} - Filtering files" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"))
$filteredBlobs = $blobs | Where-Object {
    ($_.Name -match ("^" + $srcDirectoryPath + "/" + $filterMatch)) -and ($_.Name -notmatch ("^" + $srcDirectoryPath + "/" + $filterNotMatch))
}

# Delete the filtered blobs
Write-Host
Write-Host ("   {0} - Coping files" -f (Get-Date -Format "yy-MM-dd HH:mm:ss")) -ForegroundColor Magenta
foreach ($blob in $filteredBlobs) {
    $destBlobName = $blob.Name -replace "^$srcDirectoryPath", $destDirectoryPath
    Write-Host ("      {0} - {1} to {2}" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"), $blob.Name, $destBlobName)
    $temp = Start-AzStorageBlobCopy -SrcBlob $blob.Name -SrcContainer $containerName -DestBlob $destBlobName -DestContainer $containerName -Context $ctx
}

Write-Host
Write-Host ("   {0} - Total Files Copied: {1}" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"), $filteredBlobs.Count) -ForegroundColor Green
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