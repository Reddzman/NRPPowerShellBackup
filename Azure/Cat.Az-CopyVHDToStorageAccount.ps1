<#  Cat.Az-CopyVHDToStorageAccount.ps1

This script...

Dependencies
	-	PnP.PowerShell Modules
	-	SPO PowerShell Modules
	-	Microsoft Teams Modules
	-	ShareGate PowerShell Modules
	-	Metalogix PowerShell Snapins
	-	General Functions - ..\General\Cat.GeneralFunctions.psm1
	-	SQL Server Command Utility - https://www.microsoft.com/en-us/download/details.aspx?id=53591
	-	Input CSV File - Example: ..\_Input FIles\Testing\Cat.Tech-Action-Test.csv
	-	Input XML File - Example: ..\_Input Files\Testing\Cat.Tech-Action-Test.xml

Call script with parameters
	.\Cat.Az-CopyVHDToStorageAccount.ps1 -subscriptionId "dca711d5-ed08-42b7-9cb6-580b73bc445b" -resourceGroupName "MSDN-SPLab" -diskName "SP13-001-OsDisk" -sasExpiryDuration 3600 -storageAccountName "amsdnsplab4047" -storageContainerName "vhds" -storageAccountKey "qMx+guMfzuFfPxDLakL7rJ1sD1mpFbzJKwPRwXHqwFk8G1FWiy+aqoVg3LRWTUQsc9zIanhW/f/p3TvGlJA7XQ==" -destinationVHDFileName "SP13-001-OsDisk-Backup.vhd" -useAzCopy 0

Table of Contents
	1.	Snapin(s), Modules & Environment Connection(s)
	2.	Input & Output Files
	3.	Variables
	4.	Main Script
	5.	CSV Loop Example
	6.	...
#>

#region Script Parameters
Param(	[string] $subscriptionId = "{yourSubscriptionId}", # Provide the subscription Id of the subscription where managed disk is created
		[string] $resourceGroupName = "{yourResourceGroupName}", # Provide the name of your resource group where managed is created
		[string] $diskName = "{yourDiskName}",# Provide the managed disk name 
		[string] $sasExpiryDuration = "3600", # Provide Shared Access Signature (SAS) expiry duration in seconds e.g. 3600. https://docs.microsoft.com/en-us/Az.Storage/storage-dotnet-shared-access-signature-part-1
		[string] $storageAccountName = "{yourstorageaccountName}", # Provide storage account name where you want to copy the underlying VHD of the managed disk.
		[string] $storageContainerName = "{yourstoragecontainername}", # Name of the storage container where the downloaded VHD will be stored
		[string] $storageAccountKey = "{yourStorageAccountKey}", # Provide the key of the storage account where you want to copy the VHD of the managed disk. 
		[string] $destinationVHDFileName = "{yourvhdfilename}", # Provide the name of the destination VHD file to which the VHD of the managed disk will be copied.
		[boolean] $useAzCopy = 0 # Set the value to 1 to use AzCopy tool to download the data. This is the recommended option for faster copy. If you set the value to 0 then Start-AzStorageBlobCopy will be used. Azure storage will asynchronously copy the data. Ensure that AzCopy is downloaded in the same folder as this file (https://docs.microsoft.com/en-us/azure/storage/common/storage-use-azcopy-v10)
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
Start-Transcript -Path ("..\_Log Files\Az-CopyVHDToStorageAccount-{0}.log" -f $scriptStart[1])
#endregion

#region 1.	Snapin(s), Modules & Environment Connection(s)
# Load MS / PnP PowerShell Modules
Import-Module Az -erroraction SilentlyContinue
#endregion

#region 3.	Variables
# Output Parameters & Variables
Write-Host
Write-Host "Script Parameters" -ForegroundColor Cyan
Write-Host ("   Subscription ID: {0}" -f $subscriptionId)
Write-Host ("   Resource Group Name: {0}" -f $resourceGroupName)
Write-Host ("   Disk Name: {0}" -f $resourceGroupName)
Write-Host ("   SAS Expiry Duration: {0}" -f $sasExpiryDuration)
Write-Host ("   Storage Accoount Name: {0}" -f $storageAccountName)
Write-Host ("   Storage Container Name: {0}" -f $storageContainerName)
Write-Host ("   Storage Account Key: {0}" -f $storageContainerName)
Write-Host ("   Destination VHD File Name: {0}" -f $destinationVHDFileName)
Write-Host ("   Use AzCopy: {0}" -f $useAzCopy)
#endregion
    
#region 4.	Main Script
# Connecting to Azure
Connect-AzAccount

# Set the context to the subscription Id where managed disk is created
Select-AzSubscription -SubscriptionId $SubscriptionId

#Generate the SAS for the managed disk 
$sas = Grant-AzDiskAccess -ResourceGroupName $ResourceGroupName -DiskName $diskName -DurationInSecond $sasExpiryDuration -Access Read 

#Create the context of the storage account where the underlying VHD of the managed disk will be copied
$destinationContext = New-AzStorageContext -StorageAccountName $storageAccountName -StorageAccountKey $storageAccountKey

#Copy the VHD of the managed disk to the storage account
if($useAzCopy -eq 1)
{
    $containerSASURI = New-AzStorageContainerSASToken -Context $destinationContext -ExpiryTime(get-date).AddSeconds($sasExpiryDuration) -FullUri -Name $storageContainerName -Permission rw
    .\azcopy.exe copy $sas.AccessSAS $containerSASURI

}else{

    Start-AzStorageBlobCopy -AbsoluteUri $sas.AccessSAS -DestContainer $storageContainerName -DestContext $destinationContext -DestBlob $destinationVHDFileName
}
#endregion

#region Completed Script
# Output Script Info
Complete-CatScript $scriptStart[0]

# Stop logging
Stop-Transcript
#endregion