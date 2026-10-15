<#  Cat.PnP-RestoreRecycleBinData.ps1

This script iterates through a CSV file to restore files from the recycle bin in SharePoint Online based on a specific date when a specific user deleted items.

Dependencies
	-	PnP.PowerShell
	-	General Functions - ..\General\Cat.PnP-GeneralFunctionsFunctions.psm1
	-	Input CSV File

CSV Notes
	-	For one date and one user then you only need one row in the CSV file, otherwise add as many rows as needed for the different users and dates you want to restore from
	-	The date format in the CSV file must be MMddyyyy (IE. 01012022 for January 1st, 2022)
	-	The user email in the CSV file must match the email address of the user who deleted the item in order for the script to restore the item

Call script with parameters
	cd C:\GitHub\CatPowerShell\Configuration\

	Run in test mode first to validate the items that would be restored without actually performing the restore action
	.\Cat.PnP-RestoreRecycleBinData.ps1  -inputFile "Meritage-RestoreFromRecycleBin-SFL-Landdev.csv" -clientAppID "aa73d0ee-1a51-41e7-91c4-d17a65c9531c" -clientCertPath "C:\Quisitive\PowerShell\AppRegistration\mermig.pfx" -clientTenant "mymeritage" -enableTestMode $true
	
	Run restore with enableTestMode turned off
	.\Cat.PnP-RestoreRecycleBinData.ps1  -inputFile "Meritage-RestoreFromRecycleBin-SFL-Landdev.csv" -clientAppID "aa73d0ee-1a51-41e7-91c4-d17a65c9531c" -clientCertPath "C:\Quisitive\PowerShell\AppRegistration\mermig.pfx" -clientTenant "mymeritagee" -enableTestMode $false

	Table of Contents
	1.	Snap-in(s), Modules & Environment Connection(s)
	2.	Input & Output Files
	3.	Variables
	4.	Main Script
	5.  Restore Recycle Bin Data
#>

#region Script Parameters
Param(	[string] $inputFile = "{input file path relative to _Input Files}", # IE. "sample\Cat.PnP-RestoreFromRecycleBin-Client.csv"
		[string] $clientAppID = "{Client app registration ID from Azure}", # Only Supports App Reg
		[string] $clientCertPath = "{Path to client certificate file must be installed on local user}",# Only Supports App Reg
		[string] $clientTenant = "{Client tenant SharePoint prefix}", # IE. ClientTenant.sharepoint.com
		[boolean] $enableTestMode = $false # Set to $true to enable test mode which will not perform any actions and will only log the items that would have been restored
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
Start-Transcript -Path ("..\_Log FIles\{0}.PnP-RestoreRecycleBinData-{1}.log" -f $clientTenant, $scriptStart[1])
#endregion

#region 1.	Snap-in(s), Modules & Environment Connection(s)
# Load MS / PnP PowerShell Modules
Get-Module -ListAvailable | Where-Object{($_.Name -like "*SharePoint*") -or ($_.Name -like "*PnP*") } | Remove-Module
Import-Module PnP.PowerShell
#endregion

#region 2.	Input & Output Files
# File Paths
$inPath = ("..\_Input Files\{0}" -f $inputFile)
$errorPath = ("..\_Output Files\{0}.PnP-RestoreRecycleBinDataErrors-{1}.csv" -f $clientTenant, $scriptStart[1])

# Import Input File(s)
# Load CSV file
If (Test-Path ($inPath)){
	$inFile = Import-Csv $inPath -Encoding Unicode
}
else {
	stop-transcript
	throw ("The input file does not exist: {0}" -f $inPath)	
}
#endregion

#region 3.	Variables
# User Inputs
Write-Host
Write-Host "User Input Variables" -ForegroundColor Cyan
$certPass = ConvertTo-SecureString -String (Read-Host -Prompt "   Enter certificate password for the app registration") -AsPlainText -Force
$cert = Get-PnPAzureCertificate -Path $clientCertPath -Password $certPass
Remove-Variable certPass
Remove-Variable clientCertPath

# Script Variables
$tenantDomain = ("{0}.onmicrosoft.com" -f $clientTenant)

# Output Parameters & Variables
Write-Host
Write-Host "Script Parameters" -ForegroundColor Cyan
Write-Host ("   Input File: {0}" -f $inputFile)
Write-Host ("   App ID: {0}" -f $clientAppID)
Write-Host ("   Client Tenant Prefix: {0}" -f $clientTenant)
Write-Host
Write-Host "User Inputs" -ForegroundColor Cyan
Write-Host ("   Certificate Thumbprint: {0}" -f $cert.Thumbprint)
Write-Host
Write-Host "Script Variables" -ForegroundColor Cyan
Write-Host ("   Microsoft Tenant: {0}" -f $tenantDomain)
#endregion
    
#region 4.	Main Script
Write-Host
If($enableTestMode) {
	Write-Host "      Running in TEST MODE - No changes will be made, only logging the items that would have been restored" -ForegroundColor Yellow
} else {
	Write-Host "      Restoring Files from Recycle Bin..." -ForegroundColor Cyan
}

# Looping through CSV File
:jobLoop foreach ($job in $inFile) {
	#region 5.	Connecting To SP Online
	Write-Host
	Write-Host ("   {0} - Restore from Recycle Bin from {1}" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"), $job.siteTitle) -ForegroundColor Magenta
	Write-Host "      Connecting to Site:"
	Write-Host ("	  {0}" -f $job.siteURL) -ForegroundColor DarkGray
	try {
		Connect-PnPOnline $job.siteURL -tenant $tenantDomain -ClientId $clientAppID -Thumbprint $cert.Thumbprint -ErrorAction Stop
		Get-PnPSite -ErrorAction Stop
	}
	catch {
		Write-CatErrorCSV -funcCSVItem $job -funcAction "Connecting to site" -funcException $_.Exception -funcDetails $_.ErrorDetails -funcStackTrace $_.ScriptStackTrace -funcFilePath $errorPath
		Continue jobLoop	
	}
	#endregion

	#region 5.	Restore Recycle Bin Data
	try {
		
		#Get Deleted Items from recycle bin - matching given parameters
		$deletedItemsByParams = Get-PnPRecycleBinItem -RowLimit 5000000 | Where-Object {$_.DeletedByEmail -eq $job.userEmail -and $_.DeletedDate -gt [DateTime]::ParseExact($job.specificDate, "MMddyyyy", $null) }
		
		#Restore deleted items
		$count = 0
		$api = $job.siteURL + "/_api/site/RecycleBin/RestoreByIds"
		
		$deletedItemsByParams | ForEach-Object {
			Try {  
				#Call REST API   
				$body = "{""ids"":[""$($_.ID)""]}"
				$count++
				Invoke-PnPSPRestMethod -Method Post -Url $api -Content $body | Out-Null			
				$percent = ($count/($deletedItemsByParams.count))*100
				Write-Progress -Activity "Restoring Items from Recycle Bin..." -Status "Restored File $($_.LeafName) ($count of $($deletedItemsByParams.count))" -PercentComplete $percent
			}
			catch {
				write-host "Error Restoring Item $($_.LeafName) : $($_.Exception.Message)" -foregroundcolor Red
			} 
		}		
	}
	catch {
			Write-Host "      Error getting items from recycle bin" -ForegroundColor Red
			Write-CatErrorCSV -funcCSVItem $job -funcAction "Failed to Get Items From Recycle Bin" -funcException $_.Exception -funcDetails $_.ErrorDetails -funcStackTrace $_.ScriptStackTrace -funcFilePath $errorPath
	}
	#endregion

	# Disconnect from SP Online Site
	Write-Host ("      Disconnecting from SharePoint Online site: {0}" -f $job.siteTitle)
	Disconnect-PnPOnline
} 
#endregion

#region Completed Script
# Output Script Info
Complete-CatScript $scriptStart[0]

# Stop logging
Stop-Transcript
#endregion