<#  Cat.SkySync-IgnoreFlaggedItems.ps1

This script leverages the SkySync API to ignore flagged items with the proper wave,
category, and message.

Dependencies 
	-	SkySync Rest API

Call script with parameters
	.\Cat.Tech-ScriptName.ps1 -InputFile InputFileName.csv

Table of Contents
	1.	Input & Output Files
	2.	Variables
	3.	Snapin(s), Modules & Environment Connection(s)
	4.	Main Script
#>

#region Script Parameters
Param(	
    [String] $searchCriteria = "TestWave-TestBU-TestDept",
    [String] $failCategory = "Item Not Found Failures",
    [String] $failMessage = "Cannot access a disposed object.",
    [String] $maxItemsPerRequest = 1000
)
#endregion

#region Start Date / Time
# Set Powershell run location
Set-Location "C:\Scripts\SkySync"

# Remove Existing Catapult Modules to ensure PS session has most up to date functions
Get-Module Cat.* | Remove-Module

# Load General Functions
Import-Module "..\General\Cat.GeneralFunctions.psm1"

# Get Start Date / Time
$scriptStart = Start-CatScript

# Start Transcript
Start-Transcript -Path ("..\Flagged Items Logs\IgnoreFlaggedItems-{0}.log" -f $scriptStart[1])
#endregion

#region 1.	Variables
# Script Variables
$serverURL = "http://localhost:9090/"
$baseURL = ("{0}v1" -f $serverURL)
$contentType = 'application/json'
$offset = 0
$totalItems = @()
$accessToken = skysync-cli sessions token --output-json false
$authHeader = @{"Authorization"="Bearer $accessToken";}

# Output Parameters & Variables
Write-Host
Write-Host "Script Parameters" -ForegroundColor Cyan
Write-Host ("   Max Items per Request: {0}" -f $maxItemsPerRequest)
Write-Host
Write-Host "Script Variables" -ForegroundColor Cyan
Write-Host ("   Server URL: {0}" -f $serverURL)
Write-Host ("   REST Base URL: {0}" -f $baseURL)
Write-Host ("   Content Type: {0}" -f $contentType)
Write-Host ("   Offset: {0}" -f $offset)
#endregion
    
#region 2.	Main Script
Write-Host
Write-host "Ignoring Flagged Items: {0}" -ForegroundColor Magenta

try {
    #region 3.  Get Flagged items
    Write-Host "   Getting Flagged Items" -ForegroundColor DarkGray
	$itemsRequestMethod = "$baseUrl/transfers/items/flagged?fields=all&limit=$maxItemsPerRequest&offset=$offset"
    $itemsResponse = Invoke-RestMethod $itemsRequestMethod -Headers $authHeader

    # Getting Only items with specific errors
    $totalItems += $itemsResponse.item | Where{($_.transfer.name -like "*$searchCriteria*") -and ($_.audit_category.name -eq $failCategory) -and ($_.last_failure.message -eq $failMessage)}

    # Check to see if more items exist
	if($itemsResponse.meta.has_more) #make additional job requests as necessary
	{
		do
		{
			$offset += $maxItemsPerRequest
			$itemsRequestMethod = "$baseUrl/transfers/items/flagged?fields=all&limit=$maxItemsPerRequest&offset=$offset"
            $itemsResponse = Invoke-RestMethod $itemsRequestMethod -Headers $authHeader #-Proxy 'http://localhost:8888
			$totalItems += $itemsResponse.item | Where{($_.transfer.name -like "*$searchCriteria*") -and ($_.audit_category.name -eq $failCategory) -and ($_.last_failure.message -eq $failMessage)}
		}
		while($itemsResponse.meta.has_more)
	}
    Write-Host ("      Identified {0} items to ignore" -f $totalItems.count)
    $decision = Read-Host "         Do you want to continue (Y/N)"
    #endregion

    If ($decision -eq "Y"){
        Write-Host "   Setting items to be ignored" -ForegroundColor DarkGray
        foreach ($item in $totalItems)
        {
            Write-Host ("      Configuring Job: {0}" -f $item.id)
            $retryItemMethod = "$baseUrl/transfers/items/$($item.id)?ignore=1"
            $retryItemResponse = Invoke-RestMethod $retryItemMethod -Headers $authHeader -Method Patch #-Proxy 'http://localhost:8888'
        }
        
        Write-Host
        Write-Host "All items Successfully marked for ignore." -ForegroundColor Green
    }
}
catch{
    $exMsg = $_.Exception.Message
    $line = $_.Exception.InvocationInfo.ScriptLineNumber
    $st = $_.ScriptStackTrace
    Write-Host "An error occurred while marking items for retry: ${exMsg}. Line ${line}. ${st}"
    return
}
#endregion

#region Completed Screen Output
Complete-CatScript $scriptStart[0]
Stop-Transcript
#endregion