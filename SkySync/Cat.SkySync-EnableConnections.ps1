<#  Cat.SkySync-EnableConnections.ps1

This script leverages the SkySync API to enable connections that are disabled.

Dependencies 
	-	SkySync Rest API

Call script with parameters
	.\Cat.SkySync-EnableConnections.ps1 -maxPerRequest 100

Table of Contents
	1.	Variables
    2.	Main Script
    3.  Get All Connections
    4.  Check for Disabled Connections
    5.  Enabling Connections
#>

#region Script Parameters
Param(	[String] $maxPerRequest = 1000)
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
Start-Transcript -Path ("..\Enable Connections Logs\EnableConnections-{0}.log" -f $scriptStart[1])
#endregion

#region 1.	Variables
# Script Variables
$serverURL = "http://localhost:9090/"
$baseURL = ("{0}v1" -f $serverURL)
$contentType = 'application/json'
$offset = 0
$allConnections = @()
$disabledConnections = @()
$accessToken = skysync-cli sessions token --output-json false
$authHeader = @{"Authorization"="Bearer $accessToken";}

# Output Parameters & Variables
Write-Host
Write-Host "Script Parameters" -ForegroundColor Cyan
Write-Host ("   Max Items per Request: {0}" -f $maxPerRequest)
Write-Host
Write-Host "Script Variables" -ForegroundColor Cyan
Write-Host ("   Server URL: {0}" -f $serverURL)
Write-Host ("   REST Base URL: {0}" -f $baseURL)
Write-Host ("   Content Type: {0}" -f $contentType)
Write-Host ("   Offset: {0}" -f $offset)
#endregion
    
#region 2.	Main Script
Write-Host
Write-host "Enabling Connections" -ForegroundColor Magenta

try {
    #region 3.  Get All Connections
    Write-Host "   Getting Disabled Connections" -ForegroundColor DarkGray
    do {
        Write-Host ("      Getting Connections: Offset = {0}; Limit = {1}" -f $offset, $maxPerRequest)
        $conMethod = ("{0}/connections?pools=0&fields=all&offset={1}&limit={2}" -f $baseURL, $offset, $maxPerRequest)
        $conResponse = Invoke-RestMethod $conMethod -Method GET -Headers $authHeader -ContentType $contentType
        $offset += $maxPerRequest
        $allConnections += $conResponse.connections
    } while($conResponse.meta.has_more)
    Write-Host ("      Connections Found: {0}" -f $allConnections.Count)
    #endregion

    #region 4.  Check for Disabled Connections
    Write-Host "      Checking Connections"    
    foreach ($conn in $allConnections) {
        if ($conn.disabled) {
            $disabledConnections += $conn
        }
    }
    Write-Host ("      Disabled Connections Found: {0}" -f $disabledConnections.Count)
    #endregion
    
    #region 5.  Enabling Connections
    Write-Host
    If ($disabledConnections -gt 0){
        Write-Host ("   Enabling connections: {0}" -f $disabledConnections.Count) -ForegroundColor DarkGray
        foreach ($conn in $disabledConnections) {
            Write-Host ("      Connection: {0}" -f $conn.Name)
            $endPoint = "$baseUrl/connections/" + $disabledConn.id + "?enable=1"
            $disabledResponse = Invoke-RestMethod $endPoint -Method PATCH -Headers $authHeader -ContentType 'application/json'
        }
        Write-Host
        Write-Host "   All connections where successfully enabled." -ForegroundColor Green
    }
    Else {
        Write-Host "   All Connections are currently enabled." -ForegroundColor Green
    }    
    #endregion
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