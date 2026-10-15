<#  Cat.SkySync-CreateSPOConnections.ps1

This script iterate through a CSV file to create connections and a conection pool for new SPO Site collections for the use in SkySync
migration jobs.

Dependencies
	-	PnP PowerShell https://docs.microsoft.com/en-us/powershell/sharepoint/sharepoint-pnp/sharepoint-pnp-cmdlets?view=sharepoint-ps
	-	SharePoint Online Module
	-	SkySync CLI
	-	\General\Cat.PnP-GeneralFunctions.psm1
	-	Input CSV File as a Script Parameter located in _Input Files - "Cat.PerformAction.Test.csv"

Call Script with Input File
	.\Cat.SkySync-CreateSPOConnections.ps1 -csvFile client.ITSites-Test.csv

CSV Example
	siteName,siteURL,deptName
	Infrastructure Teams,https://client.sharepoint.com/sites/ITLTBoxMigration,Internal Infrastructure

Table of Contents
	1.	Input & Output Files
	2.	Variables
	3.	Snapin(s), Modules & Environment Connection(s)
	4.	Main Script
	5.	Update Accounts JSON
	6.	Create Connections
	7.	Create Conneciton Pool
	8.	Create User Map
	9.	Create Group Map
#>

#region Script Parameters
Param(  
	[String] $connectionsFile = "Cat.SkySync-SPOConnections-client-Test.csv",
	[String] $accountsFile = "Cat.SkySync-Accounts-client.json",
	[String] $amFile = "Cat.SkySync-AccountMapTemp-client.json",
	[String] $gmFile = "Cat.SkySync-GroupMapTemp-client.json"
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
Start-Transcript -Path ("..\_Log FIles\CreateSPOConnnections-{0}.log" -f $scriptStart[1])
#endregion

#region 1.	Input & Output Files
# File Paths
$conPath = ("..\_Input Files\{0}" -f $connectionsFile)
$acctPath = ("..\_Input Files\{0}" -f $accountsFile)
$amPath = ("..\_Input Files\{0}" -f $amFile)
$gmPath = ("..\_Input Files\{0}" -f $gmFile)

# Import Input File(s)
# Load connections CSV
If (Test-Path ($conPath)){
	$conInput = Import-Csv $conPath -Encoding Unicode
}
else {
	Write-Host
	throw ("The connections file does not exist: {0}" -f $conPath)	
}

# Load accounts JSON
If (Test-Path ($acctPath)){
	$acctInput = Get-Content -Path $acctPath -Raw | ConvertFrom-Json
}
else {
	Write-Host
	throw ("The accounts file does not exist: {0}" -f $acctPath)	
}

# Load Account Map Template
If (Test-Path ($amPath)){
	$amJSON = Get-Content -Path $amPath -Raw
}
else {
	Write-Host
	throw ("The account map file does not exist: {0}" -f $amPath)	
}

# Load Group Map Template
If (Test-Path ($gmPath)){
	$gmJSON = Get-Content -Path $gmPath -Raw
}
else {
	Write-Host
	throw ("The account map file does not exist: {0}" -f $gmPath)	
}
#endregion

#region 2.	Variables
# User Inputs
Write-Host
Write-Host "User Input Variables" -ForegroundColor Cyan
Write-Host "   Enter credentials for SharePoint:"
$spCred = Get-Credential -Message "Enter your credentials for SharePoint:"

# Getting Passwords
foreach	($acct in $acctInput){
	$acct.password = Read-Host -Prompt ("   Enter the password for: {0}" -f $acct.username)
} # foreach	($acct in $acctInput)

# Script Variables
$adminURL = "https://client-admin.sharepoint.com"
$serverURL = "http://localhost:9090/"
$baseURL = ("{0}v1" -f $serverURL)
$contentType = "application/json"

# Output Parameters & Variables
Write-Host
Write-Host "Script Parameters" -ForegroundColor Cyan
Write-Host ("   Connections File: {0}" -f $connectionsFile)
Write-Host ("   Accounts File: {0}" -f $accountsFile)
Write-Host
Write-Host "Script Variables" -ForegroundColor Cyan
Write-Host ("   Admin URL: {0}" -f $adminURL)
Write-Host ("   SPO Credentials: {0}" -f $spCred.UserName)
Write-Host ("   SkySync Credentials: {0}" -f $ssUser)
Write-Host ("   SkySync URL: {0}" -f $serverURL)
Write-Host ("   Base URL: {0}" -f $baseURL)
Write-Host ("   Content Type: {0}" -f $contentType)
<#Write-Host
Write-Host "User Account Map JSON" -ForegroundColor Cyan
Write-Host $amJSON
Write-Host
Write-Host "Group Account Map JSON" -ForegroundColor Cyan
Write-Host $gmJSON #>
#endregion

#region 3.	Snapin(s), Modules & Environment Connection(s)
# Load Custom Modules
Import-Module Microsoft.Online.SharePoint.PowerShell -DisableNameChecking
Import-Module "..\SkySync\Cat.SkySyncFunctions.psm1"
#endregion

#region 4.	Main Script
# Loop through connections
foreach ($con in $conInput) {
	Write-Host
	Write-Host ("Creating Connections: {0}" -f $con.SiteName) -ForegroundColor Magenta

	# Refreshing Auth Token
	$accessToken = skysync-cli sessions token --output-json false
	$authHeader = @{"Authorization"="Bearer $accessToken";}
	
	#region 5.	Update Accounts JSON
	# Loop through JSON Accounts
	Write-Host
	Write-Host "   Updating unique connection information" -ForegroundColor DarkGray
	foreach	($acct in $acctInput){
		$acct.name = ("SPO-{0}-{1}-{2}-{3}" -f 
			$con.BusinessUnit.Replace(" ","").Replace(" ","").Replace("(","").Replace(")","").Replace("&","").Replace("-",""),
			$con.Department.Replace(" ","").Replace(" ","").Replace("(","").Replace(")","").Replace("&","").Replace("-",""), 
			$con.SiteName.Replace(" ","").Replace(" ","").Replace("(","").Replace(")","").Replace("&","").Replace("-",""),
			$acct.username.replace("o365-skysync-","").replace("@client.com","")
		)
		Write-Host ("      {0}" -f $acct.name)
		$acct.uri = $con.siteURL
	} # foreach	($acct in $acctInput)
	#endregion

	#region 6.	Create Connections
	# Loop through JSON Accounts
	Write-Host
	Write-Host "   Creating SkySync Connections" -ForegroundColor DarkGray
	$connections = @()
	Connect-SPOService -Url $adminURL -Credential $spCred
	foreach	($acct in $acctInput){
		Write-Host ("      {0}" -f $acct.name)
		if (connectionExists -findName $acct.name){
			Write-Host ("         Connection {0} already exists" -f $acct.name) -ForegroundColor Yellow
		}
		else{
			# Add as Site Collection Administrator
			Write-Host "         Adding as Site Collection Admin"
			$tempUser = Set-SPOUser -Site $con.siteURL -LoginName $acct.username -IsSiteCollectionAdmin $true

			# Create Connections
			Write-Host "         Creating Connection"
			$platform = $acct.platform
			$name = $acct.name
			$connections += $acct | ConvertTo-Json -Depth 10 | skysync-cli connections add --platform=$platform --name=$name --in --json | ConvertFrom-Json
		}
	} # foreach	($acct in $acctInput)
	#endregion

	#region 7.	Create Conneciton Pool
	$poolName = ("SPO-{0}-{1}-{2}-Pool" -f 
		$con.BusinessUnit.Replace(" ","").Replace(" ","").Replace(" ","").Replace("(","").Replace(")","").Replace("&","").Replace("-",""),
		$con.Department.Replace(" ","").Replace(" ","").Replace(" ","").Replace("(","").Replace(")","").Replace("&","").Replace("-",""),
		$con.SiteName.Replace(" ","").Replace(" ","").Replace(" ","").Replace("(","").Replace(")","").Replace("&","").Replace("-","")
	)
	Write-Host
	Write-Host ("   Creating SkySync Connection Pool: {0}" -f $poolName) -ForegroundColor DarkGray

	# loop through created Connections
	foreach	($connection in $connections){
		Write-Host ("      Adding Connection: {0}" -f $connection.name)
		$result = skysync-cli connections pool assign $connection.id --name $poolName  
	} # foreach	($connection in $connections)
	#endregion

	#region 8.	Create User Map
	Write-Host
	Write-Host ("   Creating Account Map: {0}" -f $poolName) -ForegroundColor DarkGray
	$amMapJSON = $amJSON.Replace('{account_map_name}',$poolName).Replace('{destination_Connection_id}',(GetConnectionID $poolName))
	$amEndpoint = ("{0}/account_maps" -f $baseURL)

	# REST Web Request
	$amResponse = Invoke-RestMethod $amEndpoint -Method POST -Body $amMapJSON -Headers $authHeader -ContentType $contentType
	#endregion

	#region 9.	Create Group Map
	Write-Host
	Write-Host ("   Creating Group Map: {0}" -f $poolName) -ForegroundColor DarkGray
	$grpMapJSON = $gmJSON.Replace('{group_map_name}',$poolName).Replace('{destination_Connection_id}',(GetConnectionID $poolName))
	$grpEndpoint = ("{0}/group_maps" -f $baseURL)

	# REST Web Request
	$grpResponse = Invoke-RestMethod $grpEndpoint -Method POST -Body $grpMapJSON -Headers $authHeader -ContentType $contentType
	#endregion
} # foreach ($con in $conInput)
#endregion

#region Completed Screen Output
Complete-CatScript $scriptStart[0]
Disconnect-SPOService
Stop-Transcript
#endregion