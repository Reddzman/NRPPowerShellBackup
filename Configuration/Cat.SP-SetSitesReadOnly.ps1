<#  Cat.SP-SetSitesReadOnly.ps1

This script iterates through a CSV list of sites and webs and either configures the site quota lock on site collections or 
updates existing permissions in a the web and its lists, libraries, folders and items. In addition, on a web, a JavaScript warning message
can be added to the web to resemble the OOtB capabilities with site quota locks.

*** This script has been validated with SharePoint 2013/2016 and needs to be validated on 2010, 2019

Dependencies
	-	SharePoint PowerShell Modules (must run on SP Server)
	-	General Functions - ..\General\Cat.GeneralFunctions.psm1
	-	Input CSV File - Example: ..\_Input FIles\Testing\Cat.Tech-Action-Test.csv

Call script with parameters
	cd C:\GitHub\CatPowerShell\Configuration
	.\Cat.SP-SetSitesReadOnly.ps1 -inputFile "Samples\Configuration\Cat.SP-SetSitesReadOnly-Client.csv" -client "nrpdev2" -webAddWarningMessage $true -webWarningMessage "This site is read only at the farm administrator's request."

CSV Notes:
	siteTitle - Title of Site or Web to be configured
	siteURL - URL of Site or Web to be configured
	siteType - Options: "Site" or "Web"; not case sensitive
	siteLock - Options: "Unlock, NoAdditions, ReadOnly, or NoAccess
	webPermissionLevel - Name of permissions levels configured in site collection. IE. Read

Table of Contents
	1.	Snapin(s), Modules & Environment Connection(s)
	2.	Input & Output Files
	3.	Variables
	4.  Functions
	5.  Update-CatRoleAssignments
	6.	Main Script
	7.	Site Collection Quota
	8.	Web Permissions
#>

#region Script Parameters
Param(	[string] $inputFile = "Samples\Configuration\Cat.SP-SetSitesReadOnly-Client.csv",
		[string] $client = "nrpdev2",
		[boolean] $webAddWarning = $true,
		[string] $webWarningMessage = "This site is read only at the farm administrator's request."
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
Start-Transcript -Path ("..\_Log Files\{0}.SP-SetSitesReadOnly-{1}.log" -f $client, $scriptStart[1])
#endregion

#region 1.	Snapin(s), Modules & Environment Connection(s)
# Load MS / PnP PowerShell Modules
Get-Module -ListAvailable | Where-Object{($_.Name -like "*SharePoint*") -or ($_.Name -like "*PnP*") } | Remove-Module
Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue
#endregion

#region 2.	Input & Output Files
# File Paths
$inPath = ("..\_Input Files\{0}" -f $inputFile)
$errorPath = ("..\_Output Files\{0}.SP-SetSitesReadOnly-Errors-{1}.log" -f $client, $scriptStart[1])

# Import Input File(s)
# Load CSV file
If (Test-Path ($inPath)){
	$inFile = Import-Csv $inPath -Encoding Unicode
}
else {
	Write-Host
	throw ("The input file does not exist: {0}" -f $inPath)	
}
#endregion

#region 3.	Variables
# Script Variables
$webWarningActionName = "ReadOnlyMessage"
$webWarningStartingSequence = 1100
$webWarningLocation = "ScriptLink"
$webWarningBackground = "#f8d4d4"
$webWarningJavaScript = @"
function Execute_$($webWarningActionName)_MessageBar_$('Web')()
{
	var st = document.getElementById('$webWarningActionName');
	if(st != null)
	{
		st.style.display = '';
		return;
	}
	st = document.createElement("div");
	st.id = "$($webWarningActionName)_MessageBar_$('Web')";
	st.innerHTML = "<div style='background: #f8d4d4; padding: 12px 14px 6px; border: 1px solid; font-family: \"Segoe UI\", Tahoma, sans-serif; font-size: 13px; min-height: 24px;'>$webWarningMessage</div>";
	document.body.insertBefore(st, document.body.childNodes[0]);
}
// don't show on modal dialog windows
if(!window.location.search.match("[?&]IsDlg=1"))
{
	Execute_$($webWarningActionName)_MessageBar_$('Web')();
}
"@

# Output Parameters & Variables
Write-Host
Write-Host "Script Parameters" -ForegroundColor Cyan
Write-Host ("   Input File: {0}" -f $inputFile)
Write-Host ("   Client: {0}" -f $client)
if ($webAddWarning) {
	Write-Host
	Write-Host "Script Variables" -ForegroundColor Cyan
	Write-Host ("   Warning Action Name: {0}" -f $webWarningActionName)
	Write-Host ("   Starting Sequence: {0}" -f $webWarningStartingSequence)
	Write-Host
	Write-Host "Warning JavaScript" -ForegroundColor Cyan
	Write-Host $webWarningJavaScript
}
#endregion

#region 4.  Functions

	#region 5.  Update-CatRoleAssignments

	<# This function updates the role assignments passed to the function be removing existing role definitions and adding a
	new role definition based on the provided permission level.

	Sample Call:	Update-CatRoleAssignments -functionAssignments $obj.RoleAssignments -functionWeb $web -functionPermissionLevel $csv.webPermissionLevel #>

	function Update-CatRoleAssignments { 
		param(
			[psobject] $functionAssignments,
			[psobject] $functionWeb,
			[string] $functionPermissionLevel
		)
	
		# looping provided role assignments
		foreach ($functionAssignment in $functionAssignments) {
			# retrieving role definitions for each role assignment
			$fuctionDefinitions = $functionAssignment.RoleDefinitionBindings | Where-Object {$_.Name -ne "Limited Access"}
			
			# removing existing role definitions
			foreach ($functionDefinition in $fuctionDefinitions) {
				$functionAssignment.RoleDefinitionBindings.Remove($functionDefinition)
				$functionAssignment.Update()
			}

			# Add the desired role definition
			if ($fuctionDefinitions.Count -gt 0) {
				$functionNewRole = $functionWeb.RoleDefinitions | Where-Object {$_.Name -eq $functionPermissionLevel}
				$functionAssignment.RoleDefinitionBindings.Add($functionNewRole)
				$functionAssignment.Update()
			}
		}
	}
	#endregion
	
#endregion

#region 6.	Main Script
Write-Host
Write-Host ("Configuring Sites Read Only: {0}" -f $inputFile) -ForegroundColor Cyan

# Looping through input file
:siteLoop ForEach ($site in $inFile){
	Write-Host
    Write-Host ("{0} - Site: {1}" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"), $site.siteTitle) -NoNewline

	# Determine if updating permissions on a web or setting a site quota lock
	switch ($site.siteType.ToLower()) {
		#region 7.	Site Collection Quota
		site {
			try {
				Set-SPSite -Identity $site.siteURL -LockState $site.siteLock -ErrorAction Stop
				Write-Host (" - Site Quota Lock Configured: {0}" -f $site.siteLock) -ForegroundColor Green
			}
			catch {Write-CatCSVErrorInfo -funcError $_ -funcObj $site -funcOutputPath $errorPath}
		}
		#endregion

		#region 8.	Web Permissions
		web {
			try {
				Write-Host " - Updating Web, List, Item, and Folder permissions"

				# Get web object / update perms
				Write-host ("   {0} - Retreiving Web" -f (Get-Date -Format "yy-MM-dd HH:mm:ss")) -ForegroundColor DarkGray
				$web = Get-SPWeb -Identity $site.siteURL -ErrorAction Stop

				# Breaking Inhertance on web
				if (!$web.HasUniquePerm) {
					Write-host ("   {0} - Breaking Inheritance" -f (Get-Date -Format "yy-MM-dd HH:mm:ss")) -ForegroundColor DarkGray
					$web.BreakRoleInheritance($true)
				}

				# Configure warning on web
				if ($webAddWarning) {
					Write-host ("   {0} - Configuring Warning Message" -f (Get-Date -Format "yy-MM-dd HH:mm:ss")) -ForegroundColor DarkGray

					# Removing existing custom actions
					$existingActions = $web.UserCustomActions | ? { $_.Title -eq $webWarningActionName }
					$existingActions
					foreach ($action in $existingactions) {
						$web.UserCustomActions.Item($action.Id).Delete()
					}

					# Adding new custom action
					Write-Host
					$webWarningMessage = $webWarningMessage.Replace( "`"", "'");
					$customAction = $null
					$customAction = $web.UserCustomActions.Add()
					$customAction.Location    = $webWarningLocation
					$customAction.Sequence    = $webWarningStartingSequence
					$customAction.Title       = $webWarningActionName
					$customAction.ScriptBlock = $webWarningJavaScript
					$customAction.Update()
				}
				
				# Update web permissions
				Write-host ("   {0} - Updating web permissions" -f (Get-Date -Format "yy-MM-dd HH:mm:ss")) -ForegroundColor DarkGray
				Update-CatRoleAssignments -functionAssignments $web.RoleAssignments -functionWeb $web -functionPermissionLevel $site.webPermissionLevel

				# Process each list/library
				$lists = $web.lists | where-object {$_.Hidden -eq $false}
				foreach ($list in $lists) {
					# Checking list for unique permissions
					if ($list.HasUniqueRoleAssignments -eq $true){
						Write-Host ("   {0} - Updating list permissions: {1}" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"), $list.Title) -ForegroundColor DarkGray
						Update-CatRoleAssignments -functionAssignments $list.RoleAssignments -functionWeb $web -functionPermissionLevel $site.webPermissionLevel
					}

					# Retrieve items with unique permissions and update permissions
					$items = $list.Items | Where-Object {$_.HasUniqueRoleAssignments -eq $true}
					foreach ($item in $items) {
						Write-Host ("   {0} - Updating item permissions: {1} - {2}" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"), $list.Title, $item.Id) -ForegroundColor DarkGray
						Update-CatRoleAssignments -functionAssignments $item.RoleAssignments -functionWeb $web -functionPermissionLevel $site.webPermissionLevel
					}

					# Retrieve folders with unique permissions and update permissions
					$folders = $list.Folders | Where-Object {$_.HasUniqueRoleAssignments -eq $true}
					foreach ($folder in $folders) {
						Write-Host ("   {0} - Updating folder permissions: {1} - {2}" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"), $list.Title, $folder.name) -ForegroundColor DarkGray
						Update-CatRoleAssignments -functionAssignments $folder.RoleAssignments -functionWeb $web -functionPermissionLevel $site.webPermissionLevel
					}
				}
				
				Write-Host "   Permissions updated" -ForegroundColor Green
			}
			catch {Write-CatCSVErrorInfo -funcError $_ -funcObj $site -funcOutputPath $errorPath}
		}
		#endregion

		Default {
			Write-Host " - Invalid Site Type" -ForegroundColor Red
			$errorRecord = $site
			Add-Member -InputObject $errorRecord -MemberType NoteProperty -Name "ErrorMessage" -Value "Invalid Site Type"
			$errorRecord | Export-Csv -Path $errorPath -NoTypeInformation -Append
			Remove-Variable errorRecord
		}
	}
}
#endregion

#region Completed Script
# Output Script Info
Complete-CatScript $scriptStart[0]

# Stop logging
Stop-Transcript
#endregion