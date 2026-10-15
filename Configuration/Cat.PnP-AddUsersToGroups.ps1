#region Documentation
<#  Cat.PnP-AddUsersToGroups.ps1

This script iterates a CSV file to configure group membership for both M265 groups and SharePoint Groups. This allows go live
permissions to be configured quickly based on CSV data as part of migrations or other efforts.

Dependencies
	-	PnP PowerShell https://pnp.github.io/powershell/
	-	General Functions - ..\General\Cat.GeneralFunctions.psm1
	-	Input CSV File in ..\_Input Files\ - Example: Samples\Configuration\Cat.PnP-AddUsersToGroups-Client.csv

Call script with parameters
	.\Cat.PnP-AddUsersToGroups.ps1 -inputFile "Samples\Configuration\Cat.PnP-AddUsersToGroups-Client.csv" -appID "caf6b11a-cf03-416c-b0f0-67a09a8e30ee" -certPath "C:\GitHub\NRPDev\AppRegistration\PowerShell App Registration.pfx" -clientTenant "nrpdev2"

Table of Contents
	1.	Snapin(s), Modules & Environment Connection(s)
	2.	Input & Output Files
	3.	Variables
	4.	Main Script
	5.	Adding Membership to SharePoint Groups
	6.	Adding Membership to M365 Groups
#>
#endregion

#region Script Parameters - Values are samples, remove the brackets for the actual values
Param(	[string] $inputFile = "Cat.PnP-ConfigMultipleLists-Client.csv",
		[string] $appID = "[ClientAppID]",
		[string] $certPath = "C:\GitHub\client\AppRegistration\ClientCertificate.pfx",
		[string] $clientTenant = "[ClientTenatPrefix]"
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
Start-Transcript -Path ("..\_Log Files\{0}.PnP-ConfigGroupMembership-{1}.log" -f $clientTenant, $scriptStart[1])
#endregion

#region 1.	Snapin(s), Modules & Environment Connection(s)
# Load MS / PnP PowerShell Modules
Get-Module -ListAvailable | Where-Object{($_.Name -like "*SharePoint*") -or ($_.Name -like "*PnP*") } | Remove-Module
Import-Module PnP.PowerShell

# Load Custom Modules
Import-Module "..\Reporting\Cat.ReportingFunctions.psm1"
#endregion

#region 2.	Input & Output Files
# File Paths
$inPath = ("..\_Input Files\{0}" -f $inputFile)

# Import Input File(s)
# Load CSV file
If (Test-Path ($inPath)){
	$inFile = Import-Csv $inPath -Encoding Unicode
}
else {
	Write-Host
	Stop-Transcript
	throw ("The input file does not exist: {0}" -f $inPath)	
}
#endregion

#region 3.	Variables
# User Inputs
Write-Host
Write-Host "User Input Variables" -ForegroundColor Cyan
$certPass = ConvertTo-SecureString -String (Read-Host -Prompt "   Enter certificate password for the app registration") -AsPlainText -Force

# Script Variables
$tenant = ("{0}.onmicrosoft.com" -f $clientTenant)
$tenantRootSite = ("https://{0}.sharepoint.com" -f $clientTenant)
$m365GroupClaim = "c:0o.c|federateddirectoryclaimprovider|"

# Output Parameters & Variables
Write-Host
Write-Host "Script Parameters" -ForegroundColor Cyan
Write-Host ("   Input File: {0}" -f $inputFile)
Write-Host ("   App ID: {0}" -f $appID)
Write-Host ("   Certificate Path: {0}" -f $certPath)
Write-Host
Write-Host "User Inputs" -ForegroundColor Cyan
Write-Host ("   Certificate Password: {0}" -f $certPass)
Write-Host
Write-Host "Script Variables" -ForegroundColor Cyan
Write-Host ("   Microsoft Tenant: {0}" -f $tenant)
Write-Host ("   Site URL: {0}" -f $tenantRootSite)
Write-Host ("   M365 Group Claim: {0}" -f $m365GroupClaim)
#endregion
    
#region 4.	Main Script
# Connecting To SP Online
Write-Host
Write-Host ("{0} - Updating Group Permisssions: {1}" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"), $inputFile) -ForegroundColor Magenta

# Looping CSV Records
:acctLoop ForEach ($acct in $inFile){
	# Connecting to the SharePoint site
	Write-Host
	Write-Host ("   {0} - Adding Account, {1}, to Group, {2}, on {3}" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"),$acct.loginName, $acct.groupName, $acct.siteUrl.Replace($tenantRootSite,""))
	Write-Host "      Connecting to SharePoint site" -ForegroundColor DarkGray
	<#$spSiteConnection = #> Connect-PnPOnline -ClientId $appID -Url $acct.siteUrl -Tenant $tenant -CertificatePath $certPath -CertificatePassword $certPass -WarningAction SilentlyContinue #-ReturnConnection
			
	switch ($acct.groupType) {
		sharepoint {
			#region 5.	Adding Membership to SharePoint Groups
			# Determing loginName
			switch ($acct.loginType) {
				m365group {
					Write-Host "      Retrieving M365 Group" -ForegroundColor DarkGray
					$group = Get-PnPMicrosoft365Group -Identity $acct.loginName
					$loginName = ("{0}{1}" -f $m365GroupClaim, $group.GroupId)
					Remove-Variable group
				}
				user {
					# ensuring the user exists in the SharePoint all users list
					Write-Host "      Adding User to all users list" -ForegroundColor DarkGray
					$loginName = $acct.loginName
					$tempUser = New-PnPUser -LoginName $loginName #-Connection $spSiteConnection
					Remove-Variable tempUser
				}
				Default {$loginName = $acct.loginName}
			}

			# Adding User or group to SharePoint permissions
			Write-Host "      Adding to SharePoint Group" -ForegroundColor DarkGray
			Add-PnPGroupMember -LoginName $loginName -Group $acct.groupName #-Connection $spSiteConnection

			# Disconnecting from Site & Clearing Variables
			Remove-Variable loginName
			#endregion
		}
		m365group {
			#region 6.	Adding Membership to M365 Groups
			switch ($acct.groupRole) {
				owner {
					Write-Host "      Adding as Owner to M365 Group" -ForegroundColor DarkGray
					Add-PnPMicrosoft365GroupOwner -Identity $acct.groupName -Users $acct.loginName -Connection $spSiteConnection
				}
				member {
					Write-Host "      Adding as Member to M365 Group" -ForegroundColor DarkGray
					Add-PnPMicrosoft365GroupMember -Identity $acct.groupName -Users $acct.loginName -Connection $spSiteConnection
				}
				Default {
					Write-Host ("      Invalid M365 Group Role" -f (Get-Date -Format "yy-MM-dd HH:mm:ss")) -ForegroundColor Yellow
					continue acctLoop
				}
			}
			#endregion
		}
		Default {
			Write-Host ("      Invalid Group Type" -f (Get-Date -Format "yy-MM-dd HH:mm:ss")) -ForegroundColor Yellow
			continue acctLoop
		}
	}

	# Disconnecting from Site & Clearing Variables
	Disconnect-PnPOnline #-Connection $spSiteConnection
	#Remove-Variable spSiteConnection
}
#endregion

#region Completed Script
# Output Script Info
Complete-CatScript $scriptStart[0]

# Stop logging
Stop-Transcript
#endregion