<#	Cat.CM-NewUserMappingOnPrem.ps1

This script iterates through Azure AD users and groups to generate a mapping file for content Matrix. 
AD Sync should already be in place so users exist in the cloud.

Dependencies
	-	AzureAD

Enhancements
	-	

Table of Contents
	1.	Input & Output Files
	2.	Variables
	3.	Snapin(s), Modules & Environment Connection(s)
	4.	Main Script
#>

#region Script Parameters
Param(	[String] $srcUserClaim = "i:0#.f|membership|",
		[String] $dstUserClaim = "i:0#.f|membership|" 
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
Start-Transcript -Path ("..\_Log FIles\NewUserMappingAAD-{0}.log" -f $scriptStart[1])
#endregion

#region 1.	Input & Output Files
# File Paths
$outPath = ("..\_Output Files\MetalogixUserMapping-{0}.xml" -f $scriptStart[1])

# Create Output File(s)
Out-File -FilePath $outPath -InputObject "<Mappings>"
#endregion

#region 2.	Variables
# User Inputs
Write-Host
Write-Host "User Input Variables" -ForegroundColor DarkCyan
Write-Host "   Enter credentials for Azure Active Directory:"
$aadCred = Get-Credential -Message "Enter your credentials for Azure Active Directory:"

# Script Variables
$userCount = 0
$groupCount = 0

# Output Parameters & Variables
Write-Host
Write-Host "Script Parameters" -ForegroundColor Cyan
Write-Host ("   Source User Claim: {0}" -f $srcUserClaim)
Write-Host ("   Destination User Claim: {0}" -f $dstUserClaim)
Write-Host
Write-Host "User Inputs" -ForegroundColor Cyan
Write-Host ("   AAD Credentials: {0}" -f $aadCred.UserName)
Write-Host
Write-Host "Script Variables" -ForegroundColor Cyan
Write-Host ("   User Count: {0}" -f $userCount)
#endregion

#region 3.	Snapin(s), Modules & Environment Connection(s)
# Load PowerShell Snap-In(s)
Import-Module AzureAD
#endregion
    
#region 4.	Main Script
# Screen Output
Write-Host
Write-Host "Creating User Mappings File" -ForegroundColor DarkCyan

# Connecting to Azure AD
Write-Host "   Connecting to Azure AD"
$aadConnection = Connect-AzureAD -Credential $aadCred

# Retrieving All Users
Write-Host "   Retrieving All Users"
$users = Get-AzureADUser -All:$true

# Looping through all User Accounts
Write-Host
Write-Host "   Looping through Users Accounts" -ForegroundColor Magenta
foreach ($user in $users)
{
    # Increment User Count
    $userCount++         

    If ($user.UserPrincipalName -ne $null)
    {
        $srcUser = $srcUserClaim + $user.UserPrincipalName
        $dstUser = $dstUserClaim + $user.UserPrincipalName
        $userXML = ("<Mapping Source='{0}' Target='{1}' />" -f $srcUser,$dstUser)
        out-file -filepath $outPath -append -inputobject $userXML
        Write-Host ("      {0}" -f $userXML)
    }
    else 
    {
        Write-Host ("      User Principal Name not populated: {0} - {1}" -f $user.ObjectId, $user.DisplayName) -ForegroundColor DarkRed
    }
}
out-file -filepath $outPath -append -inputobject "</Mappings>"
Write-Host
Write-Host ("   User Count: {0}" -f $userCount)
Write-Host "Mapping File Complete" -ForegroundColor Green
#endregion

#region Completed Script
# Disconnect from SharePoint
Disconnect-AzureAD

# Output Script Info
Complete-CatScript $scriptStart[0]

# Stop logging
Stop-Transcript
#endregion