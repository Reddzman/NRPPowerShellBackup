<#	Cat.CM-NewUserMappingOnPrem.ps1

This script iterates through MSOnline users and compares them to AD to generate a mapping file for users. When
migrating from on premises to online.  AD Sync should already be in place so users exist in the cloud.

Dependencies
	-	SharePoint Online Module
	-	MS Online Module
	-	Active Directory Module

Table of Contents
	1.	Input & Output Files
	2.	Variables
	3.	Snapin(s), Modules & Environment Connection(s)
	4.	Main Script
#>

#region Start Date / Time
# Set PS location
Set-Location $PSScriptRoot

# Get Script Parent Location
$LocSplit = $PSScriptRoot.Split("\")
$LocReplace = $LocSplit[($LocSplit.Count - 1)]
$Location = $PSScriptRoot.Replace($LocReplace,"")

# Remove Existing Catapult Modules to ensure PS session has most up to date functions
Get-Module Cat.* | Remove-Module

# Load General Functions
Import-Module ("{0}General\Cat.GeneralFunctions.psm1" -f $Location)

# Get Start Date / Time
$Start = Start-CatScript
#endregion

#region 1.	Input & Output Files
# File Paths
$OutPath = ("{0}_Output Files\MetalogixUserMapping-{1}.xml" -f $Location, $Start[1])

# Create Output File(s)
Out-File -FilePath $OutPath -InputObject "<Mappings>"
#endregion

#region 2.	Variables
# User Inputs
Write-Host
Write-Host "User Input Variables" -ForegroundColor DarkCyan
Write-Host "   Enter credentials for O365 Administration:"
$O365Cred = Get-Credential -Message "Enter your credentials for O365 Administration:"
Write-Host "   Enter credentials for Active Directory:"
$ADCred = Get-Credential -Message "Enter your credentials for Active Directory:" 

# Script Variables
$SP13Prefix = "i:0#.w|molycorp\" # On Premise Claim + domain\
$SPOPrefix = "i:0#.f|membership|" # Online Claim
$UserDomain = "neomaterials.com" # UPN Domain for O365 filter & internal domain name for quering AD
$UserCount = 0
#endregion

#region 3.	Snapin(s), Modules & Environment Connection(s)
# Load PowerShell Snap-In(s)
Import-Module MSOnline -erroraction SilentlyContinue
Import-Module ActiveDirectory
#endregion
    
#region 4.	Main Script
# Screen Output
Write-Host
Write-Host "Creating User Mappings" -ForegroundColor DarkCyan

# Connecting to O365
Write-Host "   Connecting to Azure AD"
Connect-MsolService -credential $O365Cred

# Retrieving All Users
Write-Host "   Retrieving All Users"
$Users = Get-MsolUser -All | ?{$_.UserPrincipalName -like ("*{0}" -f $UserDomain)}

# Looping through all User Accounts
Write-Host
Write-Host "   Looping through Users Accounts" -ForegroundColor Magenta
foreach ($User in $Users)
{
    $UserCount++         
    try
    {
		$ADUser = Get-ADUser -Credential $ADCred -Server $UserDomain -Filter {UserPrincipalName -eq $User.UserPrincipalName}
		If ($ADUser -ne $null)
		{
			$SourceUser = $SP13Prefix + $ADUser.SamAccountName
			$TargetUser = $SPOPrefix + $User.UserPrincipalName
			$Out = ("<Mapping Source='{0}' Target='{1}' />" -f $SourceUser,$TargetUser)
			out-file -filepath $OutPath -append -inputobject $Out
			Write-Host ("      {0}" -f $Out)
		}
		else 
		{
			Write-Host ("      User not found: {0}" -f $User.UserPrincipalName) -ForegroundColor Yellow
		}
    }
    catch
    {
        Write-Host ("      Error: {0}" -f $User.UserPrincipalName) -ForegroundColor Red
    }
}
out-file -filepath $OutPath -append -inputobject "</Mappings>"
Write-Host
Write-Host ("   User Count: {0}" -f $UserCount) -ForegroundColor Green
#endregion

#region Completed Screen Output
Complete-CatScript $Start[0]
#endregion