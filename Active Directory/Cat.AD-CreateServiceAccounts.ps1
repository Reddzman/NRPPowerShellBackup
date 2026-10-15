<#  Cat.AD-CreateServiceAccounts.ps1

The Script adds services accounts to your demo domain based on first and last namnes from csv

Dependencies
    -	Active Directory Module
    -   CSV File with service Accounts

Call script with parameters
	.\Cat.AD-CreateServiceAccounts.ps1 -InputFile Cat.AD-ServiceAccounts.csv

CSV Example
    Firstname,LastName
    Aaron,Alfort
    Abraham,Allendorf

Table of Contents
	1.	Input & Output Files
	2.	Variables
	3.	Snapin(s), Modules & Environment Connection(s)
	4.	Main Script
#>

#region Script Parameters
Param(	[string] $inputFile = "Dev\Cat.AD-SP19ServiceAccounts.csv",
        [string] $orgUnit = "OU=Service Accounts,OU=SP2019Users,DC=sp19test,DC=local"
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
Start-Transcript -Path ("..\_Log FIles\CreateServiceAccounts-{0}.log" -f $scriptStart[1])
#endregion

#region 1.	Input & Output Files
# File Paths
$inPath = ("..\_Input Files\{0}" -f $inputFile)

# Import Input File(s)
# Load CSV file
If (Test-Path ($inPath)){
	$inFile = Import-Csv $inPath -Encoding Unicode
}
else {
	Write-Host
	throw ("The jobs file does not exist: {0}" -f $inPath)	
}
#endregion

#region 2.	Variables
# User Inputs
Write-Host
Write-Host "User Input Variables" -ForegroundColor Cyan
Write-Host "   Enter your credentials. The password will be used for each new account:"
$password = Get-Credential -Message "Enter your credentials. The password will be used for each new account:" 

# Output Parameters & Variables
Write-Host
Write-Host "Script Parameters" -ForegroundColor Cyan
Write-Host ("   Input File: {0}" -f $inputFile)
Write-Host ("   Organizational Unit: {0}" -f $orgUnit)
Write-Host
Write-Host "User Inputs" -ForegroundColor Cyan
Write-Host ("   password: {0}" -f $password.GetNetworkCredential().Password)
#endregion

#region 3.	Snapin(s), Modules & Environment Connection(s)
# Load MS / PnP PowerShell Modules
Import-Module ActiveDirectory
#endregion
    
#region 4.	Main Script
# Connecting To SP Online
Write-Host
Write-Host "Creating Service Accounts" -ForegroundColor Magenta

# Loop through accounts in CSV file
ForEach ($acct in $inFile){
    # Create the user
    Write-Host ("   Creating User: {0} OU: {1} - " -f $username, $orgUnit) -NoNewline
    $displayName= ("{0} {1}" -f $acct.firstName, $acct.lastName)
    $upn = ("{0}@{1}" -f $acct.username, (get-addomain).DNSRoot)
    New-ADUser –Name $displayName –DisplayName $displayName –SamAccountName $acct.username -UserPrincipalName $upn -GivenName $acct.firstName -Surname $acct.lastName -description ("Service Account: {0}" -f $displayName) -Path $orgUnit –Enabled $true –ChangePasswordAtLogon $false -PasswordNeverExpires $true -Department $acct.department -AccountPassword $password.Password 
    Write-Host "Successfully Created" -ForegroundColor Green
}
#endregion

#region Completed Script
# Output Script Info
Complete-CatScript $scriptStart[0]

# Stop logging
Stop-Transcript
#endregion