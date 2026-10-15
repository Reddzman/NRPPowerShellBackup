<#  Cat.AD-CreateTestUsers.ps1

The Script creates test users to your demo domain based on first and last namnes from csv

Dependencies
    -	Active Directory Module
    -   CSV File with first & last names

Call script with parameters
	.\Cat.AD-CreateTestUsers.ps1 -InputFile Cat.AD-TestUsers.csv

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
Param(	[string] $inputFile = "Dev\Cat.AD-TestUsers.csv",
        [int] $numUsers = 100,
        [string] $orgUnit = "OU=Users,OU=SP2019Users,DC=sp19test,DC=local"
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
Start-Transcript -Path ("..\_Log FIles\CreateTestUsers-{0}.log" -f $scriptStart[1])
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

# Script Variables
$userDepts = @("IT","Finance","Logistics","Sourcing","Human Resources")
$userFirstNames = $inFile.Firstname
$userLastNames = $inFile.Lastname

# Output Parameters & Variables
Write-Host
Write-Host "Script Parameters" -ForegroundColor Cyan
Write-Host ("   Input File: {0}" -f $inputFile)
Write-Host ("   Number of Users: {0}" -f $numUsers)
Write-Host ("   Organizational Unit: {0}" -f $orgUnit)
Write-Host
Write-Host "User Inputs" -ForegroundColor Cyan
Write-Host ("   password: {0}" -f $password.GetNetworkCredential().Password)
Write-Host
Write-Host "User Departments" -ForegroundColor Cyan
Write-Host $userDepts
#endregion

#region 3.	Snapin(s), Modules & Environment Connection(s)
# Load MS / PnP PowerShell Modules
Import-Module ActiveDirectory
#endregion
    
#region 4.	Main Script
# Connecting To SP Online
Write-Host
Write-Host "Creating Test Users" -ForegroundColor Magenta

# Looping until number of users is exceeded
while ($numUsers -gt 0){
    # Picking random First & Last Name
    Write-Host
    Write-Host "   Generating First and Last Name: " -NoNewline
    $i = Get-Random -Minimum 0 -Maximum $userFirstNames.count
    $firstname = $userFirstNames[$i]
    $i = Get-Random -Minimum 0 -Maximum $userLastNames.count
    $lastname = $userLastNames[$i]
    Write-Host ("{0} {1}" -f $firstname, $lastname)

    # Picking random Department
    Write-Host "   Generating Department: " -NoNewline
    $i = get-random -Minimum 0 -Maximum $userDepts.count
    $dept = $userDepts[$i]
    Write-Host $dept
    
    # Generate username and check for duplicates
    Write-Host "   Generating username: " -NoNewline
    $username = $firstname.Substring(0,1).tolower() + $lastname.tolower()
    $exit = 0
    $count = 1
    do{ 
        try { 
            $userexists = Get-AdUser -Identity $username
            $username = $firstname.Substring(0,1).tolower() + $lastname.tolower() + $count++
        }
        catch {
            $exit = 1
        }
    }
    while ($exit -eq 0)
    Write-Host $username

    # Set Displayname
    Write-Host "   Generating Display Name: " -NoNewline
    $displayName= ("{0} {1}" -f $firstname, $lastname)
    Write-Host $displayName

    # Set Displayname
    Write-Host "   Generating UPN: " -NoNewline
    $upn = ("{0}@{1}" -f $username, (get-addomain).DNSRoot)
    Write-Host $upn

    # Create the user
    Write-Host ("   Creating User: {0} OU: {1} - " -f $username, $orgUnit) -NoNewline
    New-ADUser –Name $displayName –DisplayName $displayName –SamAccountName $username -UserPrincipalName $upn -GivenName $firstname -Surname $lastname -description "Test User" -Path $orgUnit –Enabled $true –ChangePasswordAtLogon $false -PasswordNeverExpires $true -Department $dept -AccountPassword $password.Password 
    Write-Host "Successfully Created" -ForegroundColor Green
    
    # Decrease User Count
    $numUsers-- 
}
#endregion

#region Completed Script
# Output Script Info
Complete-CatScript $scriptStart[0]

# Stop logging
Stop-Transcript
#endregion