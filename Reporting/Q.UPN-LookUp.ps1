<#  Q.UPN-LookUp.ps1

This script iterates returns the UPN of a user or security group in Entra ID. It uses as it's input file the output
file from the Q.SPOnPrem-PermissionsMapping.ps1 script. 

Dependencies
    -   Azure App Registration (Quisitive Powershell)
    -	PnP.PowerShell Module
    -   Microsoft Graph Module
    -	..\General\Cat.GeneralFunctions.psm1
    -	Input CSV File in Input Files folder
    -   ..\_Output Files folder

Call script with parameters 
    cd C:\Quisitive\PowerShell\Reporting
    .\Q.UPN-LookUp.ps1 -inputFile "Q.SPOnPrem-Client-PermissionsMapping.csv" -finalOutputFile "Wave3-AllBatches-12112024-Permissions-GoLive.csv" -getSecurityGroups $true -getUsers $false -clientAppID "6bd4f453-6253-4f7a-be0c-717ecd464321" -clientCertPath "C:\Quisitive\PowerShell\App Reg\SharePointMigration2024.pfx" -clientTenant "sanfordhealth" -clientTenantID "02a24813-640d-46cd-b583-949f041372b0"

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
Param(
    [string] $inputFile = "Q.SPOnPrem-Client-PermissionsMapping.csv",
    [string] $clientTenant = "{Client Tenant Name}",
    [string] $clientTenantID = "{Client Tenant ID}",
    [string] $clientAppID = "{Azure App Registration ID}",
    [string] $clientCertPath = "C:\certs\PnPAppRegPowerShell.pfx",
    [string] $finalOutputFile = "Client-PermissionMapping-Final.csv", #This is the name of the file this script will generate and can be used as the final permissions mapping file to use with the Q.PnP-AddUsersToGroups.ps1 when applying permissions to SPO.
    [boolean] $getSecurityGroups = $true, #set to true to get security group UPNs. Also useful to check if they exist as the LoginName will be set to Not Found in CSV
    [boolean] $getUsers = $false #set to true to get user's UPNs. The Q.SPOnPrem-PermissionsMapping.ps1 script already gets this from the claims data but useful to validate that the user exists online, if not the LoginName will be set to Not Found in CSV
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
Start-Transcript -Path ("..\_Log Files\{0}.SP-UPN-LookUpDisc-{1}.log" -f $client, $scriptStart[1])
#endregion

#region 1.	Snapin(s), Modules & Environment Connection(s)

#endregion

#region 2.	Input & Output Files
# File Paths
$inPath = ("..\_Input Files\{0}" -f $inputFile)
$errorPath = ("..\_Output Files\{0}.UPN-LookUp-Errors-{1}.log" -f $client, $scriptStart[1])

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

# Output Files
$outputErrors = ("{0}\Q.UPN-LookUp-Errors.csv" -f $outputBasePath)
$outputPath = ("..\_Output Files\{0}" -f $finalOutputFile)
#endregion

#region 3.	Variables

# User Inputs
Write-Host
Write-Host "Enter Connection Information" -ForegroundColor Cyan
$clientCertPass = ConvertTo-SecureString -String (Read-Host -Prompt "   Enter Azure Tenant certificate password for the app registration") -AsPlainText -Force

# Output Parameters & Variables
# Script Variables
$clientTenantDomain = ("{0}.onmicrosoft.com" -f $clientTenant)
$clientCert = Get-PnPAzureCertificate -CertificatePath $clientCertPath -CertificatePassword $clientCertPass

# Initialize an array to store results
$updatedResults = @()

# Initialize UPN variable
$upn = ""

# Output Parameters & Variables
Write-Host
Write-Host "Script Parameters" -ForegroundColor Cyan
Write-Host ("   Input File: {0}" -f $inputFile)
Write-Host ("	Output File: {0}" -f $outputErrors)
Write-Host ("   Client: {0}" -f $client)
Write-Host ("   App ID: {0}" -f $clientAppID)
Write-Host ("   Client Tenant Prefix: {0}" -f $clientTenant)
Write-Host
Write-Host "User Inputs" -ForegroundColor Cyan
Write-Host ("   Certificate Thumbprint: {0}" -f $clientCert.Thumbprint)
Write-Host
Write-Host "Script Variables" -ForegroundColor Cyan
Write-Host ("   Microsoft Tenant: {0}" -f $clientTenantDomain)

#endregion

#region 4.  Functions
function Write-QCSVErrorInfo { 
	param(
		[psobject] $funError,
		[psobject] $funJobInfo,
		[string] $funErrorOn,
		[string] $funOutputPath
	)

	Write-Host (" - {0}" -f $funError) -ForegroundColor Red
	Write-Host
	Write-Host ("Exception: {0}" -f $funError.Exception) -ForegroundColor Red
	Write-Host ("Error Details: {0}" -f $funError.ErrorDetails) -ForegroundColor Red
	Write-Host ("Stack Trace: {0}" -f $funError.ScriptStackTrace) -ForegroundColor Red
	$errorRecord = New-Object -TypeName PSObject -Property (@{
		srcSiteURL=$funJobInfo.srcSiteURL;
		errorOn=$funErrorOn;
		errorMessage=$funError})
	$errorRecord | Export-Csv -Path $funOutputPath -NoTypeInformation -Append
	Remove-Variable errorRecord
}
#endregion

#region 5.	Main Script

try {
    # Connect to Microsoft Graph
    Connect-MgGraph -ClientId $clientAppID -TenantId $clientTenantID -CertificateThumbprint $clientCert.Thumbprint -NoWelcome
} 
catch {
    Write-Host ("Error connecting to Azure: {0}" -f $_) -ForegroundColor Red
}

# Initialize an array to store updated results
$updatedResults = @()

# Looping through input file
:userLoop ForEach ($usrgroup in $inFile){
	Write-Host
    Write-Host ("{0} - Quering User/Group: " -f (Get-Date -Format "yy-MM-dd HH:mm:ss")) -NoNewline

    $memberName = $usrgroup.userName 

    Write-Host (" {0}" -f $memberName) -ForegroundColor Green

    try {
        # Get user info
        $user = Get-MgUser -Filter "DisplayName eq '$memberName'" 
        If($user) {   
            Write-Host("Found User: {0}" -f $user.DisplayName) -ForegroundColor Green -ErrorAction Stop
            $upn = $user.UserPrincipalName
        } else {            
            # If not found as a user, get as a group
            $group = Get-MgGroup -Filter "DisplayName eq '$memberName'" | Select-Object "DisplayName"
            If($group) {
                Write-Host("Found Group: {0}" -f $group.DisplayName) -ForegroundColor Green  -ErrorAction Stop
                $upn = $memberName
            } else {
                $upn = "Not Found"
            }
        }
    } catch {
        Write-Host
        Write-Host ("Could not find UPN for member: {0}" -f $memberName) -ForegroundColor Yellow
        Write-QCSVErrorInfo -funcError $_ -funcObj $memberName -funOutputPath $errorPath
        $upn = "Not Found"
    }

    # Add the UPN as the Login Name property to the entry
    $updatedResults = [PSCustomObject]@{
        userName    = $usrgroup.userName
        loginName   = $(
                        If (($getSecurityGroups -and ($usrGroup.loginType.ToLower() -eq "securitygroup")) -or ($getUsers -and ($usrGroup.loginType.ToLower() -eq "user"))){$upn} 
                        else {If($upn -ne "Not Found") {$usrgroup.loginName} else {$upn}}
                       )
        loginType   = $usrgroup.loginType
        srcSiteURL  = $usrgroup.srcSiteURL
        siteUrl     = $usrgroup.siteUrl 
        groupName   = $usrgroup.groupName
        groupType   = $usrgroup.groupType
        groupRole   = $usrgroup.groupRole
        groupRoles  = $usrgroup.groupRoles
    }

    # Export updated results to CSV file
    Write-Host("UPN: {0}" -f $upn) -ForegroundColor DarkGray
    $updatedResults | Export-Csv -Path $outputPath -NoTypeInformation -Append

    Clear-Variable upn
    Clear-Variable memberName

}
#endregion


Write-Host ("Updated details to file {0}" -f $outputPath) -ForegroundColor DarkGray

#region Completed Script
# Output Script Info
Complete-CatScript $scriptStart[0]

# Stop logging
Stop-Transcript
#endregion

