<#  Cat.SG-CreateUserMappingFile.ps1

This script automates the creation of a ShareGate User Mapping file from a CSV list of users. You need to adjust line 89-92 based on the source.

Dependencies
    -	Input CSV File in Input Files folder - Script Parameter

Call script with parameters
    cd C:\GitHub\CatPowerShell\ShareGate\
	.\Cat.SG-CreateUserMappingFile.ps1 -userList enabled_SAM.csv -userMapping "D4C-UserMapping.sgum" -unresolvedUserEmail "sharegate@nrpdev.com"

CSV Example
sourceName,sourceUserEmail,destinationName,destinationUserEmail
Cindy Minor,cminor@domain1.com,Cindy Minor,cminor@domain2.com

Table of Contents
	1.	Input & Output Files
	2.	Variables
	3.	Snapin(s), Modules & Environment Connection(s)
    4.	Main Script
#>

#region Script Parameters
Param(
        [String] $userList = "{Client Export from AD/Entra}",
        [String] $userMapping = "{Client}.SG-UserMapping.sgum",
        [String] $unresolvedUserEmail = "{Client Service Account Email}"
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
Start-Transcript -Path ("..\_Log Files\Cat.SG-CreateUserMappingFile-{0}.log" -f $scriptStart[1])
#endregion

#region 1.	Input & Output Files
# File Paths
$inPath = ("..\_Input Files\{0}" -f $userList)

# Import Input File(s)
# Load CSV file
If (Test-Path ($inPath)){
    $inFile = Import-Csv $inPath #-Encoding Unicode
}
else {
	Write-Host
	throw ("The users file does not exist: {0}" -f $inPath)	
}

$outPath = ("..\_Output Files\{0}" -f $userMapping)
#endregion

#region 2.	Variables
# Output Parameters & Variables
Set-Variable srcSite, dstSite, srcList, dstList
Write-Host
Write-Host "Script Parameters" -ForegroundColor Cyan
Write-Host ("   User List: {0}" -f $userList)
Write-Host ("   User Mapping: {0}" -f $userMapping)
Write-Host ("   Unresolved User Email: {0}" -f $unresolvedUserEmail)
#endregion

#region 3.	Snapin(s), Modules & Environment Connection(s)
# Load Custom Modules
Import-Module "..\Reporting\Cat.ReportingFunctions.psm1"
#endregion

#region 4.	Main Script
Write-Host
Write-Host ("Generating ShareGate Mapping File") -ForegroundColor Magenta

# ShareGate user mapping Header
$mapping = "<?xml version=""1.0""?>"
$mapping += "<UserAndGroupMappings xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"">"
$mapping += "<Unresolved AccountName=""$unresolvedUserEmail"" DisplayName=""Migration Account"" PrincipalType=""User"" />"
$mapping += "<Mappings>"

foreach ($user in $inFile){ 
    Write-Host ("Adding row for {0}" -f $user) -ForegroundColor Cyan
    # SharePoint On premises claim
    #$srcUserEmail = "i:0#.w|" + $user.sourceUserEmail
    # File share format
    $srcUserEmail = $user.sourceUserEmail.Replace("\","\\")
    $srcName = $user.sourceName
    $destUserEmail = "i:0#.f|membership|" + $user.destinationUserEmail
    $destName = $user.destinationName

    $mapping += "<Mapping>"
    $mapping += "<Source AccountName=""$srcUserEmail"" DisplayName=""$srcName"" PrincipalType=""User"" />"
    $mapping += "<Destination AccountName=""$destUserEmail"" DisplayName=""$destName"" PrincipalType=""User"" />"
    $mapping += "</Mapping>"   
}

# ShareGate user mapping footer
$mapping += "</Mappings>"
$mapping += "</UserAndGroupMappings>"

# Output user mapping information
$mapping | Out-File -FilePath $outPath -NoClobber
Write-Host
Write-Host ("Successfully exported to: {0}" -f $outPath) -ForegroundColor Green

#region Completed Script
# Output Script Info
Complete-CatScript $scriptStart[0]

# Stop logging
Stop-Transcript
#endregion