<# Cat.Files-DiscoverFiles.ps1

This script scans a list of directories from CSV and ouputs information about the files.

Dependencies
    -	Scripts Solution Folder structure
        -   _Input Files - Input file must exist here
        -   _Output Files - Output file will be written here
        -   General - Leverages the Cat.GeneralFunctions.psm1
        -   Reporting - Script lives in this folder

Call Script with Input File
	.\Cat.Files-DiscoverFiles.ps1 Cat.FileDirectories-Test.csv

CSV Example
	ServLoc,Drive,Path
    DEN,C,C:\folder1
    DEN,D,D:\folder2

Table of Contents
	1.	Input & Output Files
	2.	Variables
	3.	Main Script
#>

#region Script Parameters
Param([string]$InputFile)
#endregion

#region Start Date / Time
# Set PS location
Set-Location $PSScriptRoot

# Get Script Parent Location
$LocSplit = $PSScriptRoot.Split("\")
$LocReplace = $LocSplit[($LocSplit.Count - 1)]
$Global:Location = $PSScriptRoot.Replace($LocReplace,"")

# Remove Existing Catapult Modules to ensure PS session has most up to date functions
Get-Module Cat.* | Remove-Module

# Load General Functions
Import-Module ("{0}General\Cat.GeneralFunctions.psm1" -f $Location)

# Get Start Date / Time
$Start = Start-CatScript
#endregion

#region 1.	Input & Output Files
# File Paths
$InPath = ("{0}_Input Files\{1}" -f $Location, $InputFile)


# Import Input File(s)
$InFile = import-csv $InPath -Encoding Unicode
#endregion

#region 3.	Main Script
# Screen Output
Write-Host
Write-Host ("Scanning File Directories from {0}" -f $InPath) -ForegroundColor Magenta

# Iterate through each Directory in CSV
$TotalErrors = @()
$TotalSuccesses = 0
ForEach ($Directory in $InFile)
{
    # Get Start Date / Time
    $StartTime = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    #endregion

    # Screen output
    Write-Host ("   Scanning directory: {0} For Location: {1}" -f $Directory.Path, $Directory.ServLoc)

    # Create Output File(s)
    $OutPath = ("{0}_Output Files\FileList-{1}-{2}.txt" -f $Location, $Directory.ServLoc, $Directory.Drive)
    $PermPath = ("{0}_Output Files\FileListPerms-{1}-{2}.txt" -f $Location, $Directory.ServLoc, $Directory.Drive)
    $ErrorPath = ("{0}_Error Files\FileListErrors-{1}-{2}.txt" -f $Location, $Directory.ServLoc, $Directory.Drive)
    Out-File -FilePath $OutPath -InputObject "Folder|FullName|Path|FileName|BaseName|Extension|CreatedTime|LastWriteTime|LastAccessTime|CreationTime|Length|Owner"
    Out-File -FilePath $PermPath -InputObject "FullName|Permission|Identity|Inherited"

    # Get all child objects
    $DirErrors = @()
    $AllChild = Get-ChildItem -Recurse $Directory.Path -ErrorAction SilentlyContinue -ErrorVariable DirErrors

    # Iterate through all child objects
    ForEach($Child in $AllChild)
    {
        # Get Owner
        $Owner = Get-ACL $Child.FullName
        
        # Output to File List CSV
        Out-File -FilePath $OutPath -Append -InputObject ("{0}|{1}|{2}|{3}|{4}|{5}|{6}|{7}|{8}|{9}|{10}|{11}" -f $Child.PSIsContainer, $Child.FullName, $Child.DirectoryName, $Child.Name, $Child.BaseName, $Child.Extension, $Child.CreationTime, $Child.LastWriteTime, $Child.LastAccessTime, $Child.CreationTime, $Child.Length, $Owner.Owner)

        # Get Permissions
        $PermErrors = @()
        $Perms = Get-ACL $Child.FullName -AllCentralAccessPolicies -ErrorAction SilentlyContinue -ErrorVariable PermErrors

        # Loop through Permissions
        ForEach ($Perm in $Perms.Access)
        {
            # Output to Perm List CSV
            Out-File -FilePath $PermPath -Append -InputObject ("{0}|{1}|{2}|{3}" -f $Child.FullName, $Perm.FileSystemRights, $Perm.IdentityReference, $Perm.IsInherited)
        }
    }

    # Output Counts
    $TotalErrors += $DirErrors
    $TotalErrors += $PermErrors
    $TotalSuccesses += $AllChild.Count

    $FinishTime = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
	$TimeTaken = New-TimeSpan -Start $StartTime -End $FinishTime

    # Create Error File
    Out-File -FilePath $ErrorPath -InputObject $TotalErrors
    Out-File -FilePath $ErrorPath -Append -InputObject ("Successes: {0}   Total Errors: {1}" -f $TotalSuccesses, $TotalErrors.Count)
    Out-File -FilePath $ErrorPath -Append -InputObject ("TimeTaken: {0}" -f $TimeTaken)
    #endregion

}

#region Completed Screen Output
Complete-CatScript $Start[0]
#endregion