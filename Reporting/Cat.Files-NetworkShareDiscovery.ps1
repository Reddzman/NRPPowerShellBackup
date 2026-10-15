<# Cat.Files-NetworkShareDiscovery.ps1

This script scans a list of network shares from CSV and ouputs information about the folders and files. The shares should be broken
up into management chunks.  This is a long running script and it is helpful to be able to pick up where you left off.

Dependencies
    -	Scripts Solution Folder structure
    -   ..\_Input Files - Input file must exist here
    -   ..\_Output Files - Output file will be written here
    -   ..\General\Cat.GeneralFunctions.psm1

Call Script with Input File
	.\Cat.Files-NetworkShareDiscovery.ps1 -inputFile "Samples\Reporting\Cat.Files-NetworkShareDiscovery-Test.csv" -client "Testing"

CSV Example
	serverUNC,shareName
    Server.catapultsystems.com,Share1
    Server.catapultsystems.com,Share1\Folder1
    Server.catapultsystems.com,Share2
    Server.catapultsystems.com,Share2\Folder1
    Server.catapultsystems.com,Share2\Folder1\SubFolder1
    Server.catapultsystems.com,Share2\Folder2

Table of Contents
	1.	Input & Output Files
	2.	Variables
	3.  Functions
    4.  Get-CatDirectoryInfo
    5.  Get-CatFileInfo
    6.	Main Script
    7.  Creating Directories
    8.  Share Root Directories & Files
    9.  Level 1 Directories & Files
    10.  Level 2 Directories & Files
    11.  Remaining Directories & Files
#>

#region Script Parameters
Param(
        [string] $inputFile = "Cat.Files-NetworkShareDiscovery-Test.csv",
        [String] $client = "{Client}",
        [boolean] $getHidden = $true
)
#endregion

#region Start Date / Time
# Remove Existing Catapult Modules to ensure PS session has most up to date functions
Get-Module Cat.* | Remove-Module

# Load General Functions
Import-Module "..\General\Cat.GeneralFunctions.psm1"

# Get Start Date / Time
$global:scriptStart = Start-CatScript

# Start Transcript
Start-Transcript -Path ("..\_Log Files\{0}.Files-NetworkShareDiscovery-{1}.log" -f $client, $scriptStart[1])
#endregion

#region 1.	Input & Output Files
# File Paths
$inPath = ("..\_Input Files\{0}" -f $inputFile)

# Import Input File(s)
If (Test-Path ($inPath)){
	$inFile = Import-Csv $inPath -Encoding Unicode
}
else {
	Write-Host
	throw ("The input file does not exist: {0}" -f $inPath)	
}
#endregion

#region 2.	Variables
# Script Variables
$specialCharacters = "[``'`":]"
$outputBasePath = ("..\_Output Files\{0}\{1}" -f $client, $scriptStart[1])
$outputScanInfoFolder = ("{0}\details" -f $outputBasePath)
$outputScanInfo = ("{0}\ScanInformation.csv" -f $outputScanInfoFolder)
$outputFileInfo = ("{0}\files" -f $outputBasePath)
$outputDirInfo = ("{0}\directories" -f $outputBasePath)
$outputPermInfo = ("{0}\permissions" -f $outputBasePath)
$outputErrInfo = ("{0}\errors" -f $outputBasePath)

# Output Parameters & Variables
Write-Host
Write-Host "Script Parameters" -ForegroundColor Cyan
Write-Host ("   Input File: {0}" -f $inputFile)
Write-Host ("   Client: {0}" -f $client)
Write-Host
Write-Host "Script Variables" -ForegroundColor Cyan
Write-Host ("   Output Path: {0}" -f $outputBasePath)
Write-Host ("   Scan Info Path: {0}" -f $outputScanInfo)
Write-Host ("   File Info Path: {0}" -f $outputFileInfo)
Write-Host ("   Directory Info Path: {0}" -f $outputDirInfo)
Write-Host ("   Permissions Path: {0}" -f $outputPermInfo)
#endregion

#region 3.  Functions
    #region 4.  Get-CatDirectoryInfo
    <#  Get-CatDirectoryInfo
    This function exports directory information including permissions to CSV file for reporting.

    Command: Get-CatDirectoryInfo -directories $folders -sUNC $serverUNC -sName $shareName -dFile $outDirectoryInfo -pFile $outDirectoryPermissions -spacing $outSpacing -l1Color "DarkGreen" -l2Color "DarkYellow"

    Parameters
        directories - Array of sytem objects from Get-ChildItem -Directory,
        sUNC - UNC Name of the server the file share exists on
        sName - Share Name of the file share being scanned
        iFile - path & Name of CSV file to export directory info to
        pFile - path & Name of CSV file to export directory permissions info to
        spacing - Leading Spaced in write-host commands
        l1Color - Level one color for write-host commands
        l2Color - Level two color for write-host commands
    #>
    function Get-CatDirectoryInfo {
        param (
            [System.Collections.Queue] $directories,
            [string] $sUNC,
            [string] $sName,
            [string] $iFile,
            [string] $pFile,
            [string] $spacing = "   ",
            [string] $l1Color = "White",
            [string] $l2Color = "DarkGray"
        )

        $l2spacing = ("   {0}" -f $spacing)
        if ($directories.Count -gt 0) {
            # Exporting to CSV Directory Information
            # ((($_.SiteLogoDescription -replace $specialCharacters,'') -replace "`n",' ') -replace "`r",' ')
            Write-Host ("{0}{1} - Writing directories to CSV" -f $spacing, (Get-Date -Format "yy-MM-dd HH:mm:ss")) -ForegroundColor $l1Color
            $directories | ForEach-Object {
                $directory = New-Object -TypeName psobject -Property (@{
                    ServerUNC=$sUNC;
                    ShareName=$sName;
                    Name=((($_.Name -replace $specialCharacters,'') -replace "`n",' ') -replace "`r",' ');
                    Parent=((($_.Parent -replace $specialCharacters,'') -replace "`n",' ') -replace "`r",' ');
                    Root=((($_.Root -replace $specialCharacters,'') -replace "`n",' ') -replace "`r",' ');
                    BaseName=((($_.BaseName -replace $specialCharacters,'') -replace "`n",' ') -replace "`r",' ');
                    FullName=((($_.FullName -replace $specialCharacters,'') -replace "`n",' ') -replace "`r",' ');
                    LastAccessTime=$_.LastAccessTime;
                    LastAccessTimeUtc=$_.LastAccessTimeUtc;
                    LastWriteTime=$_.LastWriteTime;
                    LastWriteTimeUtc=$_.LastWriteTimeUtc;
                    LinkType=((($_.LinkType -replace $specialCharacters,'') -replace "`n",' ') -replace "`r",' ');
                    Mode=((($_.Mode -replace $specialCharacters,'') -replace "`n",' ') -replace "`r",' ');
                    ModeWithoutHardLink=((($_.ModeWithoutHardLink -replace $specialCharacters,'') -replace "`n",' ') -replace "`r",' ');
                    Target=((($_.Target -replace $specialCharacters,'') -replace "`n",' ') -replace "`r",' ');
                    PSChildName=((($_.PSChildName -replace $specialCharacters,'') -replace "`n",' ') -replace "`r",' ');
                    PSIsContainer=((($_.PSIsContainer -replace $specialCharacters,'') -replace "`n",' ') -replace "`r",' ');
                    PSParentPath=((($_.PSParentPath -replace $specialCharacters,'') -replace "`n",' ') -replace "`r",' ');
                    PSPath=((($_.PSPath -replace $specialCharacters,'') -replace "`n",' ') -replace "`r",' ');
                    PSProvider=((($_.PSProvider -replace $specialCharacters,'') -replace "`n",' ') -replace "`r",' ');
                    Attributes=((($_.Attributes -replace $specialCharacters,'') -replace "`n",' ') -replace "`r",' ');
                    CreationTime=$_.CreationTime;
                    CreationTimeUtc=$_.CreationTimeUtc;
                    Exists=((($_.Exists -replace $specialCharacters,'') -replace "`n",' ') -replace "`r",' ');
                    Extension=((($_.Extension -replace $specialCharacters,'') -replace "`n",' ') -replace "`r",' ')
                })
                $directory | Select-Object -Property ServerUNC,ShareName,Name,Parent,Root,BaseName,FullName,LastAccessTime,LastAccessTimeUtc,`
                    LastWriteTime,LastWriteTimeUtc,LinkType,Mode,ModeWithoutHardLink,Target,PSChildName,PSIsContainer,PSParentPath,PSPath,`
                    PSProvider,Attributes,CreationTime,CreationTimeUtc,Exists,Extension `
                    | Export-Csv -Path $iFile -UseQuotes AsNeeded -NoTypeInformation -Append
                Remove-Variable directory
                
                # Retrieving Directory Permissions Information
                Write-Host ("{0}{1} - Retrieving directory permission" -f $spacing, (Get-Date -Format "yy-MM-dd HH:mm:ss")) -ForegroundColor $l1Color
                $dirACLs = New-Object System.Collections.Queue;
                try {
                    $_ | Get-ACL | ForEach-Object{$dirACLs.Enqueue($_)}
                }
                catch {
                    Write-CatErrorCSV -funcCSVItem $fileShare -funcAction "Retrieving directory permission" -funcException $_.Exception -funcDetails $_.ErrorDetails -funcStackTrace $_.ScriptStackTrace -funcFilePath $filePath
                }
                Get-CatPermInfo -perms $dirACLs -sUNC $fileShare.serverUNC -sName $fileShare.shareName -pFile $csvPerm -spacing $l2spacing -color $l2Color
                Remove-Variable dirACLs
            }
            Remove-Variable directories
        }
        else {
            Write-Host ("{0}{1} - No directories found" -f $spacing, (Get-Date -Format "yy-MM-dd HH:mm:ss")) -ForegroundColor $l1Color
        }
    }
    #endregion

    #region 5.  Get-CatFileInfo
    <#  Get-CatFileInfo
    This function exports file information including permissions to CSV file for reporting.

    Command: Get-CatFileInfo -files $folders -sUNC $serverUNC -sName $shareName -fFile $outDirectoryInfo -pFile $outDirectoryPermissions -spacing $outSpacing -l1Color "DarkGreen" -l2Color "DarkYellow"

    Parameters
        directories - Array of sytem objects from Get-ChildItem -Directory,
        sUNC - UNC Name of the server the file share exists on
        sName - Share Name of the file share being scanned
        iFile - path & Name of CSV file to export directory info to
        pFile - path & Name of CSV file to export directory permissions info to
        spacing - Leading Spaced in write-host commands
        l1Color - Level one color for write-host commands
        l2Color - Level two color for write-host commands
    #>
    function Get-CatFileInfo {
        param (
            [System.Collections.Queue] $files,
            [string] $sUNC,
            [string] $sName,
            [string] $iFile,
            [string] $pFile,
            [string] $spacing = "   ",
            [string] $l1Color = "White",
            [string] $l2Color = "DarkGray"
        )

        $l2spacing = ("   {0}" -f $spacing)
        if ($files.Count -gt 0) {
            # Adding Properties to File Info
            Write-Host ("{0}{1} - Writing files to CSV" -f $spacing, (Get-Date -Format "yy-MM-dd HH:mm:ss")) -ForegroundColor $l1Color
            while ($files.Count -ne 0) {
                $files.Dequeue() | ForEach-Object {
                    $file = New-Object -TypeName psobject -Property (@{
                        ServerUNC=$sUNC;
                        ShareName=$sName;
                        Owner=((($_.Owner -replace $specialCharacters,'') -replace "`n",' ') -replace "`r",' ');
                        Group=((($_.Group -replace $specialCharacters,'') -replace "`n",' ') -replace "`r",' ');
                        Name=((($_.Name -replace $specialCharacters,'') -replace "`n",' ') -replace "`r",' ');
                        BaseName=((($_.BaseName -replace $specialCharacters,'') -replace "`n",' ') -replace "`r",' ');
                        LinkType=((($_.LinkType -replace $specialCharacters,'') -replace "`n",' ') -replace "`r",' ');
                        Mode=((($_.Mode -replace $specialCharacters,'') -replace "`n",' ') -replace "`r",' ');
                        ModeWithoutHardLink=((($_.ModeWithoutHardLink -replace $specialCharacters,'') -replace "`n",' ') -replace "`r",' ');
                        Target=((($_.Target -replace $specialCharacters,'') -replace "`n",' ') -replace "`r",' ');
                        PSChildName=((($_.PSChildName -replace $specialCharacters,'') -replace "`n",' ') -replace "`r",' ');
                        PSIsContainer=((($_.PSIsContainer -replace $specialCharacters,'') -replace "`n",' ') -replace "`r",' ');
                        PSParentPath=((($_.PSParentPath -replace $specialCharacters,'') -replace "`n",' ') -replace "`r",' ');
                        Attributes=((($_.Attributes -replace $specialCharacters,'') -replace "`n",' ') -replace "`r",' ');
                        CreationTime=$_.CreationTime;
                        CreationTimeUtc=$_.CreationTimeUtc;
                        Directory=((($_.Directory -replace $specialCharacters,'') -replace "`n",' ') -replace "`r",' ');
                        DirectoryName=((($_.DirectoryName -replace $specialCharacters,'') -replace "`n",' ') -replace "`r",' ');
                        Exists=((($_.Exists -replace $specialCharacters,'') -replace "`n",' ') -replace "`r",' ');
                        Extension=((($_.Extension -replace $specialCharacters,'') -replace "`n",' ') -replace "`r",' ');
                        FullName=((($_.FullName -replace $specialCharacters,'') -replace "`n",' ') -replace "`r",' ');
                        IsReadOnly=((($_.IsReadOnly -replace $specialCharacters,'') -replace "`n",' ') -replace "`r",' ');
                        LastAccessTime=$_.LastAccessTime;
                        LastAccessTimeUtc=$_.LastAccessTimeUtc;
                        LastWriteTime=$_.LastWriteTime;
                        LastWriteTimeUtc=$_.LastWriteTimeUtc;
                        Length=$_.Length
                        FileVersionRaw=((($_.VersionInfo.FileVersionRaw -replace $specialCharacters,'') -replace "`n",' ') -replace "`r",' ');
                        ProductVersionRaw=((($_.VersionInfo.ProductVersionRaw -replace $specialCharacters,'') -replace "`n",' ') -replace "`r",' ');
                        Comments=((($_.VersionInfo.Comments -replace $specialCharacters,'') -replace "`n",' ') -replace "`r",' ');
                        CompanyName=((($_.VersionInfo.CompanyName -replace $specialCharacters,'') -replace "`n",' ') -replace "`r",' ');
                        FileBuildPart=((($_.VersionInfo.FileBuildPart -replace $specialCharacters,'') -replace "`n",' ') -replace "`r",' ');
                        FileDescription=((($_.VersionInfo.FileDescription -replace $specialCharacters,'') -replace "`n",' ') -replace "`r",' ');
                        FileMajorPart=((($_.VersionInfo.FileMajorPart -replace $specialCharacters,'') -replace "`n",' ') -replace "`r",' ');
                        FileMinorPart=((($_.VersionInfo.FileMinorPart -replace $specialCharacters,'') -replace "`n",' ') -replace "`r",' ');
                        FileName=((($_.VersionInfo.FileName -replace $specialCharacters,'') -replace "`n",' ') -replace "`r",' ');
                        FilePrivatePart=((($_.VersionInfo.FilePrivatePart -replace $specialCharacters,'') -replace "`n",' ') -replace "`r",' ');
                        FileVersion=((($_.VersionInfo.FileVersion -replace $specialCharacters,'') -replace "`n",' ') -replace "`r",' ');
                        InternalName=((($_.VersionInfo.InternalName -replace $specialCharacters,'') -replace "`n",' ') -replace "`r",' ');
                        IsDebug=((($_.VersionInfo.IsDebug -replace $specialCharacters,'') -replace "`n",' ') -replace "`r",' ');
                        IsPatched=((($_.VersionInfo.IsPatched -replace $specialCharacters,'') -replace "`n",' ') -replace "`r",' ');
                        IsPrivateBuild=((($_.VersionInfo.IsPrivateBuild -replace $specialCharacters,'') -replace "`n",' ') -replace "`r",' ');
                        IsPreRelease=((($_.VersionInfo.IsPreRelease -replace $specialCharacters,'') -replace "`n",' ') -replace "`r",' ');
                        IsSpecialBuild=((($_.VersionInfo.IsSpecialBuild -replace $specialCharacters,'') -replace "`n",' ') -replace "`r",' ');
                        Language=((($_.VersionInfo.Language -replace $specialCharacters,'') -replace "`n",' ') -replace "`r",' ');
                        LegalCopyright=((($_.VersionInfo.LegalCopyright -replace $specialCharacters,'') -replace "`n",' ') -replace "`r",' ');
                        LegalTrademarks=((($_.VersionInfo.LegalTrademarks -replace $specialCharacters,'') -replace "`n",' ') -replace "`r",' ');
                        OriginalFileName=((($_.VersionInfo.OriginalFileName -replace $specialCharacters,'') -replace "`n",' ') -replace "`r",' ');
                        PrivateBuild=((($_.VersionInfo.PrivateBuild -replace $specialCharacters,'') -replace "`n",' ') -replace "`r",' ');
                        ProductBuildPart=((($_.VersionInfo.ProductBuildPart -replace $specialCharacters,'') -replace "`n",' ') -replace "`r",' ');
                        ProductMajorPart=((($_.VersionInfo.ProductMajorPart -replace $specialCharacters,'') -replace "`n",' ') -replace "`r",' ');
                        ProductMinorPart=((($_.VersionInfo.ProductMinorPart -replace $specialCharacters,'') -replace "`n",' ') -replace "`r",' ');
                        ProcuctName=((($_.VersionInfo.ProcuctName -replace $specialCharacters,'') -replace "`n",' ') -replace "`r",' ');
                        ProductPrivatePart=((($_.VersionInfo.ProductPrivatePart -replace $specialCharacters,'') -replace "`n",' ') -replace "`r",' ');
                        ProductVersion=((($_.VersionInfo.ProductVersion -replace $specialCharacters,'') -replace "`n",' ') -replace "`r",' ');
                        SpecialBuild=((($_.VersionInfo.SpecialBuild -replace $specialCharacters,'') -replace "`n",' ') -replace "`r",' ')
                    })
                    $file | Select-Object -Property ServerUNC,ShareName,Owner,Group,Name,BaseName,LinkType,Mode,ModeWithoutHardLink,Target,`
                        PSChildName,PSIsContainer,PSParentPath,Attributes,CreationTime,CreationTimeUtc,Directory,DirectoryName,Exists,Extension,`
                        FullName,IsReadOnly,LastAccessTime,LastAccessTimeUtc,LastWriteTime,LastWriteTimeUtc,Length,FileVersionRaw,ProductVersionRaw,`
                        Comments,CompanyName,FileBuildPart,FileDescription,FileMajorPart,FileMinorPart,FileName,FilePrivatePart,FileVersion,`
                        InternalName,IsDebug,IsPatched,IsPrivateBuild,IsPreRelease,IsSpecialBuild,Language,LegalCopyright,LegalTrademarks,`
                        OriginalFileName,PrivateBuild,ProductBuildPart,ProductMajorPart,ProductMinorPart,ProcuctName,ProductPrivatePart,`
                        ProductVersion,SpecialBuild | Export-Csv -Path $iFile -UseQuotes AsNeeded -NoTypeInformation -Append
                    Remove-Variable file

                    # Retrieving File Permissions Information
                    Write-Host ("{0}{1} - Retrieving file permission" -f $spacing, (Get-Date -Format "yy-MM-dd HH:mm:ss")) -ForegroundColor $l1Color
                    $fileACLs = New-Object System.Collections.Queue;
                    try {
                        $_ | Get-ACL | ForEach-Object{$fileACLs.Enqueue($_)}
                    }
                    catch {
                        Write-CatErrorCSV -funcCSVItem $fileShare -funcAction "Retrieving file permission" -funcException $_.Exception -funcDetails $_.ErrorDetails -funcStackTrace $_.ScriptStackTrace -funcFilePath $filePath
                    }
                    Get-CatPermInfo -perms $fileACLs -sUNC $fileShare.serverUNC -sName $fileShare.shareName -pFile $csvPerm -spacing $l2spacing -color $l2Color
                    Remove-Variable fileACLs
                }
            }
            Remove-Variable files
        }
        else {
            Write-Host ("{0}{1} - No files found" -f $spacing, (Get-Date -Format "yy-MM-dd HH:mm:ss")) -ForegroundColor $l1Color
        }
    }
    #endregion

    #region 6.  Get-CatPermInfo
    <#  Get-CatFileInfo
    This function exports permissions information including to CSV file for reporting.

    Command: Get-CatPermInfo -perms $dirACLs -sUNC $fileShare.serverUNC -sName $fileShare.shareName -pFile $csvPerm

    Parameters
        perms - .NET Collection containing permissions
        sUNC - UNC Name of the server the file share exists on
        sName - Share Name of the file share being scanned
        pFile - path & Name of CSV file to export directory permissions info to
        spacing - Leading Spaced in write-host commands
        l1Color - Level one color for write-host commands
        l2Color - Level two color for write-host commands
    #>
    function Get-CatPermInfo {
        param (
            [System.Collections.Queue] $perms,
            [string] $sUNC,
            [string] $sName,
            [string] $pFile,
            [string] $spacing = "   ",
            [string] $color = "White"
        )

        Write-Host ("{0}{1} - Writing permissions to CSV" -f $spacing, (Get-Date -Format "yy-MM-dd HH:mm:ss")) -ForegroundColor $color
        if ($perms.Count -gt 0) {
            # Adding Properties to Share Permissions
            while ($perms.Count -ne 0) {
                $perms.Dequeue() | ForEach-Object {
                    $perm = $_
                    $perm.Access | ForEach-Object {
                        $acl = New-Object -TypeName psobject -Property (@{
                            ServerUNC=$sUNC;
                            ShareName=$sName;
                            Owner=((($perm.Owner -replace $specialCharacters,'') -replace "`n",' ') -replace "`r",' ');
                            Group=((($perm.Group -replace $specialCharacters,'') -replace "`n",' ') -replace "`r",' ');
                            CentralAccessPolicyId=((($perm.CentralAccessPolicyId -replace $specialCharacters,'') -replace "`n",' ') -replace "`r",' ');
                            Path=((($perm.Path -replace $specialCharacters,'') -replace "`n",' ') -replace "`r",' ');
                            Sddl=((($perm.Sddl -replace $specialCharacters,'') -replace "`n",' ') -replace "`r",' ');
                            PSChildName=((($perm.PSChildName -replace $specialCharacters,'') -replace "`n",' ') -replace "`r",' ');
                            PSParentPath=((($perm.PSParentPath -replace $specialCharacters,'') -replace "`n",' ') -replace "`r",' ');
                            PSPath=((($perm.PSPath -replace $specialCharacters,'') -replace "`n",' ') -replace "`r",' ');
                            PSProvider=((($perm.PSProvider -replace $specialCharacters,'') -replace "`n",' ') -replace "`r",' ');
                            AreAccessRulesCanonical=((($perm.AreAccessRulesCanonical -replace $specialCharacters,'') -replace "`n",' ') -replace "`r",' ');
                            AreAccessRulesProtected=((($perm.AreAccessRulesProtected -replace $specialCharacters,'') -replace "`n",' ') -replace "`r",' ');
                            AreAuditRulesCanonical=((($perm.AreAuditRulesCanonical -replace $specialCharacters,'') -replace "`n",' ') -replace "`r",' ');
                            AreAuditRulesProtected=((($perm.AreAuditRulesProtected -replace $specialCharacters,'') -replace "`n",' ') -replace "`r",' ');
                            FileSystemRights=((($_.FileSystemRights -replace $specialCharacters,'') -replace "`n",' ') -replace "`r",' ');
                            AccessControlType=((($_.AccessControlType -replace $specialCharacters,'') -replace "`n",' ') -replace "`r",' ');
                            IdentityReference=((($_.IdentityReference -replace $specialCharacters,'') -replace "`n",' ') -replace "`r",' ');
                            IsInherited=((($_.IsInherited -replace $specialCharacters,'') -replace "`n",' ') -replace "`r",' ');
                            InheritanceFlags=((($_.InheritanceFlags -replace $specialCharacters,'') -replace "`n",' ') -replace "`r",' ');
                            PropagationFlags=((($_.FileSystemRights -replace $specialCharacters,'') -replace "`n",' ') -replace "`r",' ');
                        })
                        $acl | Select-Object -Property ServerUNC,ShareName,Owner,Group,CentralAccessPolicyId,Path,Sddl,PSChildName,`
                            PSParentPath,PSPath,PSProvider,AreAccessRulesCanonical,AreAccessRulesProtected,AreAuditRulesCanonical,`
                            AreAuditRulesProtected,FileSystemRights,AccessControlType,IdentityReference,IsInherited,InheritanceFlags,`
                            PropagationFlags | Export-Csv -Path $pFile -UseQuotes AsNeeded -NoTypeInformation -Append
                        Remove-Variable acl
                    }
                    Remove-Variable perm
                }
            }
            Remove-Variable perms
        }
        else {
            Write-Host ("{0}{1} - No permissions found" -f $spacing, (Get-Date -Format "yy-MM-dd HH:mm:ss")) -ForegroundColor $color
        }
    }
    #endregion
#endregion

#region 6.	Main Script
#region 7.  Creating Directories
Write-Host
Write-Host ("Scanning File Directories: {0}" -f $InPath) -ForegroundColor Cyan

# Creating Directories
Write-Host "   Creating Directories"
New-Item -Path $outputFileInfo -ItemType Directory
New-Item -Path $outputDirInfo -ItemType Directory
New-Item -Path $outputPermInfo -ItemType Directory
New-Item -Path $outputScanInfoFolder -ItemType Directory
New-Item -Path $outputErrInfo -ItemType Directory

# Outputing Scan Information
Write-Host
Write-Host "   Exporting Scan Information"
$scanInfo = @()
$scanObj = New-Object -TypeName PSObject -Property (@{customerName=$client;scanDateTime=$scriptStart[1];fileType="Files"})
$scanInfo += $scanObj
$scanObj = New-Object -TypeName PSObject -Property (@{customerName=$client;scanDateTime=$scriptStart[1];fileType="Directories"})
$scanInfo += $scanObj
$scanObj = New-Object -TypeName PSObject -Property (@{customerName=$client;scanDateTime=$scriptStart[1];fileType="Permissions"})
$scanInfo += $scanObj
$scanInfo | Export-Csv -Path $outputScanInfo -NoTypeInformation -Append
Remove-Variable scanInfo
Remove-Variable scanObj
#endregion

# Iterate through each Directory in CSV
:loopShare foreach ($fileShare in $InFile)
{
    # Screen output
    Write-Host
    Write-Host ("{0} - Scanning Server: {1}   Share: {2}" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"), $fileShare.serverUNC, $fileShare.shareName) -ForegroundColor Magenta
    $csvFile = ("{0}\{1}-File-{2}.csv" -f $outputFileInfo, ($fileShare.shareName -replace '[\W]',''), (Get-Date -Format "yyMMddHHmmss"))
    $csvDir = ("{0}\{1}-Dir-{2}.csv" -f $outputDirInfo, ($fileShare.shareName -replace '[\W]',''), (Get-Date -Format "yyMMddHHmmss"))
    $csvPerm = ("{0}\{1}-Perm-{2}.csv" -f $outputPermInfo, ($fileShare.shareName -replace '[\W]',''), (Get-Date -Format "yyMMddHHmmss"))
    $csvErr = ("{0}\{1}-Err-{2}.csv" -f $outputErrInfo, ($fileShare.shareName -replace '[\W]',''), (Get-Date -Format "yyMMddHHmmss"))

    #region 8.  Share Root Permissions & Files
    $sharePath = ("\\{0}\{1}" -f $fileShare.serverUNC, $fileShare.shareName)
    $outSpacing = "      "

    #region Exporting to CSV Share Root Permissions
    $dirACLs = New-Object System.Collections.Queue;
    Write-Host ("   {0} - Retrieving Share Root Permissions: {1}" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"),$sharePath)
    try {
        Get-ACL -Path $sharePath | ForEach-Object{$dirACLs.Enqueue($_)}
    }
    catch {
        Write-CatErrorCSV -funcCSVItem $fileShare -funcAction "Retrieving Share Root Permissions" -funcException $_.Exception -funcDetails $_.ErrorDetails -funcStackTrace $_.ScriptStackTrace -funcFilePath $csvErr
    }
    Get-CatPermInfo -perms $dirACLs -sUNC $fileShare.serverUNC -sName $fileShare.shareName -pFile $csvPerm -spacing $outSpacing -color "DarkBlue"
    Remove-Variable dirACLs
    #endregion
  
    #region Exporting to CSV Share Root Files
    Write-Host ("   {0} - Retrieving Share Root Files: {1}" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"),$sharePath)
    $dirFiles = New-Object System.Collections.Queue;
    try {
        Get-ChildItem -Path $sharePath -File -Force:$getHidden -ErrorAction Stop | ForEach-Object{$dirFiles.Enqueue($_)}
    }
    catch {
        Write-CatErrorCSV -funcCSVItem $fileShare -funcAction "Retrieving Share Root Files" -funcException $_.Exception -funcDetails $_.ErrorDetails -funcStackTrace $_.ScriptStackTrace -funcFilePath $csvErr
    }
    Get-CatFileInfo -files $dirFiles -sUNC $fileShare.serverUNC -sName $fileShare.shareName -iFile $csvFile -pFile $csvPerm -spacing $outSpacing -l1Color "DarkBlue" -l2Color "DarkCyan"
    Remove-Variable dirFiles
    #endregion

    #region Retrieving Level 1 Directory Information
    Write-Host ("   {0} - Retrieving level 1 directories: {1}" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"),$sharePath)
    $dirLevel1Folders = New-Object System.Collections.Queue;
    try {
        Get-ChildItem -Path $sharePath -Directory -Force:$getHidden -ErrorAction Stop | ForEach-Object{$dirLevel1Folders.Enqueue($_)}
    }
    catch {
        Write-CatErrorCSV -funcCSVItem $fileShare -funcAction "Retrieving level 1 directories" -funcException $_.Exception -funcDetails $_.ErrorDetails -funcStackTrace $_.ScriptStackTrace -funcFilePath $csvErr
    }
    Get-CatDirectoryInfo -directories $dirLevel1Folders -sUNC $fileShare.serverUNC -sName $fileShare.shareName -iFile $csvDir -pFile $csvPerm -spacing $outSpacing -l1Color "DarkBlue" -l2Color "DarkCyan"
    #endregion
    #endregion

    :loopL1 while ($dirLevel1Folders.Count -ne 0) {
        $dirLevel1Folders.Dequeue() | ForEach-Object {
            #region 9.  Level 1 Directories & Files
            $dirL1 = $_
            $outSpacing = "      "

            #region Retrieving Files Information
            Write-Host
            Write-Host ("   {0} - Retrieving level 1 files: {1}" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"), $dirL1.FullName)
            $dirFiles = New-Object System.Collections.Queue;
            try {
                Get-ChildItem -Path $dirL1.FullName -File -Force:$getHidden -ErrorAction Stop | ForEach-Object{$dirFiles.Enqueue($_)}
            }
            catch {
                Write-CatErrorCSV -funcCSVItem $fileShare -funcAction "Retrieving level 1 files" -funcException $_.Exception -funcDetails $_.ErrorDetails -funcStackTrace $_.ScriptStackTrace -funcFilePath $csvErr
            }
            Get-CatFileInfo -files $dirFiles -sUNC $fileShare.serverUNC -sName $fileShare.shareName -iFile $csvFile -pFile $csvPerm -spacing $outSpacing -l1Color "DarkBlue" -l2Color "DarkCyan"
            Remove-Variable dirFiles
            #endregion

            #region Retrieving Directory Information
            Write-Host ("   {0} - Retrieving level 2 directories: {1}" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"), $dirL1.FullName)
            $dirLevel2Folders = New-Object System.Collections.Queue;
            try {
                Get-ChildItem -Path $dirL1.FullName -Directory -Force:$getHidden -ErrorAction Stop | ForEach-Object{$dirLevel2Folders.Enqueue($_)}
            }
            catch {
                Write-CatErrorCSV -funcCSVItem $fileShare -funcAction "Retrieving level 2 directories" -funcException $_.Exception -funcDetails $_.ErrorDetails -funcStackTrace $_.ScriptStackTrace -funcFilePath $csvErr
            }
            Get-CatDirectoryInfo -directories $dirLevel2Folders -sUNC $fileShare.serverUNC -sName $fileShare.shareName -iFile $csvDir -pFile $csvPerm -spacing $outSpacing -l1Color "DarkBlue" -l2Color "DarkCyan"
            #endregion
            #endregion
            
            :loopL2 while ($dirLevel2Folders.Count -ne 0) {
                $dirLevel2Folders.Dequeue() | ForEach-Object {
                    #region 10.  Level 2 Directories & Files
                    $dirL2 = $_
                    $outSpacing = "         "

                    #region Retrieving Files Information
                    Write-Host
                    Write-Host ("      {0} - Retrieving level 2 files: {1}" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"), $dirL2.FullName) -ForegroundColor DarkBlue                    
                    $dirFiles = New-Object System.Collections.Queue;
                    try {
                        Get-ChildItem -Path $dirL2.FullName -File -Force:$getHidden -ErrorAction Stop | ForEach-Object{$dirFiles.Enqueue($_)}
                    }
                    catch {
                        Write-CatErrorCSV -funcCSVItem $fileShare -funcAction "Retrieving level 2 files" -funcException $_.Exception -funcDetails $_.ErrorDetails -funcStackTrace $_.ScriptStackTrace -funcFilePath $csvErr
                    }
                    Get-CatFileInfo -files $dirFiles -sUNC $fileShare.serverUNC -sName $fileShare.shareName -iFile $csvFile -pFile $csvPerm -spacing $outSpacing -l1Color "DarkCyan" -l2Color "DarkGray"
                    Remove-Variable dirFiles
                    #endregion

                    #region Retrieving Directory Information
                    Write-Host ("      {0} - Retrieving level 3 directories: {1}" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"), $dirL2.FullName) -ForegroundColor DarkBlue
                    $dirLevel3Folders = New-Object System.Collections.Queue;
                    try {
                        Get-ChildItem -Path $dirL2.FullName -Directory -Force:$getHidden -ErrorAction Stop | ForEach-Object{$dirLevel3Folders.Enqueue($_)}
                    }
                    catch {
                        Write-CatErrorCSV -funcCSVItem $fileShare -funcAction "Retrieving level 3 directories" -funcException $_.Exception -funcDetails $_.ErrorDetails -funcStackTrace $_.ScriptStackTrace -funcFilePath $csvErr
                    }
                    Get-CatDirectoryInfo -directories $dirLevel3Folders -sUNC $fileShare.serverUNC -sName $fileShare.shareName -iFile $csvDir -pFile $csvPerm -spacing $outSpacing -l1Color "DarkCyan" -l2Color "DarkGray"
                    #endregion
                    #endregion

                    :loopL3 while ($dirLevel3Folders.Count -ne 0) {
                        $dirLevel3Folders.Dequeue() | ForEach-Object {
                            #region 11.  Remaining Directories & Files
                            $dirL3 = $_
                            $outSpacing = "            "

                            #region Retrieving Files Information
                            Write-Host
                            Write-Host ("         {0} - Retrieving remaining files: {1}" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"), $dirL3.FullName) -ForegroundColor DarkCyan
                            $dirFiles = New-Object System.Collections.Queue;
                            try {
                                Get-ChildItem -Path $dirL3.FullName -File -Force:$getHidden -Recurse -ErrorAction Stop | ForEach-Object{$dirFiles.Enqueue($_)}
                            }
                            catch {
                                Write-CatErrorCSV -funcCSVItem $fileShare -funcAction "Retrieving remaining files" -funcException $_.Exception -funcDetails $_.ErrorDetails -funcStackTrace $_.ScriptStackTrace -funcFilePath $csvErr
                            }
                            Get-CatFileInfo -files $dirFiles -sUNC $fileShare.serverUNC -sName $fileShare.shareName -iFile $csvFile -pFile $csvPerm -spacing $outSpacing -l1Color "DarkGray" -l2Color "DarkMagenta"
                            Remove-Variable dirFiles
                            #endregion

                            #region Retrieving Directory Information
                            Write-Host ("         {0} - Retrieving remaining directories: {1}" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"), $dirL3.FullName) -ForegroundColor DarkCyan
                            $dirRemainingFolders = New-Object System.Collections.Queue;
                            try {
                                Get-ChildItem -Path $dirL3.FullName -Directory -Force:$getHidden -Recurse -ErrorAction Stop | ForEach-Object{$dirRemainingFolders.Enqueue($_)}
                            }
                            catch {
                                Write-CatErrorCSV -funcCSVItem $fileShare -funcAction "Retrieving remaining directories" -funcException $_.Exception -funcDetails $_.ErrorDetails -funcStackTrace $_.ScriptStackTrace -funcFilePath $csvErr
                            }
                            Get-CatDirectoryInfo -directories $dirRemainingFolders -sUNC $fileShare.serverUNC -sName $fileShare.shareName -iFile $csvDir -pFile $csvPerm -spacing $outSpacing -l1Color "DarkGray" -l2Color "DarkMagenta"
                            #endregion
                            #endregion

                            Remove-Variable dirL3
                            Remove-Variable dirRemainingFolders
                        }
                    }
                    Remove-Variable dirL2
                    Remove-Variable dirLevel3Folders
                }
            }
            Remove-Variable dirL1
            Remove-Variable dirLevel2Folders
        }
    } 
    Remove-Variable dirLevel1Folders
}

Remove-Variable outSpacing
Remove-Variable outputFileInfo
Remove-Variable outputDirInfo
Remove-Variable outputPermInfo
Remove-Variable outputScanInfoFolder
Remove-Variable outputErrInfo
#endregion

#region Completed Script
# Output Script Info
Complete-CatScript $scriptStart[0]

# Stop logging
Stop-Transcript
#endregion