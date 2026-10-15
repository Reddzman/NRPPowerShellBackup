<# Cat.Files-NetworkShareDiscovery.ps1

This script scans a list of network shares from CSV and ouputs information about the folders and files. The shares should be broken
up into management chunks.  This is a long running script and it is helpful to be able to pick up where you left off.

Dependencies
    -	Scripts Solution Folder structure
    -   ..\_Input Files - Input file must exist here
    -   ..\_Output Files - Output file will be written here

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
	1.  Functions
    2.  Get-CatDirectoryInfo
    3.  Get-CatFileInfo
    4.	Start Date/Time
    5.  Input & Output Files
	6.	Variables
	7.	Main Script
    8.  Creating Directories
    9.  Share Root Directories & Files
    10.  Level 1 Directories & Files
    11.  Level 2 Directories & Files
    12.  Remaining Directories & Files
#>

#region Script Parameters
Param(
        [string] $inputFile = "Cat.Files-NetworkShareDiscovery-Test.csv",
        [String] $client = "{Client}",
        [boolean] $getHidden = $true
)
#endregion

#region 1.  Functions
    #region General Functions
    <# Start-CatScript

    This function returns the script start time as an array.  The [0] member of the array is normal date time
    format for use in the script and the [1] member is formated for file names.

    Sample Call:
        $scriptStart = Start-CatScript

    #>
    function Start-CatScript{
        # Collecting Start Time
        $SDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        $FDate = Get-Date -Format "yyMMddHHmm"
        
        # Return Start Time
        $DateArray = @()
        $DateArray += $SDate
        $DateArray += $FDate
        return $DateArray
    } # Start-CatScript

    <# Complete-CatScript

    This function uses the start time array to calculate and output duration at the end of a script.

    Sample Call: 
        Complete-CatScript $scriptStart[0]

    #>
    function Complete-CatScript{
        param(
            [parameter(Mandatory=$true, Position=0, ValueFromPipeline=$true)][datetime]$STime
        )
        
        # Finish Time & Time Taken
        [datetime]$FTime = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        $TTime = New-TimeSpan -Start $STime -End $FTime

        # Screen Output
        Write-Host
        Write-Host "-----------------------------------------------------------------------------------------------------" -ForegroundColor Cyan
        Write-Host ("Started: {0}   Finished: {1}   Time Taken: {2}" -f $STime, $FTime, $TTime) -ForegroundColor Cyan
        Write-Host "-----------------------------------------------------------------------------------------------------" -ForegroundColor Cyan
        Write-Host

        # Stopping Transcript
        #Stop-Transcript
    } # Complete-CatScript
    #endregion

    #region 2.  Get-CatDirectoryInfo
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
            [System.Object] $directories,
            [string] $sUNC,
            [string] $sName,
            [string] $iFile,
            [string] $pFile,
            [string] $spacing = "   ",
            [string] $l1Color = "White",
            [string] $l2Color = "DarkGray"
        )

        if ($directories.Count -gt 0) {
            # Adding Properties to Directory Info
            Write-Host ("{0}{1} - Adding share properites to directories" -f $spacing, (Get-Date -Format "yy-MM-dd HH:mm:ss")) -ForegroundColor $l1Color
            $directories | ForEach-Object {$_ | Add-Member -MemberType NoteProperty -Name "ServerUNC" -Value $sUNC}
            $directories | ForEach-Object {$_ | Add-Member -MemberType NoteProperty -Name "ShareName" -Value $sName}

            # Exporting to CSV Directory Information
            Write-Host ("{0}{1} - Exporting directories to CSV" -f $spacing, (Get-Date -Format "yy-MM-dd HH:mm:ss")) -ForegroundColor $l1Color
            $directories | Select-Object -Property ServerUNC,ShareName,Name,Parent,Root,BaseName,FullName,LastAccessTime,LastAccessTimeUtc,LastWriteTime,LastWriteTimeUtc,LinkType,`
                Mode,ModeWithoutHardLink,Target,PSChildName,PSIsContainer,PSParentPath,PSPath,PSProvider,Attributes,CreationTime,CreationTimeUtc,Exists,Extension  | Export-Csv -Path $iFile -UseQuotes AsNeeded -NoTypeInformation -Append

            # Retrieving Directory Permissions Information
            Write-Host ("{0}{1} - Retrieving directory permission" -f $spacing, (Get-Date -Format "yy-MM-dd HH:mm:ss")) -ForegroundColor $l1Color
            $acls = @()
            $acls += $directories | ForEach-Object {$_ | Get-Acl -ErrorAction Continue -ErrorVariable errLog}

            # Adding Properties to Directory Permissions
            Write-Host ("{0}   {1} - Adding share properites to directory permissions" -f $spacing, (Get-Date -Format "yy-MM-dd HH:mm:ss")) -ForegroundColor $l2Color
            $acls | ForEach-Object {$_ | Add-Member -MemberType NoteProperty -Name "ServerUNC" -Value $sUNC}
            $acls | ForEach-Object {$_ | Add-Member -MemberType NoteProperty -Name "ShareName" -Value $sName}

            # Exporting to CSV Directory Permissions
            Write-Host ("{0}   {1} - Exporting directory permissions to CSV" -f $spacing, (Get-Date -Format "yy-MM-dd HH:mm:ss")) -ForegroundColor $l2Color
            $acls | Select-Object -Property ServerUNC,ShareName,Owner,Group,CentralAccessPolicyId,Path,Sddl,PSChildName,PSParentPath,PSPath,PSProvider,AreAccessRulesCanonical,AreAccessRulesProtected,`
                AreAuditRulesCanonical,AreAuditRulesProtected -ExpandProperty Access | Export-Csv -Path $pFile -UseQuotes AsNeeded -NoTypeInformation -Append
        }
        else {
            Write-Host ("{0}{1} - No directories found" -f $spacing, (Get-Date -Format "yy-MM-dd HH:mm:ss")) -ForegroundColor $l1Color
        }
    }
    #endregion

    #region 3.  Get-CatFileInfo
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
            [System.Object] $files,
            [string] $sUNC,
            [string] $sName,
            [string] $iFile,
            [string] $pFile,
            [string] $spacing = "   ",
            [string] $l1Color = "White",
            [string] $l2Color = "DarkGray"
        )

        if ($files.Count -gt 0) {
            # Adding Properties to File Info
            Write-Host ("{0}{1} - Adding share properites to files" -f $spacing, (Get-Date -Format "yy-MM-dd HH:mm:ss")) -ForegroundColor $l1Color
            $files | ForEach-Object {$_ | Add-Member -MemberType NoteProperty -Name "ServerUNC" -Value $sUNC}
            $files | ForEach-Object {$_ | Add-Member -MemberType NoteProperty -Name "ShareName" -Value $sName}

            # Exporting to CSV File Information
            Write-Host ("{0}{1} - Exporting files to CSV" -f $spacing, (Get-Date -Format "yy-MM-dd HH:mm:ss")) -ForegroundColor $l1Color
            $files | Select-Object -Property ServerUNC,ShareName,Name,BaseName,LinkType,Mode,ModeWithoutHardLink,Target,PSChildName,PSIsContainer,PSParentPath,PSPath,PSProvider,Attributes,`
                CreationTime,CreationTimeUtc,Directory,DirectoryName,Exists,Extension,FullName,IsReadOnly,LastAccessTime,LastAccessTimeUtc,LastWriteTime,LastWriteTimeUtc,Length `
                -ExpandProperty VersionInfo | Export-Csv -Path $iFile -UseQuotes AsNeeded -NoTypeInformation -Append
            
            # Retrieving File Permissions Information
            Write-Host ("{0}{1} - Retrieving file permission" -f $spacing, (Get-Date -Format "yy-MM-dd HH:mm:ss")) -ForegroundColor $l1Color
            $acls = @()
            $acls += $files | ForEach-Object {$_ | Get-Acl -ErrorAction Continue -ErrorVariable errLog}

            # Adding Properties to File Permissions
            Write-Host ("{0}   {1} - Adding share properites to file permissions" -f $spacing, (Get-Date -Format "yy-MM-dd HH:mm:ss")) -ForegroundColor $l2Color
            $acls | ForEach-Object {$_ | Add-Member -MemberType NoteProperty -Name "ServerUNC" -Value $sUNC}
            $acls | ForEach-Object {$_ | Add-Member -MemberType NoteProperty -Name "ShareName" -Value $sName}

            # Exporting to CSV File Permissions
            Write-Host ("{0}   {1} - Exporting file permissions to CSV" -f $spacing, (Get-Date -Format "yy-MM-dd HH:mm:ss")) -ForegroundColor $l2Color
            $acls | Select-Object -Property ServerUNC,ShareName,Owner,Group,CentralAccessPolicyId,Path,Sddl,PSChildName,PSParentPath,PSPath,PSProvider,AreAccessRulesCanonical,AreAccessRulesProtected,`
                AreAuditRulesCanonical,AreAuditRulesProtected -ExpandProperty Access | Export-Csv -Path $pFile -UseQuotes AsNeeded -NoTypeInformation -Append

        }
        else {
            Write-Host ("{0}{1} - No files found" -f $spacing, (Get-Date -Format "yy-MM-dd HH:mm:ss")) -ForegroundColor $l1Color
        }
    }
    #endregion
#endregion

#region 4.  Start Date / Time
# Get Start Date / Time
$global:scriptStart = Start-CatScript

# Start Transcript
Start-Transcript -Path ("..\_Log Files\{0}.Files-NetworkShareDiscovery-{1}.log" -f $client, $scriptStart[1])
#endregion

#region 5.	Input & Output Files
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

#region 6.	Variables
# Script Variables
$outputBasePath = ("..\_Output Files\{0}\{1}" -f $client, $scriptStart[1])
$outputScanInfoFolder = ("{0}\details" -f $outputBasePath)
$outputScanInfo = ("{0}\ScanInformation.csv" -f $outputScanInfoFolder)
$outputFileInfo = ("{0}\files" -f $outputBasePath)
$outputDirInfo = ("{0}\directories" -f $outputBasePath)
$outputPermInfo = ("{0}\permissions" -f $outputBasePath)
$errLog = @()

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

#region 7.	Main Script
#region 8.  Creating Directories
Write-Host
Write-Host ("Scanning File Directories: {0}" -f $InPath) -ForegroundColor Cyan

# Creating Directories
Write-Host "   Creating Directories"
New-Item -Path $outputFileInfo -ItemType Directory
New-Item -Path $outputDirInfo -ItemType Directory
New-Item -Path $outputPermInfo -ItemType Directory
New-Item -Path $outputScanInfoFolder -ItemType Directory

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

    #region 9.  Share Root Directories & Files
    # Retrieving Share Permissions Information
    $sharePath = ("\\{0}\{1}" -f $fileShare.serverUNC, $fileShare.shareName)
    Write-Host ("   {0} - Retrieving directory permission" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"))
    $acl = Get-ACL -Path $sharePath -ErrorAction Continue -ErrorVariable errLog

    # Adding Properties to Share Permissions
    Write-Host ("   {0} - Adding share properites to directory permissions" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"))
    $acl | Add-Member -MemberType NoteProperty -Name "ServerUNC" -Value $fileShare.serverUNC
    $acl | Add-Member -MemberType NoteProperty -Name "ShareName" -Value $fileShare.shareName

    # Exporting to CSV Share Permissions
    Write-Host ("   {0} - Exporting directory permissions to CSV" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"))
    $acl | Select-Object -Property ServerUNC,ShareName,Owner,Group,CentralAccessPolicyId,Path,Sddl,PSChildName,PSParentPath,PSPath,PSProvider,AreAccessRulesCanonical,AreAccessRulesProtected,`
        AreAuditRulesCanonical,AreAuditRulesProtected -ExpandProperty Access | Export-Csv -Path $csvPerm -UseQuotes AsNeeded -NoTypeInformation -Append
    
    # Retrieving Files Information
    Write-Host ("   {0} - Retrieving root files" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"))
    $outSpacing = "   "
    $dirFiles = @()
    $dirFiles += Get-ChildItem -Path $sharePath -File -Force:$getHidden -ErrorAction Continue -ErrorVariable errLog
    Get-CatFileInfo -files $dirFiles -sUNC $fileShare.serverUNC -sName $fileShare.shareName -iFile $csvFile -pFile $csvPerm

    # Retrieving Directory Information
    Write-Host ("   {0} - Retrieving level 1 directories" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"))
    $outSpacing = "   "
    $dirLevel1Folders = @()
    $dirLevel1Folders += Get-ChildItem -Path $sharePath -Directory -Force:$getHidden -ErrorAction Continue -ErrorVariable errLog
    Get-CatDirectoryInfo -directories $dirLevel1Folders -sUNC $fileShare.serverUNC -sName $fileShare.shareName -iFile $csvDir -pFile $csvPerm
    #endregion

    :loopL1 foreach ($dirL1 in $dirLevel1Folders) {
        #region 10.  Level 1 Directories & Files        
        # Retrieving Files Information
        Write-Host
        Write-Host ("   {0} - Retrieving level 1 files: {1}" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"), $dirL1.FullName)
        $outSpacing = "      "
        $dirFiles = @()
        $dirFiles += Get-ChildItem -Path $dirL1.FullName -File -Force:$getHidden -ErrorAction Continue -ErrorVariable errLog
        Get-CatFileInfo -files $dirFiles -sUNC $fileShare.serverUNC -sName $fileShare.shareName -iFile $csvFile -pFile $csvPerm -spacing $outSpacing -l1Color "DarkGray" -l2Color "DarkBlue"

        # Retrieving Directory Information
        Write-Host ("   {0} - Retrieving level 2 directories: {1}" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"), $dirL1.FullName)
        $outSpacing = "      "
        $dirLevel2Folders = @()
        $dirLevel2Folders = Get-ChildItem -Path $dirL1.FullName -Directory -Force:$getHidden -ErrorAction Continue -ErrorVariable errLog
        Get-CatDirectoryInfo -directories $dirLevel2Folders -sUNC $fileShare.serverUNC -sName $fileShare.shareName -iFile $csvDir -pFile $csvPerm -spacing $outSpacing -l1Color "DarkGray" -l2Color "DarkBlue"
        #endregion

        :loopL2 foreach ($dirL2 in $dirLevel2Folders) {
            #region 11.  Level 2 Directories & Files
            # Retrieving Files Information
            Write-Host
            Write-Host ("      {0} - Retrieving level 2 files: {1}" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"), $dirL2.FullName) -ForegroundColor DarkGray
            $outSpacing = "         "
            $dirFiles = @()
            $dirFiles += Get-ChildItem -Path $dirL2.FullName -File -Force:$getHidden -ErrorAction Continue -ErrorVariable errLog
            Get-CatFileInfo -files $dirFiles -sUNC $fileShare.serverUNC -sName $fileShare.shareName -iFile $csvFile -pFile $csvPerm -spacing $outSpacing -l1Color "DarkBlue" -l2Color "DarkCyan"

            # Retrieving Directory Information
            Write-Host ("      {0} - Retrieving level 3 directories: {1}" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"), $dirL2.FullName) -ForegroundColor DarkGray
            $outSpacing = "         "
            $dirLevel3Folders = @()
            $dirLevel3Folders = Get-ChildItem -Path $dirL2.FullName -Directory -Force:$getHidden -ErrorAction Continue -ErrorVariable errLog
            Get-CatDirectoryInfo -directories $dirLevel3Folders -sUNC $fileShare.serverUNC -sName $fileShare.shareName -iFile $csvDir -pFile $csvPerm -spacing $outSpacing -l1Color "DarkBlue" -l2Color "DarkCyan"
            #endregion

            :loopL3 foreach ($dirL3 in $dirLevel3Folders) {
                #region 12.  Remaining Directories & Files
                # Retrieving Files Information
                Write-Host
                Write-Host ("         {0} - Retrieving remaining files: {1}" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"), $dirL3.FullName) -ForegroundColor DarkBlue
                $outSpacing = "            "
                $dirFiles = @()
                $dirFiles += Get-ChildItem -Path $dirL3.FullName -File -Force:$getHidden -Recurse -ErrorAction Continue -ErrorVariable errLog
                Get-CatFileInfo -files $dirFiles -sUNC $fileShare.serverUNC -sName $fileShare.shareName -iFile $csvFile -pFile $csvPerm -spacing $outSpacing -l1Color "DarkCyan" -l2Color "Gray"

                # Retrieving Directory Information
                Write-Host ("         {0} - Retrieving remaining directories: {1}" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"), $dirL3.FullName) -ForegroundColor DarkBlue
                $outSpacing = "            "
                $dirRemainingFolders = @()
                $dirRemainingFolders = Get-ChildItem -Path $dirL3.FullName -Directory -Force:$getHidden -Recurse -ErrorAction Continue -ErrorVariable errLog
                Get-CatDirectoryInfo -directories $dirRemainingFolders -sUNC $fileShare.serverUNC -sName $fileShare.shareName -iFile $csvDir -pFile $csvPerm -spacing $outSpacing -l1Color "DarkCyan" -l2Color "Gray"
                #endregion
            }
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