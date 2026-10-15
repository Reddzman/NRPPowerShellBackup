#region Documentation
<#  Cat.PS-BreakUpLargeCSVFiles.ps1

This script can be used to break up large CSV files.

Dependencies
	-	General Functions - ..\General\Cat.GeneralFunctions.psm1
	-	Input CSV File - Example: ..\_Input FIles\Testing\Cat.Tech-Action-Test.csv

Call script with parameters
	cd 'C:\GitHub\CatPowerShell\Sample Data\'
	.\Cat.PS-BreakUpLargeCSVFiles.ps1 -inputFile "CUST-Dir-220216125338.csv" -client "Client" -maxRows 15000

Table of Contents
	1.	Variables
	2.	Main Script
	3.	Count Rows in CSV
	4.	Break Up CSV
#>
#endregion

#region Script Parameters
Param(	[string] $inputFile = "CUST-Dir-220216125338.csv",
		[string] $client = "Client",
		[int] $maxRows = 15000
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
Start-Transcript -Path ("..\_Log Files\{0}.PS-BreakUpLargeCSVFiles-{1}.log" -f $client, $scriptStart[1])
#endregion

#region 1.	Variables
# Script variables
$outputBasePath = ("..\_Output Files\{0}\{1}" -f $client, $inputFile.Replace(".csv",""))
$inPath = ("C:\GitHub\CatPowerShell\_Input Files\{0}" -f $inputFile)
$startRow = 0
$counter = 1
$csvRows = 0

# Output Parameters & Variables
Write-Host
Write-Host "Script Parameters" -ForegroundColor Cyan
Write-Host ("   Input File: {0}" -f $inputFile)
Write-Host ("   Client: {0}" -f $client)
Write-Host ("   Max Rows: {0}" -f $maxRows)
Write-Host
Write-Host "Script Variables" -ForegroundColor Cyan
Write-Host ("   Input Path: {0}" -f $inPath)
Write-Host ("   Output Path: {0}" -f $outputBasePath)
Write-Host ("   Start Row: {0}" -f $startRow)
Write-Host ("   Counter: {0}" -f $counter)
Write-Host ("   CSV Rows: {0}" -f $counter)
#endregion
    
#region 2.	Main Script
#region 3.	Count Rows in CSV
If (Test-Path ($inPath)){
	Write-Host
	Write-Host ("Breaking up CSV File: {0}" -f $inputFile) -ForegroundColor Cyan

	# Creating Directory for Files
	Write-Host ("   {0} - Creating Directory: {1}" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"), $outputBasePath)
	$tempPath = New-Item -Path $outputBasePath -ItemType Directory

	# Retreiving number of rows in file
	Write-Host ("   {0} - Retreiving total number of rows in CSV: {1}" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"), $inputFile) -NoNewline
	$csvReader = New-Object IO.StreamReader $inPath
	while($null -ne $csvReader.ReadLine()){
		$csvRows++
	}
	Write-Host (" - {0}" -f $csvRows)
}
else {
	Write-Host
	throw ("The input file does not exist: {0}" -f $inPath)	
}
#endregion

#region 4.	Break Up CSV
while ($startrow -lt $csvRows)
{
	$outputFile = ("{0}\{1}-{2}.csv" -f $outputBasePath, $inputFile.Replace(".csv",""), $counter)	
	Write-Host ("   {0} - Splitting rows {1} through {2}: {3}" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"), $startRow, ($startRow + $maxRows), $outputFile)
	Import-CSV $inPath | select-object -skip $startrow -first $maxRows | Export-CSV $outputFile -NoClobber;

	# Increment Variables
	$startrow += $maxRows
	$counter++
}
#endregion

#endregion

#region Completed Script
# Output Script Info
Complete-CatScript $scriptStart[0]

# Stop logging
Stop-Transcript
#endregion