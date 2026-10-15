<#  Cat.SP13-CreateColumns.ps1

This script iterates through a CSV File to create columns in a SharePoint 2013 site collection.

Dependencies
	-	SharePoint 2013 Module
	-	Input CSV File in Input Files folder - inputFile - "Cat.PerformAction.Test.csv"

Call script with parameters
	.\Cat.SP13-CreateColumns.ps1 -inputFile Cat.CreateColumns-Test.csv -siteURL https://test.company.com/sites/test

CSV Example
	displayName,internalName,groupName,type,diplayFormat,Description
	Test Text 01,TestText01,Test Columns,Text,,This is a sample text column

Table of Contents
	1.	Input & Output Files
	2.	Variables
	3.	Snapin(s), Modules & Environment Connection(s)
	4.	Main Script
#>

#region Script Parameters
Param(	[String] $inputFile = "Cat.SP13-CreateColumns-Test.csv",
    	[String] $siteURL = "https://devportal.azdps.gov/sites/cthub"  
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
Start-Transcript -Path ("..\_Log FIles\Cat.SP13-CreateColumns-{0}.log" -f $scriptStart[1])
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
# Script Variables
[XML] $defaultXML = '<Field Type="%Type%" Name="%Name%" Description="%Description%" DisplayName="%DisplayName%" StaticName="%StaticName%" Group="%GroupName%" Hidden="FALSE" Required="FALSE" Sealed="FALSE" ShowInDisplayForm="TRUE" ShowInEditForm="TRUE" ShowInListSettings="TRUE" ShowInNewForm="TRUE"></Field>'

# Output Parameters & Variables
Write-Host
Write-Host "Script Parameters" -ForegroundColor Cyan
Write-Host ("   Input File: {0}" -f $inputFile)
Write-Host ("   Site URL: {0}" -f $siteURL)
Write-Host
Write-Host "Default XML" -ForegroundColor Cyan -NoNewline
$defaultXML.Field
#endregion

#region 3.	Snapin(s), Modules & Environment Connection(s)
# Load PowerShell Snap-In(s)
Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue
#endregion
    
#region 4.	Main Script
Write-Host
Write-Host "Configuring Columns" -ForegroundColor Magenta
Write-Host ("   Retrieving Site: {0}" -f $siteURL)
Write-Host
$site = Get-SPSite $siteURL
# Loop Through CSV File
foreach ($column in $inFile) {
	# Update XML
	Write-Host ("   Creating Column: {0}" -f $column.displayName)
	[XML]$columnXML = $defaultXML
	$columnXML.field.type = $column.type
	$columnXML.field.Name = $column.internalName
	$columnXML.field.StaticName = $column.internalName
	$columnXML.field.DisplayName = $column.displayName
	$columnXML.field.Description = $column.description
	$columnXML.field.Group = $column.groupName
	#$columnXML.field # Use for testing values

	# Creating Column
	$temp = $site.rootweb.Fields.AddFieldAsXml($columnXML.OuterXml)
}
#endregion

#region Completed Script
# Stop logging
Stop-Transcript

# Output Script Info
Complete-CatScript $scriptStart[0]
#endregion