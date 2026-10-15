<#  Cat.SP16-WebParts.ps1 Description

This script scans a site collection and looks for the user of web parts of the type listed in the $desiredWPTypes 
variable. It was orgionally created to scan an intranet looking for Content Editor and Script editor web parts.

Tested on SP16 & SP19

Dependencies
	-	SharePoint PowerShell Modules
	-	General Functions - ..\General\Cat.GeneralFunctions.psm1

Call script with parameters
	cd C:\GitHub\CatPowerShell\Reporting
	.\Cat.SP16-WebParts.ps1  -SiteURL "http://portal.sp19test.local" -pageLibrary "Pages"

Table of Contents
	1.	Snapin(s), Modules & Environment Connection(s)
	2.	Input & Output Files
	3.	Variables
	4.	Main Script
	5.	CSV Loop Example
	6.	...
#>

#region Script Parameters
Param(
        [string] $SiteURL = "http://your-sharepoint-site", # Site collection you want to scan
        [string] $pageLibrary = "Pages" # Pages or SitePages
)
#endregion Script Parameters

#region Start Date / Time
# Version Update Check
$env:PNPPOWERSHELL_UPDATECHECK="false"

# Remove Existing Catapult Modules to ensure PS session has most up to date functions
Get-Module Cat.* | Remove-Module

# Load General Functions
Import-Module "..\General\Cat.GeneralFunctions.psm1"

# Get Start Date / Time
$scriptStart = Start-CatScript

# Start Transcript
Start-Transcript -Path ("..\_Log Files\SP16-WebParts-{0}.log" -f $scriptStart[1])
#endregion Start Date / Time

#region 2.	Input & Output Files
# File Paths
$ReportOutput = ("..\_Output Files\SP16-WebParts-{0}.csv" -f $scriptStart[1])
#endregion Input & Output Files

#region 3.	Variables
$desiredWPTypes = @("ContentEditorWebPart";"ScriptEditorWebPart";"PageViewerWebPart";)
$ResultCollection = @()

#region Output Parameters & Variables
Write-Host
Write-Host "Script Parameters" -ForegroundColor Cyan
Write-Host ("   Site URL: {0}" -f $SiteURL)
Write-Host
Write-Host "Output Variables" -ForegroundColor Cyan
Write-Host ("   Report: {0}" -f $ReportOutput)
Write-Host
Write-Host "Web Part Types" -ForegroundColor Cyan
Write-host $desiredWPTypes
#endregion Output Parameters & Variables
#endregion
    
#region 4.	Main Script
# Loop through all webs in site collection
Write-Host
Write-Host "Processing Site: $($SiteURL)" -ForegroundColor Cyan
$Site = Get-SPSite $SiteURL
foreach ($Web in $Site.AllWebs) {
	Write-Host
    Write-Host ("   {0} - Processing Web: {1}" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"),$Web.Url) -ForegroundColor Magenta

	#for each web, loop through each page
    foreach ($Page in $Web.GetFolder($pageLibrary).Files) {
		Write-Host ("      Processing Page: {0}" -f $Page.ServerRelativeUrl)
        $PageUrl = $Page.ServerRelativeUrl
        $File = $Web.GetFile($PageUrl)
        if ($File.Exists) {
			Write-Host
            $Manager = $File.GetLimitedWebPartManager([System.Web.UI.WebControls.WebParts.PersonalizationScope]::Shared)

			#For each page, loop through each Web Part
            foreach ($WebPart in $Manager.WebParts) {
				if ($WebPart.GetType().Name -in $desiredWPTypes){
					Write-Host ("         Processing Web Part: {0} - {1}" -f $WebPart.Title,$WebPart.GetType().Name) -ForegroundColor DarkGray

					# Creating output object
					$Result = New-Object PSObject
					$Result | Add-Member -Type NoteProperty -Name "Site URL" -Value $Web.Url
					$Result | Add-Member -Type NoteProperty -Name "Page URL" -Value $PageUrl
					$Result | Add-Member -Type NoteProperty -Name "Web Part Title" -Value $WebPart.Title
					$Result | Add-Member -Type NoteProperty -Name "Web Part Type" -Value $WebPart.GetType().Name
					switch ($WebPart.GetType().Name) {
						ContentEditorWebPart {$Result | Add-Member -Type NoteProperty -Name "Content" -Value (($WebPart.Content.InnerText -replace "`n"," ") -replace "`r"," ")}
						ScriptEditorWebPart {$Result | Add-Member -Type NoteProperty -Name "Content" -Value (($WebPart.Content.Substring(0, [Math]::Min(100, $WebPart.Content.Length)) -replace "`n"," ") -replace "`r"," ")}
						PageViewerWebPart {$Result | Add-Member -Type NoteProperty -Name "Content" -Value $WebPart.ContentLink}
					}
					$ResultCollection += $Result
				}
            }
        }
		else {Write-Host "File not Found" -ForegroundColor Yellow}
    }
}

# Export results
$ResultCollection | Export-Csv -Path $ReportOutput -NoTypeInformation
Write-Host
Write-Host "Scan complete. Results saved to $ReportOutput" -ForegroundColor Green
#endregion Main Script

#region Completed Script
# Output Script Info
Complete-CatScript $scriptStart[0]

# Stop logging
Stop-Transcript
#endregion Completed Script