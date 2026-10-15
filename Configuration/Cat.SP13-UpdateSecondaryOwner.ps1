<#  Cat.SP13-UpdateSecondaryOwner.ps1

This script iterates through all site collections in SharePoint 2013 and updates the secondary administrators for every site collection in the Farm.
This is similar to setting primary and secondary site collection administrators through Central Administration.

Dependencies
	-	SharePoint Server Modules

Call script with parameters
	cd C:\GitHub\CatPowerShell\Configuration
	.\Cat.SP13-UpdateSecondaryOwner.ps1 -client "nrpdev2" -secondaryOwner "sp13test\sharegate" -includeCentralAdmin $false -reportOnly $false -skipWebAppsURLs @("https://mysites.sp13test.local/";"https://other.client.com/";)

Table of Contents
	1.	Input & Output Files
	2.	Variables
	3.	Snapin(s), Modules & Environment Connection(s)
	4.	Main Script
#>

#region Script Parameters
Param(	
		[string] $client = "{client}",
		[string] $secondaryOwner = "domain\username",
		[boolean] $includeCentralAdmin = $false,
		[boolean] $reportOnly = $false,
		[string] $outputLocation = "c:\GitHub\CatPowerShell\_Output Files",
		$skipWebAppsURLs = @("https://my.client.com";"https://other.client.com";)

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
Start-Transcript -Path ("..\_Log FIles\{0}.SP13-UpdateSecondaryOwner-{1}.log" -f $client, $scriptStart[1])
#endregion

#region 2.	Variables
# File Paths
$outputPath = ("{0}\{1}\{2}" -f $outputLocation, $client, $scriptStart[1])
$errorPath = ("{0}\errors.csv" -f $outputPath)
$outputSites = ("{0}\AdminReport.csv" -f $outputPath)

# Output Parameters & Variables
Write-Host
Write-Host "Script Parameters" -ForegroundColor Cyan
Write-Host ("   Client: {0}" -f $client)
Write-Host ("   Secondary Owner: {0}" -f $secondaryOwner)
Write-Host ("   Include Central Admin: {0}" -f $includeCentralAdmin)
Write-Host ("   Report Only: {0}" -f $reportOnly)
Write-Host
Write-Host "File Paths" -ForegroundColor Cyan
Write-Host ("   Output Location: {0}" -f $outputPath)
Write-Host ("   Error File: {0}" -f $errorPath)
Write-Host ("   Admin Report: {0}" -f $outputSites)
Write-Host
Write-Host "Skip Web Applications" -ForegroundColor Cyan
$skipWebAppsURLs
#endregion

#region 3.	Snapin(s), Modules & Environment Connection(s)
# Load MS / PnP PowerShell Modules
Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue
#endregion
    
#region 4.	Main Script
#region Created directories for additional reports
Write-Host
Write-Host ("{0} - Creating Report Directory" -f (Get-Date -Format "yy-MM-dd HH:mm:ss")) -ForegroundColor Cyan
New-Item -Path $outputPath -ItemType Directory
#endregion

# Screen Output
Write-Host
if ($reportOnly) {
	Write-Host ("{0} - Reporting current secondary admin: {1}" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"), $secondaryOwner) -ForegroundColor Cyan
} else {
	Write-Host ("{0} - Configuring secondary admin" -f (Get-Date -Format "yy-MM-dd HH:mm:ss")) -ForegroundColor Cyan
}

#region Get Web Application Objects
Write-Host ("   {0} - Retrieving web applciations" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"))
try {
	$webApps = Get-SPWebApplication -IncludeCentralAdministration:$includeCentralAdmin -ErrorAction Stop
}
catch {
	$obj = New-Object -TypeName PSObject -Property(@{`
		siteTitle=$null;`
		siteURL=$null;`
		primaryAdminName=$null;`
		primaryAdminAccount=$null;`
		secondaryAdminName=$null;`
		secondaryAdminAccount=$null;`
		ID=$null;})
	Write-CatErrorCSV -funcCSVItem $obj -funcAction "Retrieving web applciations" -funcException $_.Exception -funcDetails $_.ErrorDetails -funcStackTrace $_.ScriptStackTrace -funcFilePath $errorPath
}
#endregion

# Loop through all web applications in the farm
:webAppLoop ForEach ($webApp in $webApps)
{
	# Screen Output for web application
	Write-Host
	Write-Host ("   {0} - Web application: {1}" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"),$webApp.URL)  -ForegroundColor Magenta

	if ($webApp.URL -notin $skipWebAppsURLs) {
		#region Get site collection objects
		Write-Host ("      {0} - Retrieving site collections" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"))
		try {
			$sites = Get-SPSite -WebApplication $WebApp.URL -limit all -ErrorAction Stop
		}
		catch {
			$obj = New-Object -TypeName PSObject -Property(@{`
				siteTitle=$null;`
				siteURL=$null;`
				primaryAdminName=$null;`
				primaryAdminAccount=$null;`
				secondaryAdminName=$null;`
				secondaryAdminAccount=$null;`
				ID=$null;})
			Write-CatErrorCSV -funcCSVItem $obj -funcAction "Retrieving site collections" -funcException $_.Exception -funcDetails $_.ErrorDetails -funcStackTrace $_.ScriptStackTrace -funcFilePath $errorPath
			Continue webAppLoop
		}
		#endregion

		# Loop through site collections for web application
		:sitesLoop ForEach ($site in $sites)
		{
			# Output Admin Information
			Write-Host ("      {0} - Writing admin information: {1}" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"),$site.URL)
			$obj = New-Object -TypeName PSObject -Property(@{`
				siteTitle=$site.RootWeb.Title;`
				siteURL=$site.Url;`
				primaryAdminName=$site.Owner.DisplayName;`
				primaryAdminAccount=$site.Owner.UserLogin;`
				secondaryAdminName=$site.SecondaryContact.DisplayName;`
				secondaryAdminAccount=$site.SecondaryContact.UserLogin;`
				ID=$site.ID.GUID;})
			$obj | Export-Csv -Path $outputSites -NoTypeInformation -Append

			# Updating Secondary owner
			if (-not $reportOnly) {
				Write-Host ("      {0} - Updating secondary owner: {1}" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"),$site.URL)
				try {
					Set-SPSite -Identity $site.URL -SecondaryOwnerAlias $secondaryOwner -ErrorAction Stop
				}
				catch {
					Write-CatErrorCSV -funcCSVItem $obj -funcAction "Updating secondary owner" -funcException $_.Exception -funcDetails $_.ErrorDetails -funcStackTrace $_.ScriptStackTrace -funcFilePath $errorPath
					Continue sitesLoop
				}
			}

			Remove-Variable obj
		} # ForEach ($Site in $Sites)
		Remove-Variable sites
	}
	else {
		# Get site collection objects
		Write-Host "      Skipping web application" -ForegroundColor Yellow
	}		
} # ForEach ($WebApp in $WebApps)
Remove-Variable webApps
#endregion

#region Completed Script
# Output Script Info
Complete-CatScript $scriptStart[0]

# Stop logging
Stop-Transcript
#endregion