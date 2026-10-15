<#
Cat.PnP-ConfigInfoArch.ps1

This script iterates through a list of Web objects and configures content types on all lists in
the web.  It performs the following tasks:
	1.	Enables editing content types on the list
	2.	Adds content type to list configuration
	3.	Configures default values at list level
	4.	Adds any list columns to the content type at a list level
	5.	If configured in XML, Updates list Items for new content type & Adds default value to columns on each item

Dependencies
	-	PnP PowerShell https://msdn.microsoft.com/en-us/pnp_powershell/pnp-powershell-overview
	-	Input CSV File in Input Files folder - Script Parameter - "Cat.IALibraryCnfg.Test.csv"

Call Script with Input File
	.\Cat.PnP-ConfigInfoArch.ps1 -InputFile Cat.IALibraryCnfg-Test.xml

**** XML Example at end of Script ****

Table of Contents
	1.	Input & Output Files
	2.	Variables
	3.	Snapin(s), Modules & Environment Connection(s)
	4.	Main Script
	5.	XML Example
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
$Location = $PSScriptRoot.Replace($LocReplace,"")

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
$OutPath = ("{0}_Output Files\Cat.IALibraryCnfgErrs-Test-{1}.txt" -f $Location, $Start[1])

# Import Input File(s)
[xml]$InFile = get-content $InPath

# Create Output File(s)
Out-File -FilePath $OutPath -InputObject "Site|List|Name|RelativeURL"
#endregion

#region 2.	Variables
# User Inputs
Write-Host
Write-Host "User Input Variables" -ForegroundColor DarkCyan
Write-Host "   Enter credentials for SharePoint sites to be configured:"
$SPCred = Get-Credential -Message "Enter your credentials for SharePoint sites to be configured:"

# Script Variables
#Script will skip non-document libraries, hidden document libraries, picture libraries, form libraries, and the below list of libraries.
$SkipLists = @("Content and Structure Reports";"Customized Reports";"DataConnectionInfoPath";"Drop Off Library";"Form Templates";"Forms Library";"Images";"pages";
				"Site Assets";"Site Collection Images";"Site Pages";"Style Library";"Workflow Tasks";"Style Library";)
#endregion

#region 3.	Snapin(s), Modules & Environment Connection(s)
# Load Custom Modules
Import-Module ("{0}Configuration\Cat.ConfigurationFunctions.psm1" -f $Location)
#endregion

#region 4.	Main Script
# Connect To SPO Admin Site For Term Store
Write-Host
Write-Host ("Configuring Content Types & Default Values" -f $SPAdminURL) -ForegroundColor DarkCyan

# Add Term Groups
# Looping through all sites in input file
ForEach ($Site in $InFile.InfoArchCnfg.Sites.Site)
{
	# Connect to site
	Write-Host
	Write-Host ("   Configuring SharePoint Site: {0}" -f $Site.Name) -ForegroundColor Magenta
	Connect-PnPOnline $Site.URL -Credentials $SPCred
	$PnPCtx = Get-PnPContext

	# Get desired lists / Libraries from site
	$Web = Get-PnPWeb -Includes Lists
	$Lists = $Web.Lists | Where-Object{($_.Hidden -ne $true) -and ($_.Title -notin $SkipLists) -and ($_.BaseType -eq "DocumentLibrary") -and ($_.BaseTemplate -ne "115") -and ($_.BaseTemplate -ne "109")}
	
	# Loop through lists
	ForEach ($List in $Lists)
	{
		# Load List information
		Write-Host ("      Configuring Library: {0}" -f $List.Title) -ForegroundColor DarkGray
		$CurList = Get-PnPList -Identity $List.ID -Includes BaseType, ContentTypes, ContentTypesEnabled, Fields

		# Enable Content Types
		If (-($CurList.ContentTypesEnabled))
		{
			# Configure Library for Content Types
			Write-Host "         Enabling Content Type Management"
			Set-PnPList -Identity $CurList.ID -EnableContentTypes:$true
		}

		# Check for Content Type
		$CT = $CurList.ContentTypes | Where-Object{$_.Name -eq $Site.ContentType}

		If ($CT -eq $null)
		{
			# Add Content Type As Default
			Write-Host ("         Adding Content Type to Library: {0}" -f $Site.ContentType)
			Add-PnPContentTypeToList -List $CurList.ID -ContentType $Site.ContentType -DefaultContentType
		}
		else 
		{
			# Set as Default Content Type
			Write-Host ("         Setting Default Content Type: {0}" -f $Site.ContentType)
			Set-PnPDefaultContentTypeToList -List $CurList.ID -ContentType $Site.ContentType	
		}

		# Get List Content Type and custom list columns
		$CT = $CurList.ContentTypes | Where-Object{$_.Name -eq $Site.ContentType}
		$Columns = $CurList.Fields | Where-Object{($_.Hidden -eq $false) -and ($_.CanBeDeleted -eq $true)}

		# Loop through columns
		ForEach ($Column in $Columns)
		{
			# Adding column to content type
			Write-Host ("            Adding Column to Content Type: {0}" -f $Column.Title)
			Add-ListColumnToListCT -ClientContext $PnPCtx -ContentType $CT -ListColumn $Column
		}

		# Loop through and Set Default Values on list
		$ItemDefaults = @{}
		ForEach ($Default in $Site.Defaults.Default)
		{
			# Set column default values
			Write-Host ("            Setting Default Value: {0}" -f $Default.Column)
			Set-PnPDefaultColumnValues -List $CurList.ID -Field $Default.Column -Value $Default.Value 

			# Building default values for items
			$ItemDefaults += @{("{0}" -f $Default.Column) = ("{0}" -f $Default.Value)}
		}

		# Check if Items should be configured
		If ($Site.ConfigItems -eq "Yes")
		{
			# Screen Output
			Write-Host ("         Looping through items: {0}" -f $List.Title)
			$Items = Get-PnPListItem -List $CurList.ID
			
			# Loop through all list items
			ForEach ($Item in $Items)
			{
				try 
				{
					# Configuring Content type and Adding Column Values
					Write-Host ("            Configuring content type: {0}" -f $Item["FileLeafRef"])
					Set-PnPFileCheckedOut -Url $Item["FileRef"]
					$Temp = Set-PnPListItem -List $CurList.ID -Identity $Item.ID -ContentType $CT -Values $ItemDefaults # Output to temp variable instead of screen
					Set-PnPFileCheckedIn -Url $Item["FileRef"] -CheckinType MajorCheckIn -Comment "Information Architecture: Updating Metadata"

					# Due to a bug in the PnP Commands that resets the content type to document, Reconfiguring just the content type
					Set-PnPFileCheckedOut -Url $Item["FileRef"]
					$Temp = Set-PnPListItem -List $CurList.ID -Identity $Item.ID -ContentType $CT # Output to temp variable instead of screen
					Set-PnPFileCheckedIn -Url $Item["FileRef"] -CheckinType MajorCheckIn -Comment "Information Architecture: Updating Content Type"
				}
				catch 
				{
					# Error Output
					Write-Host ("            Unable to configure content type - Likely In Use or Checked out: {0}" -f $Item["FileLeafRef"]) -ForegroundColor Red
					Out-File -FilePath $OutPath -Append -InputObject ("{0}|{1}|{2}|{3}" -f $Site.Name, $List.Title, $Item["FileLeafRef"], $Item["FileRef"])

				}	
			}
		}
			
		# Screen Output for Library
		Write-Host ("         Library Configuration Complete: {0}" -f $List.Title) -ForegroundColor Green
	}

	# Screen Output for Library
	Write-Host ("      Site Configuration Complete: {0}" -f $Site.Name) -ForegroundColor Green
}
#endregion

#region Completed Screen Output
Complete-CatScript $Start[0]
#endregion

#region 5.	XML Example
<#
<InfoArchCnfg> <!-- Configuration of Content Types and Default Values -->
    <Sites>
        <Site>
            <Name>Test SP Migration Site</Name>
            <URL>https://testclient.sharepoint.com/sites/spmig</URL>
            <ContentType>Test Document</ContentType>
            <Defaults>
                <Default><Column>TestCol</Column><Value>Test Term Group|Test Term Set|Test Term</Value></Default>
                <Default><Column>TestCol2</Column><Value>Test Term Group 2|Test Term Set 2|Test Term 2</Value></Default>
			</Defaults>
			<ConfigItems>Yes</ConfigItems>
        </Site>
    </Sites>
</InfoArchCnfg>
#>
#endregion