#region Documentation
<#  Cat.PnP-CreateInfoArch.ps1

This script script iterates through and XML file to create Term stores, Term Groups, Terms, Site Columns, and Site Content Types
in a SharePoint Site collection.  If used in conjunction with a content type hub, content types will be replcated to all site
collections.
This script currently does not support Azure App Registrations. Writing to the Term Store is not yet supported. This script should be updated when MS allows read/write in the TermStore with an app registration.
This script can be run repeatedly with no issues, if the objects already exist they will be reported in the PowerShell window. 

Dependencies
	-	PnP PowerShell https://msdn.microsoft.com/en-us/pnp_powershell/pnp-powershell-overview
	-	XML Input File
	-	Global Admin Access if SharePoint Online
	-	Site Collection Admin rights

**** XML Example at end of script ****

Call script with parameters
	.\Cat.PnP-CreateInfoArch.ps1 -inputFile "Cat.InfoArch-OnSemi-Provisioning.xml" -siteUrl "https://nrpdev2.sharepoint.com/sites/DemoSiteProvisioning" -skipTermStore $false -appID "caf6b11a-cf03-416c-b0f0-67a09a8e30ee" -certPath "C:\GitHub\NRPDev\AppRegistration\PowerShell App Registration.pfx" -clientTenant "nrpdev2"

Improvements:
    -   Set Metadata to Open
    -   Configure Column for Multiple values and adhoc values
    -   Add Sub terms or import config
    -   Re Use terms previously created
    -   Remove un wanted columns from children content types

Table of Contents
	1.	Snapin(s), Modules & Environment Connection(s)
    2.	Variables
	3.	Input & Output Files
	4.	Main Script
	5.	Configure Term Store
	6.	Configure Columns & Content Types
#>
#endregion

#region Script Parameters
Param(	[string] $inputFile = "Cat.InfoArch-Test.xml", # Must exist in _Input Files folder
        [string] $siteUrl = "{SharePoint Site URL}", # SP site to create objects in
        [boolean] $skipTermStore = $false, # If the terms already exist you will want to skip them in future runs of the script
        [string] $appID = "{Azure App Registration ID}", # no {}
		[string] $certPath = "C:\AppRegistration\AzureAppRegistrationCert.pfx", # full local path
        [string] $clientTenant = "{client}" # IE. {client}.sharepoint.com  No {}
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
Start-Transcript -Path ("..\_Log Files\{0}.PnP-CreateInfoArch{1}.log" -f $clientTenant, $scriptStart[1])
#endregion

#region 1.	Snapin(s), Modules & Environment Connection(s)
# Load MS / PnP PowerShell Modules
Get-Module -ListAvailable | Where-Object{($_.Name -like "*SharePoint*") -or ($_.Name -like "*PnP*") } | Remove-Module
Import-Module PnP.PowerShell
#endregion

#region 2.	Variables
# User Inputs
If ($appID) {
	Write-Host
	Write-Host "User Input Variables" -ForegroundColor Cyan
	$certPass = ConvertTo-SecureString -String (Read-Host -Prompt "   Enter certificate password for the app registration") -AsPlainText -Force
	$cert = Get-PnPAzureCertificate -Path $certPath -Password $certPass
	Remove-Variable certPass
	Remove-Variable certPath
}

# Script Parameters
$spoAdminURL = ("https://{0}-admin.sharepoint.com" -f $clientTenant)
$tenantDomain = ("{0}.onmicrosoft.com" -f $clientTenant)

# Script Variables
$termGroupSleep = 10
$termSetSleep = 10
$termSleep = 10
$ctSleep = 10

# Output Parameters & Variables
Write-Host
Write-Host "Script Parameters" -ForegroundColor Cyan
Write-Host ("   Input File: {0}" -f $inputFile)
Write-Host ("   Site URL: {0}" -f $siteUrl)
if ($appID) {
	Write-Host ("   App ID: {0}" -f $AppID)
	Write-Host ("   Certificate Thumbprint: {0}" -f "Only output for troubleshooting", $cert.Thumbprint)
}
Write-Host ("   Client Tenant: {0}" -f $clientTenant)
Write-Host
Write-Host "Script Variables" -ForegroundColor Cyan
Write-Host ("   SP Admin URL: {0}" -f $spoAdminURL)
Write-Host ("   Tenant Domain: {0}" -f $tenantDomain)
Write-Host ("   Term Group Sleep (Seconds): {0}" -f $termGroupSleep)
Write-Host ("   Term Set Sleep (Seconds): {0}" -f $termSetSleep)
Write-Host ("   Term Sleep (Seconds): {0}" -f $termSleep)
Write-Host ("   Content Type Sleep (Seconds): {0}" -f $ctSleep)
#endregion

#region 3.	Input & Output Files
# File Paths
$inPath = ("..\_Input Files\{0}" -f $inputFile)

# Import Input File(s)
[xml]$InFile = get-content $InPath
#endregion
    
#region 4.	Main Script

#region 5.	Configure Term Store
If (!$skipTermStore){
    # Connect To SPO Admin Site For Term Store
    Write-Host
    Write-Host ("Connecting to SharePoint Admin site: {0}" -f $spoAdminURL) -ForegroundColor Magenta
    if ($AppID) {
		Connect-PnPOnline $spoAdminURL -tenant $tenantDomain -ClientId $AppID -Thumbprint $cert.Thumbprint -ValidateConnection -ErrorAction Stop # PnP Modules w/ App Reg
	} else {
		Connect-PnPOnline -Url $spoAdminURL -Interactive -ValidateConnection -ErrorAction Stop # PnP Modules w/o App Reg
	} 

    # Add Term Groups
    ForEach ($Group in $InFile.InfoArch.TermGroups.TermGroup)
    {
        Write-Host ("   Adding Term Group: {0}" -f $Group)
        try {
            $TempGroup = New-PnPTermGroup -Name $Group -ErrorAction Stop
            Start-Sleep -Seconds $termGroupSleep
        }
        catch {
            Write-Host ("      Term Group error: {0}" -f $_) -ForegroundColor Yellow
        }
    }

    # Add Term Sets
    ForEach ($Set in $InFile.InfoArch.TermSets.TermSet){
        Write-Host ("   Adding Term Set: {0}" -f $Set.Name)
        try {
            $TempSet = New-PnPTermSet -TermGroup $Set.Group -Name $Set.Name -ErrorAction Stop
            Start-Sleep -Seconds $termSetSleep
        }
        catch {
            Write-Host ("      Term Set error: {0}" -f $_) -ForegroundColor Yellow
        }        
    }

    # Add Terms
    ForEach ($Term in $InFile.InfoArch.Terms.Term){
        Write-Host
        Write-Host ("Adding Term: {0}" -f $Term.Name) -ForegroundColor DarkGray
        $customProperties = @{}
        foreach ($property in $Term.CustomProperties.Property){$customProperties.Add($property.Name,$property.Value)}
        try {
            $customProperties
            $TempTerm = New-PnPTerm -TermGroup $Term.Group -TermSet $Term.Set -Name $Term.Name -CustomProperties $customProperties -ErrorAction Stop
            Start-Sleep -Seconds $termSleep
        }
        catch {
            Write-Host ("   Term error: {0}" -f $_) -ForegroundColor Yellow
        }          
    }
    Disconnect-PnPOnline
}
#endregion #>

#region 6.	Configure Columns & Content Types
# Connect to SharePoint Site for Columns & Content Types
Write-Host
Write-Host ("Connecting to SharePoint site: {0}" -f $siteUrl) -ForegroundColor Magenta
if ($AppID) {
    Connect-PnPOnline $siteUrl -tenant $tenantDomain -ClientId $appID -Thumbprint $cert.Thumbprint -ValidateConnection -ErrorAction Stop # PnP Modules w/ App Reg
} else {
    Connect-PnPOnline -Url $siteUrl -Interactive -ValidateConnection -ErrorAction Stop # PnP Modules w/o App Reg
} 

# Add Columns
ForEach ($Column in $InFile.InfoArch.Fields.Field)
{
	If ($Column.Type -eq "Taxonomy")
	{
		Write-Host ("   Creating Column: {0}" -f $Column.DisplayName)
        try {
            $TempCol = Add-PnPTaxonomyField -DisplayName $Column.DisplayName -InternalName $Column.StaticName -Group $Column.Group -TermSetPath $Column.TermSetPath -ErrorAction Stop
            Start-Sleep -Seconds $CTSleep
        }
        catch {
            Write-Host ("      Column error: {0}" -f $_) -ForegroundColor Yellow
        }           
	}
	else 
	{
		Write-Host ("   Creating Column: {0}" -f $Column.DisplayName)
        #$TempCol = Add-PnPField -DisplayName $Column.Display -InternalName $Column.Internal -Group $Column.Group -Type $Column.Type
        try {
            $TempCol = Add-PnPFieldFromXml -FieldXml $Column.OuterXml -ErrorAction Stop
            Start-Sleep -Seconds $CTSleep
        }
        catch {
            Write-Host ("      Column error: {0}" -f $_) -ForegroundColor Yellow
        }          
	}
}

# Add Content Types
ForEach ($ContentType in $InFile.InfoArch.ContentTypes.ContentType)
{
	# Get Parent Content Type
	Write-Host ("   Creating Content Type: {0}" -f $ContentType.Name)
    try {
        $Parent = Get-PnPContentType -Identity $ContentType.Parent -ErrorAction Stop
        Start-Sleep -Seconds $CTSleep
    }
    catch {
        Write-Host ("      Parent Content Type error: {0}" -f $_) -ForegroundColor Yellow
    } 
    # Create Content Type
    try {
        $TempCT = Add-PnPContentType -Name $ContentType.Name -Group $ContentType.Group -ParentContentType $Parent -ErrorAction Stop
        Start-Sleep -Seconds $CTSleep
    }
    catch {
        $TempCT = Get-PnPContentType -Identity $ContentType.Name -ErrorAction Stop
        Write-Host ("      Parent Content Type error: {0}" -f $_) -ForegroundColor Yellow
    }    

	# Add Columns to Content Type
	ForEach ($Field in $ContentType.Fields.Field)
	{
		Write-Host ("      Adding Column to Content Type: {0}" -f $Field) -ForegroundColor DarkGray
        try {
            Add-PnPFieldToContentType -Field $Field -ContentType $TempCT -ErrorAction Stop
            Start-Sleep -Seconds $CTSleep
        }
        catch {
            Write-Host ("         Adding Column error: {0}" -f $_) -ForegroundColor Yellow
        }         
	}
}
#endregion #>

# Disconnecting from SharePoint Online
Disconnect-PnPOnline
#endregion

#region Completed Script
# Output Script Info
Complete-CatScript $scriptStart[0]

# Stop logging
Stop-Transcript
#endregion