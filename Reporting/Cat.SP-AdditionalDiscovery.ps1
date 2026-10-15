<#region Cat.SP-Additional Discovery.ps1

This script leveraged SMAT reports as the starting point for discovery of on Premsises SharePoint environments.  The intent is to loop through 
SMAT reports to gather additional information.  This script leverages servers side SharePoint Modules in PowerShell

*** This script has been validated with SharePoint 2013 and needs to be validated on 2010, 2016, 2019

Reports
	-	Farm Solutions
	-	Webs
	-	Complex Columns
	-	Item Last Modified (List object property in accurate)
	-	Checked Out Items

Dependencies
	-	Microsoft.SharePoint.PowerShell
	-	SMAT Results including (Site Assessment, Large lists set to 0, checked out documents, workflows, infopath, etc.)
		*	It is expecting the default structure and file names generated when running SMAT
	-	General Functions - ..\General\Cat.GeneralFunctions.psm1
	-	Permissions: SharePoint Farm Administrator, DB Owner on all SP Databases, Full Control Web Application Policy

Improvements Needed
	-	Use .NET Queues to handle CSV (Similar to webs report)
    -   Update all reports to build array
	-	Upgrade Where clauses to Where-Object
	-	Add JSLinkfrom On Semi
    -   InfoPath: Update nested if/else to use Switch instead
    -   InfoPath: Use workflow association to identify list on WF Forms not associated directly to the list

Call script with parameters
    cd C:\GitHub\CatPowerShell\Reporting
	.\Cat.SP-AdditionalDiscovery.ps1 -SMATPath "..\..\SMAT\_210420" -SMATScanID "720864fe-c275-4a65-bbd3-4fab48edb4cf" -client "Client" -spVersion "SP19" -outputLocation "C:\GitHub\CatPowerShell\_Output Files"

Table of Contents
	1.	Snapin(s), Modules & Environment Connection(s)
	2.	Input & Output Files
	3.	Variables
	4.	Main Script
	5.	Farm Solutions
    6.  InfoPath Details
    7.  Workflow 2010 Details
    8.  Workflow 2013 Details
	9.	Webs
	10.	Checked Out Documents
	11.	Scanning Lists
		11a.	Complex Columns
		11b.	Last Item Modified

****************************************************************************************************************
****    This script has been validated with SP2013 and needs to be fully tested in other versions of SP     ****
****************************************************************************************************************
    
#endregion #>

#region Script Parameters
Param(	[string] $SMATPath = "..\..\SMAT\_210409",
		[string] $SMATScanID = "5f02d932-4042-4b7d-99bf-3b9241dc1c42",
		[string] $client = "Client",
        [string] $spVersion = "SP13", # SP19, SP16, SP13 or SP10
        [string] $outputLocation = "C:\GitHub\CatPowerShell\_Output Files" # the script will create directories underneath this directory
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
Start-Transcript -Path ("..\_Log Files\{0}.SP-AdditionalDiscovery-{1}.log" -f $client, $scriptStart[1])
#endregion

#region 1.	Snapin(s), Modules & Environment Connection(s)
# Load PowerShell Modules
Get-Module -ListAvailable | Where-Object{($_.Name -like "*SharePoint*") -or ($_.Name -like "*PnP*") } | Remove-Module
Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue
#endregion

#region 2.	Input & Output Files
# Input Files
$inputSiteAssessment = ("{0}\SiteAssessmentReport.csv" -f $SMATPath)
$inputLists = ("{0}\ScannerReports\LargeLists-detail.csv" -f $SMATPath)
$inputCheckedOutDocs = ("{0}\ScannerReports\CheckedOutFiles-detail.csv" -f $SMATPath)
$inputInfoPathReport = ("{0}\ScannerReports\InfoPath-detail.csv" -f $SMATPath)
$inputWF2010Report = ("{0}\ScannerReports\WorkflowAssociations2010-detail.csv" -f $SMATPath)
$inputWF2013Report = ("{0}\ScannerReports\WorkflowAssociations2013-detail.csv" -f $SMATPath)

# Output Files
$outputPath = ("{0}\{1}\{2}" -f $outputLocation, $client, $scriptStart[1])
$tempPath = ("{0}\Temp" -f $outputPath)
$outputFarmSolutions = ("{0}\catFarmSolutions.csv" -f $outputPath)
$outputInfoPathDetails = ("{0}\catInfoPathDetails.csv" -f $outputPath)
$outputWebs = ("{0}\catWebs.csv" -f $outputPath)
$outputColumns = ("{0}\catColumns.csv" -f $outputPath)
$outputLastModifiedItem = ("{0}\catLastModifiedItem.csv" -f $outputPath)
$outputCheckedOutDocs = ("{0}\catCheckedOut.csv" -f $outputPath)
$outputWorkflowDetails2010 = ("{0}\catWorkflowDetails2010.csv" -f $outputPath)
$outputWorkflowDetails2013 = ("{0}\catWorkflowDetails2013.csv" -f $outputPath)
#endregion

#region 3.	Variables
$specialCharacters = "[~``!@#$%^&*()_-``""']"
$complexColumns = @("Calculated";"TaxonomyFieldType";"Lookup")
$skipListURLs = "/_catalogs/|/Access Requests|/WmaAggregatorList_User"
$batchExportThreshold = 5000
$currentPSVersion = $PSVersionTable.PSVersion.Major
$batchPSVersion = 3
$queryLastItemModified = New-Object Microsoft.SharePoint.SPQuery
$queryLastItemModified.ViewXml = "@<View Scope='RecursiveAll'>
	<Query>
		<OrderBy>
			<FieldRef Name='Modified' Ascending='False'/>
		</OrderBy>
	</Query>
	<RowLimit>1</RowLimit>
</View>"
# Script Variables

# Output Parameters & Variables
Write-Host
Write-Host "Script Parameters" -ForegroundColor Cyan
Write-Host ("   SMAT Output Path: {0}" -f $SMATPath)
Write-Host ("   SMAT Scan ID: {0}" -f $SMATScanID)
Write-Host ("   Client: {0}" -f $client)
Write-Host ("   SharePoint Version: {0}" -f $spVersion)
Write-Host ("   Output Base Path: {0}" -f $outputLocation)
Write-Host
Write-Host "Input Files" -ForegroundColor Cyan
Write-Host ("   SMAT Site Assessment: {0}" -f $inputSiteAssessment)
Write-Host ("   SMAT Large Lists: {0}" -f $inputLists)
Write-Host ("   SMAT Checked Out Files: {0}" -f $inputCheckedOutDocs)
Write-Host ("   SMAT InfoPath Report: {0}" -f $inputInfoPathReport)
Write-Host ("   SMAT WF 2010 Report: {0}" -f $inputWF2010Report)
Write-Host ("   SMAT WF 2013 Report: {0}" -f $inputWF2013Report)
Write-Host
Write-Host "Output Files" -ForegroundColor Cyan
Write-Host ("   Output Directory: {0}" -f $outputPath)
Write-Host ("   Temp Directory: {0}" -f $tempPath)
Write-Host ("   Farm Solutions Report: {0}" -f $outputFarmSolutions)
Write-Host ("   Webs Report: {0}" -f $outputWebs)
Write-Host ("   Complex Columns Report: {0}" -f $outputColumns)
Write-Host ("   Last Modified Item Report: {0}" -f $outputLastModifiedItem)
Write-Host ("   Checked OutFiles Report: {0}" -f $outputCheckedOutDocs)
Write-Host ("   InfoPath Report: {0}" -f $outputInfoPathDetails)
Write-Host ("   SP 2010 WF Report: {0}" -f $outputWorkflowDetails2010)
Write-Host ("   SP 2013 wF Report: {0}" -f $outputWorkflowDetails2013)
Write-Host
Write-Host "Script Variables" -ForegroundColor Cyan
Write-Host ("   Special Characters to Remove: {0}" -f $specialCharacters)
Write-Host ("   Skip List URLs: {0}" -f $skipListURLs)
Write-Host ("   Batch Export Threshold: {0}" -f $batchExportThreshold)
Write-Host ("   PS Version: {0}" -f $currentPSVersion)
Write-Host ("   PS Version required for Batching: {0}" -f $batchPSVersion)
Write-Host
Write-Host "Complex Column Types" -ForegroundColor Cyan
Write-Host $complexColumns
Write-Host
Write-Host "Last Item Modified CAML Query" -ForegroundColor Cyan
Write-Host $queryLastItemModified.ViewXml
#endregion
    
#region 4.	Main Script
#region Output warning if running version of PS that doesn't support append
if ($currentPSVersion -lt $batchPSVersion) {
    Write-Host
    Write-Host "****" -ForegroundColor Yellow
    Write-Host
    Write-Host "This version of PowerShell doesn't support batch exporting. Server memory may be an issue and should be monitored." -ForegroundColor Yellow
    Write-Host
    Write-Host "****" -ForegroundColor Yellow
    $userProceed = Read-Host "Are you sure you want to proceed with the scan? (Yes or No)" 
	if ($userProceed.ToLower() -eq "yes") {
		Write-Host "   Continuing environment scan" -ForegroundColor Green
	}
	else {
		Write-Host "   Stopping environment scan" -ForegroundColor Red
        Break
	}
    Write-Host
}
#endregion

# Created directories for additional reports
Write-Host
Write-Host ("{0} - Creating Report and Temp Directories" -f (Get-Date -Format "yy-MM-dd HH:mm:ss")) -ForegroundColor Cyan
New-Item -Path $outputPath -ItemType Directory
New-Item -Path $tempPath -ItemType Directory

#region 5.	Farm Solutions
Write-Host
Write-Host ("{0} - Gathering Farm Solutions" -f (Get-Date -Format "yy-MM-dd HH:mm:ss")) -ForegroundColor Magenta -NoNewline 
$farmSolutions = New-Object System.Collections.Queue;
Get-SPSolution | ForEach-Object{$farmSolutions.Enqueue($_)}
If ($farmSolutions.Count -gt 0) {
	#loop through solutions queue
    $farmSolutionDetails = @()
	while($farmSolutions.Count -ne 0) {
		$farmSolutions.Dequeue() | ForEach-Object{
            $solution = New-Object -TypeName PSObject -Property (@{ScanId=$SMATScanID;DisplayName=($_.DisplayName -replace $specialCharacters,'');ID=$_.Id;Status=$_.Status;Name=$_.Name;Farm=$_.Farm;
                Deployed=$_.Deployed;DeploymentState=$_.DeploymentState;ContainsGlobalAssembly=$_.ContainsGlobalAssembly;ContainsCasPolicy=$_.ContainsCasPolicy;
                ContainsWebApplicationResource=$_.ContainsWebApplicationResource;SolutionId=$_.SolutionId;IsWebPartPackage=$_.IsWebPartPackage;LanguagePacks=$_.LanguagePacks;
                Added=$_.Added;LastOperationResult=$_.LastOperationResult;LastOperationEndTime=$_.LastOperationEndTime;JobStatus=$_.JobStatus;JobExists=$_.JobExists;
                CanSelectForBackup=$_.CanSelectForBackup;CanSelectForRestore=$_.CanSelectForRestore;Parent=$_.Parent;Version=$_.Version;CanBackupRestoreAsConfiguration=$_.CanBackupRestoreAsConfiguration;
                DiskSizeRequired=$_.DiskSizeRequired;CanRenameOnRestore=$_.CanRenameOnRestore})
            $farmSolutionDetails += $solution
            Clear-Variable solution
		}
	}

    #exporting to CSV file
    $farmSolutionDetails | Export-Csv -Path $outputFarmSolutions -NoTypeInformation
	Write-Host " - Success" -ForegroundColor Green
}
Else {
	Write-Host " - No Farm Solutions Found" -ForegroundColor Yellow
}

# Cleaning up memory
Remove-Variable farmSolutions
#endregion #>

#region 6.  InfoPath Details
# Load InfoPath file
If (Test-Path ($inputInfoPathReport)){
	$InfoPathReport = Import-Csv $inputInfoPathReport

    # Loop InfoPath Reprot and collect Details
    Write-Host
    Write-Host ("{0} - Scanning InfoPath: {1} forms" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"),$InfoPathReport.count) -ForegroundColor Magenta
    $InfoPathFormDetails = @()
    $infopathIndex = 0

    foreach ($InfoPathForm in $InfoPathReport)
    {
        Write-Host ("   {0} - Analyzing Form: {1}" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"),$InfoPathForm.URL)
        
        # Get the BaseTemplate
        $site = New-Object Microsoft.SharePoint.SPSite($InfoPathForm.URL)
        $web = $site.OpenWeb()
        $parentListURL = $InfoPathForm.URL.Replace("/Task/template.xsn","").Replace("/Item/template.xsn","").Replace("/Forms/template.xsn","").Replace("/Issue/template.xsn","").Replace("/Contact/template.xsn","")
        $parentListRootFolder = $parentListURL.Replace($web.Url + "/","")
        $list = $web.lists | Where-Object {$parentListRootFolder -eq  $_.RootFolder}

        # If it failed to get the BaseTemplate, use second method to grab it (For Workflow Forms)
        if(!$list.BaseTemplate)
        {
            $list = $web.lists | Where-Object {$parentListURL -match  $_.RootFolder}
        }

        # If it failed to get the BaseTemplate again, use third method to grab it (For non-latin characters on Custom List Forms)
        if(!$list.BaseTemplate)
        {
            $parentListURL = $InfoPathForm.URL.Replace("?","").Replace("/template.xsn","")
            $site = New-Object Microsoft.SharePoint.SPSite($parentListURL)
            $web = $site.OpenWeb()
            $parentListRootFolder = $parentListURL.Replace($web.Url + "/","").Replace("/","").Replace("Lists","Lists/")
            $list = $web.lists | Where-Object {$parentListRootFolder -eq  $_.RootFolder}
        }

        # If it failed to get the BaseTemplate again again, use fourth method to grab it (For non-latin characters on Form Libraries)
        if(!$list.BaseTemplate)
        {
            $parentListURL = $InfoPathForm.URL.Replace("/Forms/","").Replace("?","").Replace("/template.xsn","")
            $site = New-Object Microsoft.SharePoint.SPSite($parentListURL)
            $web = $site.OpenWeb()
            $parentListRootFolder = $parentListURL.Replace($web.Url + "/","").Replace("/","")
            $list = $web.lists | Where-Object {$parentListRootFolder -eq  $_.RootFolder}
        }

        # If a custom list form (Generic List, Issues List, Task List, etc), save last modified date from the list
        if($list.BaseTemplate -eq "GenericList" -or $list.BaseTemplate -eq "IssueTracking" -or $list.BaseTemplate -eq "CustomGrid" -or $list.BaseTemplate -eq "Contacts" -or $list.BaseTemplate -match "Tasks")
        {
            $InfoPathFormType = "Custom List Form"
    
            # Get the last modified List Items
            $LastModifiedItems = $List.GetItems($queryLastItemModified)
    
            # Loop through the one last modified item
            ForEach($LastModifiedItem in $LastModifiedItems) 
            {
                $itemLastModified = $LastModifiedItem["Modified"]
            }
        }

        # If it is a Form Library
        elseif($list.BaseTemplate -eq "XMLForm")
        {
        $InfoPathFormType = "Form Library"

            # If it is a Forms library with the default URL, we can't get to the file object, have to hit the list via URL
            if($InfoPathForm.URL -like "*/Forms/*.xsn")
            { 
                #Get the last modified List Items
                $LastModifiedItems = $List.GetItems($queryLastItemModified)
    
                #Loop through the one last modified item
                ForEach($LastModifiedItem in $LastModifiedItems) 
                {
                    $itemLastModified = $LastModifiedItem["Modified"]
                }

            }
            # If it is a Forms library with custom name, we can get the list via the file object
            else
            {
                $item = $web.GetListItem($InfoPathForm.URL)
                $parentListURL = $site.Protocol + "//" + $site.HostName + $item.ParentList.RootFolder.ServerRelativeUrl
                $parentListRootFolder = $parentListURL.Replace($web.Url + "/","")
                $list = $web.lists | Where-Object {$parentListRootFolder -eq  $_.RootFolder}
    
                #Get the last modified List Items
                $LastModifiedItems = $List.GetItems($queryLastItemModified)
    
                #Loop through the one last modified item
                ForEach($LastModifiedItem in $LastModifiedItems) 
                {
                    $itemLastModified = $LastModifiedItem["Modified"]
                }

            }

        }

        #If Workflow Form: the parentListURL will not be correct since the form is not associated with a list. Modifed data will be for the Workflow Initiation Form itself. Not for the List.
        elseif($list.BaseTemplate -eq "NoCodeWorkflows" -or $list.BaseTemplate -eq "WorkflowProcess" -or $list.BaseTemplate -eq "NoCodePublic")
        {
            $InfoPathFormType = "Workflow Initiation Form"
            $item = $web.GetListItem($InfoPathForm.URL)
            $parentListURL = "No List Association"
            $itemLastModified = $item["Modified"]
        }

        #If an actual InfoPath Document, grab the list from the file object, save last modified date from the file itself
        elseif($list.BaseTemplate -eq "DocumentLibrary")
        {
            $InfoPathFormType = "InfoPath Document"
            $item = $web.GetListItem($InfoPathForm.URL)
            $parentListURL = $site.Protocol + "//" + $site.HostName + $item.ParentList.RootFolder.ServerRelativeUrl
            $itemLastModified = $item["Modified"]
        }

        #Else - Catch All on random stuff
        else
        {
            $InfoPathFormType = "CATCH-ALL"
            $item = $web.GetListItem($InfoPathForm.URL)
            $parentListURL = $site.Protocol + "//" + $site.HostName + $item.ParentList.RootFolder.ServerRelativeUrl
            Write-Host " = Catch All on this one..." -ForegroundColor Red
            Write-Host " = Parent List URL:" $parentListURL -ForegroundColor Red
        }
        
        # Extracting XML
        #Load SharePoint CSOM Assemblies 
        switch ($spVersion) {
            {$_ -in "SP16", "SP19"} {
                Add-Type -Path "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll"
                Add-Type -Path "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Runtime.dll"
            }
            SP13 {
                Add-Type -Path "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\15\ISAPI\Microsoft.SharePoint.Client.dll"
                Add-Type -Path "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\15\ISAPI\Microsoft.SharePoint.Client.Runtime.dll"
            }
            SP10 {
                Add-Type -Path "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\14\ISAPI\Microsoft.SharePoint.Client.dll"
                Add-Type -Path "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\14\ISAPI\Microsoft.SharePoint.Client.Runtime.dll"
            }
            Default { 
                Write-Host "Incorrect SP Version"
                Break
            }
        }  

        #Set parameter values 
        $SiteURL = $site.Protocol + "//" + $site.HostName
        $SourceRelativePath = $InfoPathForm.URL.Replace($SiteURL,"")
        $TargetFilePath =  ("{0}\Temporary-InfoPath.xsn" -f $tempPath)
        $CabFilePath = $TargetFilePath.Replace(".xsn",".cab")  
        
        #Setup the context         
        $Ctx = New-Object Microsoft.SharePoint.Client.ClientContext($SiteURL)
                
        #Download file from SP
        switch ($spVersion) {
            {$_ -in "SP13","SP16","SP19"} {
                $FileInfo = [Microsoft.SharePoint.Client.File]::OpenBinaryDirect($Ctx,$SourceRelativePath)
                $WriteStream = [System.IO.File]::Open($TargetFilePath,[System.IO.FileMode]::Create)
                $FileInfo.Stream.CopyTo($WriteStream)
                $WriteStream.Close()
            }
            SP10 {
                $downloadWeb = Get-SPWeb $SiteURL
                $downloadFile = $downloadWeb.GetFile($SourceRelativePath)
                $downloadbytes = $downloadFile.OpenBinary()
                $downloadStream = New-Object System.IO.FileStream($TargetFilePath,"Create")
                $downloadWriter = New-Object System.IO.BinaryWriter($downloadStream)
                $downloadWriter.write($downloadbytes)
                $downloadWriter.Close()
            }
        }      
        
        #Rename the XSN to CAB so we can extract it
        Rename-Item -Path $TargetFilePath -NewName $CabFilePath     
        
        #Create folder in the Output folder
        $NewFolder = $tempPath + "\" + $SourceRelativePath.Replace("/","-")
        $createfolder = md $NewFolder
        $ManifestPath = $NewFolder + "\manifest.xsf"

        #Extract the CAB
        $extract = cmd.exe /c "C:\Windows\System32\expand.exe" -F:* $CabFilePath $NewFolder
    
        #Delete the temporary CAB to avoid file name duplicates
        $delete = Remove-Item -Path $CabFilePath -Force

        #Loop all the views to count Rules, Buttons, and Controls
        $ViewsArray = @()
        [xml]$ManifestContent = Get-Content -Path $ManifestPath  
        $ViewsArray = $ManifestContent.xDocumentClass.views.view.mainpane.transform
        
        foreach($View in $ViewsArray){
            #Write-Host "Looking inside View:" $View -f Green
            $ViewPath = $NewFolder + "\" + $View 
                
            #Count number of Rules in View
            $XMLRules = Select-String -Path $ViewPath -pattern "xsl:if test" -AllMatches
            $XMLRules1 = Select-String -Path $ViewPath -pattern "caption:" -AllMatches
            $XMLRulesTotal += $XMLRules.Matches.count
            $XMLRulesTotal += $XMLRules1.Matches.count
            #Write-Host " = Number of Rules found:" $XMLRulesTotal -f White 

            #Count number of Buttons in View
            $XMLButtons = Select-String -Path $ViewPath -pattern 'type="button"' -AllMatches
            $XMLButtonsTotal += $XMLButtons.Matches.count
            #Write-Host " = Number of Buttons found:" $XMLButtons.Matches.count -f White 

            #Count number of Form Controls in View
            $XMLFormControls = Select-String -Path $ViewPath -pattern "xd:CtrlId" -AllMatches
            $XMLFormControlsTotal += $XMLFormControls.Matches.count 
            #Write-Host " = Number of Form Controls found:" $XMLFormControls.Matches.Count -f White    
        }

            #Subtract Buttons from the Controls count
            $XMLFormControlsTotal -= $XMLButtonsTotal
            #Write-Host " = TOTAL Number of Form Controls removed buttons:" $XMLFormControlsTotal -f Green 


        #Count number of Views in Manifest.xsf
        $XMLViews = $ViewsArray.Count
        #Write-Host " = Number of Views found:" $XMLViews -f Yellow 
    
        #Count number of Data Connections in Manifest.xsf
        $XMLDataConnections = Select-String -Path $ManifestPath -pattern "</xsf:dataObject>" -AllMatches
        $XMLDataConnections1 = Select-String -Path $ManifestPath -pattern "</xsf:davAdapter>" -AllMatches
        $XMLDataConnections2 = Select-String -Path $ManifestPath -pattern "</xsf:webServiceAdapter>" -AllMatches
        $XMLDataConnections3 = Select-String -Path $ManifestPath -pattern 'name="Main Data Connection"' -AllMatches
        $XMLDataConnections4 = Select-String -Path $ManifestPath -pattern "</xsf:submitToHostAdapter>" -AllMatches #This is being used in Workflow Initiation form
        $XMLDataConnectionsTotal += $XMLDataConnections.Matches.count
        $XMLDataConnectionsTotal += $XMLDataConnections1.Matches.count
        $XMLDataConnectionsTotal += $XMLDataConnections2.Matches.count
        $XMLDataConnectionsTotal += $XMLDataConnections3.Matches.count
        $XMLDataConnectionsTotal += $XMLDataConnections4.Matches.count
        If($XMLDataConnections4.Matches.count -gt 0) # If Workflow Initiation form, we need to remove 1 since Main Data Connection is gets reported twice
        {
            $XMLDataConnectionsTotal -= 1
        }
        #Write-Host " = Number of Data Connections found:" $XMLDataConnectionsTotal -f Yellow 

        #Count number of Data Fields in Manifest.xsf
        $XMLDataFields = Select-String -Path $ManifestPath -pattern "</xsf:field>" -AllMatches
        
        #Count number of Calculations in Manifest.xsf
        $XMLCalculations = Select-String -Path $ManifestPath -pattern "</xsf:calculatedField>" -AllMatches

        #Delete Unzipped folder
        Remove-Item -Path $NewFolder -Recurse -Force -Confirm:$false

        # Add to Excel Export Array
        $itemArrayInfoPath = New-Object PSObject
        $itemArrayInfoPath | Add-Member -type NoteProperty -Name 'SiteURL' -Value $InfoPathForm.SiteURL
        $itemArrayInfoPath | Add-Member -type NoteProperty -Name 'FormURL' -Value $InfoPathForm.URL
        $itemArrayInfoPath | Add-Member -type NoteProperty -Name 'ListURL' -Value $parentListURL
        $itemArrayInfoPath | Add-Member -type NoteProperty -Name 'LastModified' -Value $itemLastModified
        $itemArrayInfoPath | Add-Member -type NoteProperty -Name 'ListItemCount' -Value $list.ItemCount
        $itemArrayInfoPath | Add-Member -type NoteProperty -Name 'FormType' -Value $InfoPathFormType
        $itemArrayInfoPath | Add-Member -type NoteProperty -Name 'Mode' -Value $InfoPathForm.Mode
        $itemArrayInfoPath | Add-Member -type NoteProperty -Name 'ListBaseTemplate' -Value $list.BaseTemplate
        $itemArrayInfoPath | Add-Member -type NoteProperty -Name 'XMLViews' -Value $XMLViews
        $itemArrayInfoPath | Add-Member -type NoteProperty -Name 'XMLDataConnections' -Value $XMLDataConnectionsTotal
        $itemArrayInfoPath | Add-Member -type NoteProperty -Name 'XMLDataFields' -Value $XMLDataFields.Matches.count
        $itemArrayInfoPath | Add-Member -type NoteProperty -Name 'XMLCalculations' -Value $XMLCalculations.Matches.count
        $itemArrayInfoPath | Add-Member -type NoteProperty -Name 'XMLRules' -Value $XMLRulesTotal
        $itemArrayInfoPath | Add-Member -type NoteProperty -Name 'XMLButtons' -Value $XMLButtonsTotal
        $itemArrayInfoPath | Add-Member -type NoteProperty -Name 'XMLFormControls' -Value $XMLFormControlsTotal
        $itemArrayInfoPath | Add-Member -type NoteProperty -Name 'ScanID' -Value $SMATScanID
        $InfoPathFormDetails += $itemArrayInfoPath
        $infopathIndex ++

        # Clear Variables between loops
        Remove-Variable parentListURL
        Remove-Variable InfoPathFormType
        Remove-Variable XMLDataConnectionsTotal
        Remove-Variable XMLRulesTotal
        Remove-Variable XMLFormControlsTotal
        Remove-Variable XMLButtonsTotal
        Remove-Variable XMLViews

        if(Test-Path variable:itemLastModified){
            Remove-Variable itemLastModified
        }

        # Optimize memory usage to write to CSV in batches if PS supports it
        if ($currentPSVersion -ge $batchPSVersion) {
            if(($infopathIndex % $batchExportThreshold) -eq 0){
                $InfoPathFormDetails | Export-Csv -Path $outputInfoPathDetails -NoTypeInformation -Append
                $InfoPathFormDetails = @()
                Write-Host "      Batch-point detected and wrote results to disk."
            }
        }
    }

    # Write to CSV if PS Doesn't support batches
    if ($currentPSVersion -lt $batchPSVersion) {
        $InfoPathFormDetails | Export-Csv -Path $outputInfoPathDetails -NoTypeInformation
    } Elseif (($infopathIndex % $batchExportThreshold) -gt 0) {
        $InfoPathFormDetails | Export-Csv -Path $outputInfoPathDetails -NoTypeInformation -Append
        Write-Host "      Final batch-point detected and wrote results to disk."
    }
}
else {
	Write-Host
    Write-Host ("{0} - InfoPath Report Missing: {1}" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"), $inputInfoPathReport) -ForegroundColor Yellow
}
#endregion #>

#region 7.  Workflow 2010 Details
# Test amd Import the Input files
If (Test-Path ($inputWF2010Report)){
    # Import SMAT 2010 Workflow Report
	$WF2010Report = Import-Csv $inputWF2010Report | Select-Object 'SiteUrl','Scope','RunningInstances','WebUrl','ListTitle','ListUrl','IsReusable','ContentTypeName','WorkflowName'

    Write-Host
    Write-Host ("{0} - Scanning 2010 Workflow: {1} workflows" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"),$WF2010Report.count) -ForegroundColor Magenta
    $WorkflowDetails2010 = @()
    $WF2010XML =@()
    $Workflow2010Index = 1

    foreach ($WF2010 in $WF2010Report)
    {
        Write-Host ("   {0} - {1} of {2} - List: {3} - Workflow: {0}" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"),$Workflow2010Index,$WF2010Report.count, $WF2010.ListUrl, $WF2010.WorkflowName)

        # Get the Site, Web, and List Objects - This is a very expensive call so we only want to do it if needed.
        If($oldListURL -ne $WF2010.ListUrl){
            
            # Dispose Objects (if they exist) since we are about to load a new one
            if($site){
                $site.Dispose()
                $web.Dispose()
            }
            # Create the site and web objects
            $site = New-Object Microsoft.SharePoint.SPSite($WF2010.ListUrl)
            $web = $site.OpenWeb()
        } 

        $parentListURL = $WF2010.ListUrl.Replace($web.Url + "/","")
        $list = $web.lists | Where-Object {$parentListURL -eq $_.RootFolder}
        #Write-Host " = List Count:" $list.ItemCount -f Yellow

        # Get Last Modified from List Items (if there are any items there)
        if($list.ItemCount -gt 0){
            #Get the last modified List Items, need to save the date in a variable since it will fail on the Array if we refer to first item [0]
            $LastModifiedItems = $List.GetItems($queryLastItemModified)
            #Write-Host " = List ITEM Last Modified:" $LastModifiedItems[0]["Modified"] "(GOOD)" -f Yellow
            $LastModifiedItemDate = $LastModifiedItems[0]["Modified"]
        }

        # Get Workflow Association for this specific workflow. Slightly different from a List WF and a Content Type WF.     
        if ($WF2010.Scope -eq "List"){
            foreach ($wflowAssociation in $list.WorkflowAssociations){
                # Keep looping definitions if no match was found and TaskListTitle set
                if (!$TaskListTitle){
                    #Write-Host " = Matching Workflow Names:" $wflowAssociation.InternalName ". Contains:" $WF2010.WorkflowName

                    if($wflowAssociation.Name -eq $WF2010.WorkflowName -and $wflowAssociation.Name -notmatch "Previous Version"){
                        #Write-Host " = Association ListWF Name found:" $wflowAssociation.Name
                        #Write-Host " = Found ListWF Association Task List:" $wflowAssociation.TaskListTitle  
                        #Write-Host " = Found ListWF Association History List:" $wflowAssociation.HistoryListTitle
                        $ActivewflowAssociation = $wflowAssociation 
                        $TaskListTitle = $wflowAssociation.TaskListTitle  
                        $HistoryListTitle = $wflowAssociation.HistoryListTitle
                    }
                        else{
                            #Write-Host "Did not find ListWF wflowAssociation for:" $WF2010.WorkflowName ".That matches:" $wflowAssociation.InternalName -f Magenta
                        }
                }
            }   
        }           
        elseif ($WF2010.Scope -eq "ContentType"){
            foreach ($wflowAssociation in $list.ContentTypes.WorkflowAssociations){
                # Keep looking if no match was found
                if (!$TaskListTitle){
                    #Write-Host " = Matching Workflow Names:" $wflowAssociation.InternalName ". Contains the value from the CSV:" $WF2010.WorkflowName
                    if(($wflowAssociation.Name -eq $WF2010.WorkflowName) -and ($wflowAssociation.ParentContentType.Name -eq $wf2010.ContentTypeName) -and ($wflowAssociation.Name -notmatch "Previous Version")){
                        # Write Workflow Info for troubleshooting
                        #Write-Host " = CT-Workflow Name:" $wflowAssociation.Name  
                        #Write-Host " = Found CT-WF Association Task List:" $wflowAssociation.TaskListTitle  
                        #Write-Host " = Found CT-WF Association History List:" $wflowAssociation.HistoryListTitle #
                        $ActivewflowAssociation = $wflowAssociation 
                        $TaskListTitle = $wflowAssociation.TaskListTitle  
                        $HistoryListTitle = $wflowAssociation.HistoryListTitle
                        }
                        else{
                            #Write-Host "Did not find CT-WF wflowAssociation for:" $WF2010.WorkflowName ".That matches:" $wflowAssociation.InternalName -f Magenta
                        }
                }
            }              
        }         

        # Get last modified from Task list (if one was found)
        if($TaskListTitle){
            $TaskList = $web.lists | Where-Object {$TaskListTitle -eq $_.Title}
            # Write Task List Info for troubleshooting
            #Write-Host " = Task List Found:" $TaskList.RootFolder -f Green
            #Write-Host " = Task List Item Count:" $TaskList.ItemCount -f Green
            #Write-Host " = Task List Last Modified:" $TaskList.LastItemModifiedDate "(BAD)" -f Green #

            # Get Task List Last Modified from Items (if there are any items there)
            if($TaskList.ItemCount -gt 0){
                #Get the last modified Task List Items, need to save the date in a variable since it will fail on the Array if we refer to first item [0]
                $LastTaskModifiedItems = $TaskList.GetItems($queryLastItemModified)
                $TaskListItemLastModified = $LastTaskModifiedItems[0]["Modified"]
                #Write-Host " = Task List ITEM Last Modified:" $LastTaskModifiedItems[0]["Modified"] "(GOOD)"-f Green
    
            } 
        } 

        # Get last modified from History list (if one was found)
        if($HistoryListTitle){
            $HistoryList = $web.lists | Where-Object {$HistoryListTitle -eq $_.Title}
            # Write Task List Info for troubleshooting
            #Write-Host " = History List Found:" $HistoryList.RootFolder -f Magenta
            #Write-Host " = History List Item Count:" $HistoryList.ItemCount -f Magenta
            #Write-Host " = History List Last Modified:" $HistoryList.LastItemModifiedDate "(BAD)" -f Magenta #

            # Get History List Last Modified from Items (if there are any items there)
            if($HistoryList.ItemCount -gt 0){
                #Get the last modified Task List Items, need to save the date in a variable since it will fail on the Array if we refer to first item [0]
                $LastHistoryModifiedItems = $HistoryList.GetItems($queryLastItemModified)
                $HistoryListItemLastModified = $LastHistoryModifiedItems[0]["Modified"]
                #Write-Host " = History List ITEM Last Modified:" $LastHistoryModifiedItems[0]["Modified"] "(GOOD)"-f Magenta\
            }
        }

        # Identify Workflow Type and Get complexity stuff for Non OOTB workflows
        switch ($ActivewflowAssociation)
        {
            {($_.SoapXml -match "8ad4d8f0-93a7-4941-9657-cf3706f00409") -or ($_.SoapXml -match "c6964bff-bf8d-41ac-ad5e-b61ec111731c")}{$WorkflowType = "OOTB - Approval"}
            {$_.SoapXml -match "e43856d2-1bb4-40ef-b08b-016d89a00409"}{$WorkflowType = "OOTB - Approve Content for Publishing"} #TODO: Need to get Content Type Association GUID. Don't currently have an example
            {($_.SoapXml -match "3bfb07cb-5c6a-4266-849b-8d6711700409") -or ($_.SoapXml -match "46c389a4-6e18-476c-aa17-289b0c79fb8f")}{$WorkflowType = "OOTB - Collect Feedback"}
            {($_.SoapXml -match "77c71f43-f403-484b-bcb2-303710e00409") -or ($_.SoapXml -match "2f213931-3b93-4f81-b021-3022434a3114")}{$WorkflowType = "OOTB - Collect Signature"}
            {$_.SoapXml -match "dd19a800-37c1-43c0-816d-f8eb5f4a4145"}{$WorkflowType = "OOTB - Disposition Approval"} #TODO: Need to get Content Type Association GUID. Don't currently have an example
            {$_.SoapXml -match "c6964bff-bf8d-41ac-ad5e-b61ec111731a"}{$WorkflowType = "OOTB - Three State Approval"} #TODO: Need to get Content Type Association GUID. Don't currently have an example
            
            # Since this is a custom 2010 workflow, run the complexity
            Default {        
                $WorkflowType = "Custom"
                if($list.ItemCount -gt 0){
                    [xml]$xmldocument =  $ActivewflowAssociation.SoapXml
                    $name = $xmldocument.FirstChild.GetAttribute("Name")
                    if (!$name) {$name = $xmldocument.WorkflowTemplates.WorkflowTemplate.Name}
                    $wfName = $name.Replace(" ", "%20")
                    $webRelativeFolder = "Workflows/" + $wfName
                    $xomlFileName = $wfName + ".xoml"
                    $wfFolder = $wflowAssociation.ParentWeb.GetFolder($webRelativeFolder)
                    
                    # Even if the association extists, the file might not.
                    foreach($folderItem in $wfFolder){
                        if($folderItem.Exists){
                            $wflowAssociationFileExist = "Yes"
                        }
                        else{
                            $wflowAssociationFileExist = "No"
                        }
                    }

                    # Grab the workflow definition file
                    $xomlFile = $wfFolder.Files[$xomlFileName]
                    if ($xomlFile.Exists){
                        $xomlStream = $xomlFile.OpenBinaryStream()
                        $xmldocument.Load($xomlStream)
                        $xomlStream.Close()
                        #$filepath = $tempPath + "\" + $xomlFileName
                        $gotxoml = $xmldocument.OuterXml
                        #$gotxoml | Out-File $filepath -Append
                        #Write-Host " = Found XOML and output it to here:" $filepath -f Green
            
                
                        # Count number of Variables, seems like 8 system variables is a pattern, so subtracting that.
                        $XOMLVariables = $gotxoml.Split([string[]]@('ns0:WorkflowDataField'),[StringSplitOptions]"None").Count - 1
                        $XOMLVariables -= 8
                        #Write-Host " = Variable Count:" $XOMLVariables -f Magenta         

                        # Count number of Steps, remove any SequenceActivity nested inside TaskProcess
                        $XOMLSteps =  $gotxoml.Split([string[]]@('<SequenceActivity Description='),[StringSplitOptions]"None").Count - 1
                        $XOMLSteps -=  $gotxoml.Split([string[]]@('<SequenceActivity Description="When the Item'),[StringSplitOptions]"None").Count - 1
                        #Write-Host " = Step Count:" $XOMLSteps -f Magenta

                        # Count Activity: SetField
                        $XOMLActivitySetField = $gotxoml.Split([string[]]@('</ns0:SetFieldActivity>'),[StringSplitOptions]"None").Count - 1
                        $XOMLActivitySetField += $gotxoml.Split([string[]]@('</ns1:SetFieldActivity>'),[StringSplitOptions]"None").Count - 1
                        $XOMLActivitySetField += $gotxoml.Split([string[]]@('</ns2:SetFieldActivity>'),[StringSplitOptions]"None").Count - 1
                        #Write-Host " = Found SetField:" $XOMLActivitySetField -f Magenta

                        # Count Activity: SetVariable
                        $XOMLActivitySetVariable = $gotxoml.Split([string[]]@('</ns0:SetVariableActivity>'),[StringSplitOptions]"None").Count - 1
                        $XOMLActivitySetVariable += $gotxoml.Split([string[]]@('</ns1:SetVariableActivity>'),[StringSplitOptions]"None").Count - 1
                        $XOMLActivitySetVariable += $gotxoml.Split([string[]]@('</ns2:SetVariableActivity>'),[StringSplitOptions]"None").Count - 1
                        #Write-Host " = Found SetVariable:" $XOMLActivitySetVariable -f Magenta        
    
                        # Count Activity: If/Else
                        $XOMLActivityIfElse = $gotxoml.Split([string[]]@('</IfElseActivity>'),[StringSplitOptions]"None").Count - 1
                        #Write-Host " = Found If/Else:" $XOMLActivityIfElse -f Magenta        
                
                        # Count Activity: E-mail
                        $XOMLActivityEmail = $gotxoml.Split([string[]]@('ns0:EmailActivity'),[StringSplitOptions]"None").Count - 1
                        $XOMLActivityEmail += $gotxoml.Split([string[]]@('ns1:EmailActivity'),[StringSplitOptions]"None").Count - 1
                        $XOMLActivityEmail += $gotxoml.Split([string[]]@('ns2:EmailActivity'),[StringSplitOptions]"None").Count - 1
                        #Write-Host " = Found E-mail:" $XOMLActivityEmail -f Magenta      

                        # Count Activity: AddComment
                        $XOMLActivityAddComment = $gotxoml.Split([string[]]@('ns0:CommentActivity'),[StringSplitOptions]"None").Count - 1
                        $XOMLActivityAddComment += $gotxoml.Split([string[]]@('ns1:CommentActivity'),[StringSplitOptions]"None").Count - 1
                        $XOMLActivityAddComment += $gotxoml.Split([string[]]@('ns2:CommentActivity'),[StringSplitOptions]"None").Count - 1
                        #Write-Host " = Found AddComment:" $XOMLActivityAddComment -f Magenta      
    
                        # Count Activity: AddTime 
                        $XOMLActivityAddTime = $gotxoml.Split([string[]]@('ns0:AddTimeToDateActivity'),[StringSplitOptions]"None").Count - 1
                        $XOMLActivityAddTime += $gotxoml.Split([string[]]@('ns1:AddTimeToDateActivity'),[StringSplitOptions]"None").Count - 1
                        $XOMLActivityAddTime += $gotxoml.Split([string[]]@('ns2:AddTimeToDateActivity'),[StringSplitOptions]"None").Count - 1
                        #Write-Host " = Found AddTime:" $XOMLActivityAddTime -f Magenta      
    
                        # Count Activity: Calculate
                        $XOMLActivityCalculate = $gotxoml.Split([string[]]@('</ns0:MathActivity>'),[StringSplitOptions]"None").Count - 1
                        $XOMLActivityCalculate += $gotxoml.Split([string[]]@('</ns1:MathActivity>'),[StringSplitOptions]"None").Count - 1
                        $XOMLActivityCalculate += $gotxoml.Split([string[]]@('</ns2:MathActivity>'),[StringSplitOptions]"None").Count - 1
                        #Write-Host " = Found Calculate:" $XOMLActivityCalculate -f Magenta      
    
                        # Count Activity: BuidString (Custom?)
                        $XOMLActivityBuidString = $gotxoml.Split([string[]]@('<ns0:BuildStringActivity'),[StringSplitOptions]"None").Count - 1
                        $XOMLActivityBuidString += $gotxoml.Split([string[]]@('<ns1:BuildStringActivity'),[StringSplitOptions]"None").Count - 1
                        $XOMLActivityBuidString += $gotxoml.Split([string[]]@('<ns2:BuildStringActivity'),[StringSplitOptions]"None").Count - 1
                        #Write-Host " = Found BuidString:" $XOMLActivityBuidString -f Magenta      
    
                        # Count Activity: PauseFor
                        $XOMLActivityPauseFor = $gotxoml.Split([string[]]@('<ns0:DelayForActivity'),[StringSplitOptions]"None").Count - 1
                        $XOMLActivityPauseFor += $gotxoml.Split([string[]]@('<ns1:DelayForActivity'),[StringSplitOptions]"None").Count - 1
                        $XOMLActivityPauseFor += $gotxoml.Split([string[]]@('<ns2:DelayForActivity'),[StringSplitOptions]"None").Count - 1
                        #Write-Host " = Found PauseFor:" $XOMLActivityPauseFor -f Magenta      
    
                        # Count Activity: PauseUntil
                        $XOMLActivityPauseUntil = $gotxoml.Split([string[]]@('ns0:DelayUntilActivity'),[StringSplitOptions]"None").Count - 1
                        $XOMLActivityPauseUntil += $gotxoml.Split([string[]]@('ns1:DelayUntilActivity'),[StringSplitOptions]"None").Count - 1
                        $XOMLActivityPauseUntil += $gotxoml.Split([string[]]@('ns2:DelayUntilActivity'),[StringSplitOptions]"None").Count - 1
                        #Write-Host " = Found PauseUntil:" $XOMLActivityPauseUntil -f Magenta      
    
                        # Count Activity: SetTime
                        $XOMLActivitySetTime = $gotxoml.Split([string[]]@('ns0:SetTimeFieldActivity'),[StringSplitOptions]"None").Count - 1
                        $XOMLActivitySetTime += $gotxoml.Split([string[]]@('ns1:SetTimeFieldActivity'),[StringSplitOptions]"None").Count - 1
                        $XOMLActivitySetTime += $gotxoml.Split([string[]]@('ns2:SetTimeFieldActivity'),[StringSplitOptions]"None").Count - 1
                        #Write-Host " = Found SetTime:" $XOMLActivitySetTime -f Magenta      
    
                        # Count Activity: DocSetCapture
                        $XOMLActivityDocSetCapture = $gotxoml.Split([string[]]@('ns0:DocSetSnapshotActivity'),[StringSplitOptions]"None").Count - 1
                        $XOMLActivityDocSetCapture += $gotxoml.Split([string[]]@('ns1:DocSetSnapshotActivity'),[StringSplitOptions]"None").Count - 1
                        $XOMLActivityDocSetCapture += $gotxoml.Split([string[]]@('ns2:DocSetSnapshotActivity'),[StringSplitOptions]"None").Count - 1
                        #Write-Host " = Found DocSetCapture:" $XOMLActivityDocSetCapture -f Magenta      
                
                        # Count Activity: DocSetSubmit
                        $XOMLActivityDocSetSubmit = $gotxoml.Split([string[]]@('ns0:SubmitDocSetActivity'),[StringSplitOptions]"None").Count - 1
                        $XOMLActivityDocSetSubmit += $gotxoml.Split([string[]]@('ns1:SubmitDocSetActivity'),[StringSplitOptions]"None").Count - 1
                        $XOMLActivityDocSetSubmit += $gotxoml.Split([string[]]@('ns2:SubmitDocSetActivity'),[StringSplitOptions]"None").Count - 1
                        #Write-Host " = Found DocSetSubmit:" $XOMLActivityDocSetSubmit -f Magenta      
    
                        # Count Activity: DocSetApproval
                        $XOMLActivityDocSetApproval = $gotxoml.Split([string[]]@('ns0:SetDocSetModerationStatusActivity'),[StringSplitOptions]"None").Count - 1
                        $XOMLActivityDocSetApproval += $gotxoml.Split([string[]]@('ns1:SetDocSetModerationStatusActivity'),[StringSplitOptions]"None").Count - 1
                        $XOMLActivityDocSetApproval += $gotxoml.Split([string[]]@('ns2:SetDocSetModerationStatusActivity'),[StringSplitOptions]"None").Count - 1
                        #Write-Host " = Found DocSetApproval:" $XOMLActivityDocSetApproval -f Magenta      
    
                        # Count Activity: DocSetStartProcess
                        $XOMLActivityDocSetStartProcess = $gotxoml.Split([string[]]@('</ns0:DocSetContentsTask>'),[StringSplitOptions]"None").Count - 1
                        $XOMLActivityDocSetStartProcess += $gotxoml.Split([string[]]@('</ns1:DocSetContentsTask>'),[StringSplitOptions]"None").Count - 1
                        $XOMLActivityDocSetStartProcess += $gotxoml.Split([string[]]@('</ns2:DocSetContentsTask>'),[StringSplitOptions]"None").Count - 1
                        #Write-Host " = Found DocSetStartProcess:" $XOMLActivityDocSetStartProcess -f Magenta      
    
                        # Count Activity: ListCheckInItem
                        $XOMLActivityListCheckInItem = $gotxoml.Split([string[]]@('ns0:CheckInItemActivity'),[StringSplitOptions]"None").Count - 1
                        $XOMLActivityListCheckInItem += $gotxoml.Split([string[]]@('ns1:CheckInItemActivity'),[StringSplitOptions]"None").Count - 1
                        $XOMLActivityListCheckInItem += $gotxoml.Split([string[]]@('ns2:CheckInItemActivity'),[StringSplitOptions]"None").Count - 1
                        #Write-Host " = Found ListCheckInItem:" $XOMLActivityListCheckInItem -f Magenta      
    
                        # Count Activity: ListCheckOutItem
                        $XOMLActivityListCheckOutItem = $gotxoml.Split([string[]]@('ns0:CheckOutItemActivity'),[StringSplitOptions]"None").Count - 1
                        $XOMLActivityListCheckOutItem += $gotxoml.Split([string[]]@('ns1:CheckOutItemActivity'),[StringSplitOptions]"None").Count - 1
                        $XOMLActivityListCheckOutItem += $gotxoml.Split([string[]]@('ns2:CheckOutItemActivity'),[StringSplitOptions]"None").Count - 1
                        #Write-Host " = Found ListCheckOutItem:" $XOMLActivityListCheckOutItem -f Magenta      
    
                        # Count Activity: ListCopyItem
                        $XOMLActivityListCopyItem = $gotxoml.Split([string[]]@('ns0:CopyItemActivity'),[StringSplitOptions]"None").Count - 1
                        $XOMLActivityListCopyItem += $gotxoml.Split([string[]]@('ns1:CopyItemActivity'),[StringSplitOptions]"None").Count - 1
                        $XOMLActivityListCopyItem += $gotxoml.Split([string[]]@('ns2:CopyItemActivity'),[StringSplitOptions]"None").Count - 1
                        #Write-Host " = Found ListCopyItem:" $XOMLActivityListCopyItem -f Magenta      
    
                        # Count Activity: ListCreateItem
                        $XOMLActivityListCreateItem = $gotxoml.Split([string[]]@('ns0:CreateItemActivity'),[StringSplitOptions]"None").Count - 1
                        $XOMLActivityListCreateItem += $gotxoml.Split([string[]]@('ns1:CreateItemActivity'),[StringSplitOptions]"None").Count - 1
                        $XOMLActivityListCreateItem += $gotxoml.Split([string[]]@('ns2:CreateItemActivity'),[StringSplitOptions]"None").Count - 1
                        #Write-Host " = Found ListCreateItem:" $XOMLActivityListCreateItem -f Magenta      
    
                        # Count Activity: ListDeleteItem
                        $XOMLActivityListDeleteItem = $gotxoml.Split([string[]]@('ns0:DeleteItemActivity'),[StringSplitOptions]"None").Count - 1
                        $XOMLActivityListDeleteItem += $gotxoml.Split([string[]]@('ns1:DeleteItemActivity'),[StringSplitOptions]"None").Count - 1
                        $XOMLActivityListDeleteItem += $gotxoml.Split([string[]]@('ns2:DeleteItemActivity'),[StringSplitOptions]"None").Count - 1
                        #Write-Host " = Found ListDeleteItem:" $XOMLActivityListDeleteItem -f Magenta      
    
                        # Count Activity: ListDiscardCheckout
                        $XOMLActivityListDiscardCheckout = $gotxoml.Split([string[]]@('ns0:UndoCheckOutItemActivity'),[StringSplitOptions]"None").Count - 1
                        $XOMLActivityListDiscardCheckout += $gotxoml.Split([string[]]@('ns1:UndoCheckOutItemActivity'),[StringSplitOptions]"None").Count - 1
                        $XOMLActivityListDiscardCheckout += $gotxoml.Split([string[]]@('ns2:UndoCheckOutItemActivity'),[StringSplitOptions]"None").Count - 1
                        #Write-Host " = Found ListDiscardCheckout:" $XOMLActivityListDiscardCheckout -f Magenta      
    
                        # Count Activity: ListUpdateItem
                        $XOMLActivityListUpdateItem = $gotxoml.Split([string[]]@('ns0:UpdateItemActivity'),[StringSplitOptions]"None").Count - 1
                        $XOMLActivityListUpdateItem += $gotxoml.Split([string[]]@('ns1:UpdateItemActivity'),[StringSplitOptions]"None").Count - 1
                        $XOMLActivityListUpdateItem += $gotxoml.Split([string[]]@('ns2:UpdateItemActivity'),[StringSplitOptions]"None").Count - 1
                        #Write-Host " = Found ListUpdateItem:" $XOMLActivityListUpdateItem -f Magenta    
    
                        # Count Activity: TaskAssignGroup
                        $XOMLActivityTaskAssignGroup = $gotxoml.Split([string[]]@('ns0:GroupAssignedTask'),[StringSplitOptions]"None").Count - 1
                        $XOMLActivityTaskAssignGroup += $gotxoml.Split([string[]]@('ns1:GroupAssignedTask'),[StringSplitOptions]"None").Count - 1
                        $XOMLActivityTaskAssignGroup += $gotxoml.Split([string[]]@('ns2:GroupAssignedTask'),[StringSplitOptions]"None").Count - 1
                        #Write-Host " = Found TaskAssignGroup:" $XOMLActivityTaskAssignGroup -f Magenta      
    
                        # Count Activity: TaskAssignUser
                        $XOMLActivityTaskAssignUser = $gotxoml.Split([string[]]@('ns0:TodoItemTask'),[StringSplitOptions]"None").Count - 1
                        $XOMLActivityTaskAssignUser += $gotxoml.Split([string[]]@('ns1:TodoItemTask'),[StringSplitOptions]"None").Count - 1
                        $XOMLActivityTaskAssignUser += $gotxoml.Split([string[]]@('ns2:TodoItemTask'),[StringSplitOptions]"None").Count - 1
                        #Write-Host " = Found TaskAssignUser:" $XOMLActivityTaskAssignUser -f Magenta    
    
                        # Count Activity: TaskCollectData
                        $XOMLActivityTaskCollectData = $gotxoml.Split([string[]]@('ns0:TodoItemTask'),[StringSplitOptions]"None").Count - 1
                        $XOMLActivityTaskCollectData += $gotxoml.Split([string[]]@('ns1:TodoItemTask'),[StringSplitOptions]"None").Count - 1
                        $XOMLActivityTaskCollectData += $gotxoml.Split([string[]]@('ns2:TodoItemTask'),[StringSplitOptions]"None").Count - 1
                        #Write-Host " = Found TaskCollectData:" $XOMLActivityTaskCollectData -f Magenta      
    
                        # Count Activity: TaskStartProcess
                        $XOMLActivityTaskStartProcess = $gotxoml.Split([string[]]@('</ns0:ApprovalTaskProcess>'),[StringSplitOptions]"None").Count - 1
                        $XOMLActivityTaskStartProcess += $gotxoml.Split([string[]]@('</ns1:ApprovalTaskProcess>'),[StringSplitOptions]"None").Count - 1
                        $XOMLActivityTaskStartProcess += $gotxoml.Split([string[]]@('</ns2:ApprovalTaskProcess>'),[StringSplitOptions]"None").Count - 1
                        #Write-Host " = Found TaskStartProcess:" $XOMLActivityTaskStartProcess -f Magenta    
    
                        # Count Activity: TaskStartCustomProcess
                        $XOMLActivityTaskStartCustomProcess = $gotxoml.Split([string[]]@('</ns0:OfficeTask>'),[StringSplitOptions]"None").Count - 1
                        $XOMLActivityTaskStartCustomProcess += $gotxoml.Split([string[]]@('</ns1:OfficeTask>'),[StringSplitOptions]"None").Count - 1
                        $XOMLActivityTaskStartCustomProcess += $gotxoml.Split([string[]]@('</ns2:OfficeTask>'),[StringSplitOptions]"None").Count - 1
                        #Write-Host " = Found TaskStartCustomProcess:" $XOMLActivityTaskStartCustomProcess -f Magenta      
    
                        # Count Activity: TaskStartFeedbackProcess
                        $XOMLActivityTaskStartFeedbackProcess = $gotxoml.Split([string[]]@('</ns0:CollectFeedbackTaskProcess>'),[StringSplitOptions]"None").Count - 1                  
                        $XOMLActivityTaskStartFeedbackProcess += $gotxoml.Split([string[]]@('</ns1:CollectFeedbackTaskProcess>'),[StringSplitOptions]"None").Count - 1  
                        $XOMLActivityTaskStartFeedbackProcess += $gotxoml.Split([string[]]@('</ns2:CollectFeedbackTaskProcess>'),[StringSplitOptions]"None").Count - 1  
                        #Write-Host " = Found TaskStartFeedbackProcess:" $XOMLActivityTaskStartFeedbackProcess -f Magenta

                        # Count Activity: StringFromEnd, StringFromIndex, StringFromStart, StringLenght
                        $XOMLActivityStringFunctions = $gotxoml.Split([string[]]@('ns0:ExtractSubstringFrom'),[StringSplitOptions]"None").Count - 1  
                        $XOMLActivityStringFunctions += $gotxoml.Split([string[]]@('ns1:ExtractSubstringFrom'),[StringSplitOptions]"None").Count - 1  
                        $XOMLActivityStringFunctions += $gotxoml.Split([string[]]@('ns2:ExtractSubstringFrom'),[StringSplitOptions]"None").Count - 1  
                        #Write-Host " = Found StringFunctions:" $XOMLActivityStringFunctions -f Magenta    
    
                        # Count Activity: FindBetweenDates
                        $XOMLActivityFindBetweenDates = $gotxoml.Split([string[]]@('ns0:DateIntervalActivity'),[StringSplitOptions]"None").Count - 1 
                        $XOMLActivityFindBetweenDates += $gotxoml.Split([string[]]@('ns1:DateIntervalActivity'),[StringSplitOptions]"None").Count - 1 
                        $XOMLActivityFindBetweenDates += $gotxoml.Split([string[]]@('ns2:DateIntervalActivity'),[StringSplitOptions]"None").Count - 1 
                        #Write-Host " = Found FindBetweenDates:" $XOMLActivityFindBetweenDates -f Magenta    
    
                        # Total Activity Count
                        $XOMLActivityTotalCount += $XOMLActivitySetField
                        $XOMLActivityTotalCount += $XOMLActivitySetVariable
                        $XOMLActivityTotalCount += $XOMLActivityIfElse
                        $XOMLActivityTotalCount += $XOMLActivityEmail
                        $XOMLActivityTotalCount += $XOMLActivityAddComment
                        $XOMLActivityTotalCount += $XOMLActivityAddTime
                        $XOMLActivityTotalCount += $XOMLActivityCalculate
                        $XOMLActivityTotalCount += $XOMLActivityBuidString
                        $XOMLActivityTotalCount += $XOMLActivityPauseFor
                        $XOMLActivityTotalCount += $XOMLActivityPauseUntil
                        $XOMLActivityTotalCount += $XOMLActivitySetTime
                        $XOMLActivityTotalCount += $XOMLActivityDocSetCapture
                        $XOMLActivityTotalCount += $XOMLActivityDocSetSubmit
                        $XOMLActivityTotalCount += $XOMLActivityDocSetApproval
                        $XOMLActivityTotalCount += $XOMLActivityDocSetStartProcess
                        $XOMLActivityTotalCount += $XOMLActivityListCheckInItem
                        $XOMLActivityTotalCount += $XOMLActivityListCheckOutItem
                        $XOMLActivityTotalCount += $XOMLActivityListCopyItem
                        $XOMLActivityTotalCount += $XOMLActivityListCreateItem
                        $XOMLActivityTotalCount += $XOMLActivityListDeleteItem
                        $XOMLActivityTotalCount += $XOMLActivityListDiscardCheckout
                        $XOMLActivityTotalCount += $XOMLActivityListUpdateItem
                        $XOMLActivityTotalCount += $XOMLActivityTaskAssignGroup
                        $XOMLActivityTotalCount += $XOMLActivityTaskAssignUser
                        $XOMLActivityTotalCount += $XOMLActivityTaskCollectData
                        $XOMLActivityTotalCount += $XOMLActivityTaskStartProcess
                        $XOMLActivityTotalCount += $XOMLActivityTaskStartCustomProcess
                        $XOMLActivityTotalCount += $XOMLActivityTaskStartFeedbackProcess
                        $XOMLActivityTotalCount += $XOMLActivityStringFunctions
                        $XOMLActivityTotalCount += $XOMLActivityFindBetweenDates

                        Write-Host ("      Total Activity Count: {0}" -f $XOMLActivityTotalCount) -ForegroundColor DarkGray
                    }

                    else{
                        Write-Host "         ERROR: Didn't find the XOML file:" $xomlFileName "in the web:" $WF2010.WebUrl  -f Red    
                        #$ActivewflowAssociation.SoapXml   
                    }

                }
            }
        }
        Write-Host ("      Detected Workflow Type: {0}" -f $WorkflowType) -ForegroundColor DarkGray                      
                    
        # Add to Array
        $itemArrayWorkflow = New-Object PSObject
        $itemArrayWorkflow | Add-Member -type NoteProperty -Name 'SiteURL' -Value $WF2010.SiteURL
        $itemArrayWorkflow | Add-Member -type NoteProperty -Name 'WebURL' -Value $WF2010.WebUrl
        $itemArrayWorkflow | Add-Member -type NoteProperty -Name 'ListURL' -Value $WF2010.ListUrl
        $itemArrayWorkflow | Add-Member -type NoteProperty -Name 'ListTitle' -Value ($WF2010.ListTitle -replace $specialCharacters,'')
        $itemArrayWorkflow | Add-Member -type NoteProperty -Name 'ListItemCount' -Value $list.ItemCount
        $itemArrayWorkflow | Add-Member -type NoteProperty -Name 'ListLastModified' -Value $list.LastItemModifiedDate
        $itemArrayWorkflow | Add-Member -type NoteProperty -Name 'ListItemLastModified' -Value $LastModifiedItemDate
        $itemArrayWorkflow | Add-Member -type NoteProperty -Name 'WorkflowName' -Value ($WF2010.WorkflowName -replace $specialCharacters,'')
        $itemArrayWorkflow | Add-Member -type NoteProperty -Name 'WorkflowType' -Value $WorkflowType
        $itemArrayWorkflow | Add-Member -type NoteProperty -Name 'Scope' -Value $WF2010.Scope
        $itemArrayWorkflow | Add-Member -type NoteProperty -Name 'RunningInstances' -Value $WF2010.RunningInstances
        $itemArrayWorkflow | Add-Member -type NoteProperty -Name 'IsReusable' -Value $WF2010.IsReusable
        $itemArrayWorkflow | Add-Member -type NoteProperty -Name 'ContentTypeName' -Value ($WF2010.ContentTypeName -replace $specialCharacters,'')
        $itemArrayWorkflow | Add-Member -type NoteProperty -Name 'ScanID' -Value $SMATScanID
        $itemArrayWorkflow | Add-Member -type NoteProperty -Name 'TaskListRootFolder' -Value $TaskList.RootFolder
        $itemArrayWorkflow | Add-Member -type NoteProperty -Name 'TaskListItemCount' -Value $TaskList.ItemCount
        $itemArrayWorkflow | Add-Member -type NoteProperty -Name 'TaskListLastModified' -Value $TaskList.LastItemModifiedDate
        $itemArrayWorkflow | Add-Member -type NoteProperty -Name 'TaskListItemLastModified' -Value $TaskListItemLastModified
        $itemArrayWorkflow | Add-Member -type NoteProperty -Name 'HistoryListRootFolder' -Value $HistoryList.RootFolder
        $itemArrayWorkflow | Add-Member -type NoteProperty -Name 'HistoryListItemCount' -Value $HistoryList.ItemCount
        $itemArrayWorkflow | Add-Member -type NoteProperty -Name 'HistoryListLastModified' -Value $HistoryList.LastItemModifiedDate
        $itemArrayWorkflow | Add-Member -type NoteProperty -Name 'HistoryListItemLastModified' -Value $HistoryListItemLastModified
        $itemArrayWorkflow | Add-Member -type NoteProperty -Name 'WFAssociationFileExist' -Value $wflowAssociationFileExist
        $itemArrayWorkflow | Add-Member -type NoteProperty -Name 'XOMLVariables' -Value $XOMLVariables
        $itemArrayWorkflow | Add-Member -type NoteProperty -Name 'XOMLSteps' -Value $XOMLSteps
        $itemArrayWorkflow | Add-Member -type NoteProperty -Name 'XOMLActivities' -Value $XOMLActivityTotalCount

        # Building an array for everything, adding...
        $WorkflowDetails2010 += $itemArrayWorkflow

        # Increment index with 1
        $Workflow2010Index ++

        # Set the oldListURL so that we can compare up-top when we open the web and site object
        $oldListURL = $WF2010.ListURL

        # Clear Variables that might exist  
        if(Test-Path variable:XOMLActivityTotalCount){Remove-Variable XOMLActivityTotalCount}
        if(Test-Path variable:XOMLVariables){Remove-Variable XOMLVariables}
        if(Test-Path variable:XOMLSteps){Remove-Variable XOMLSteps}
        if(Test-Path variable:xmldocument){Remove-Variable xmldocument}
        if(Test-Path variable:xomlFile){Remove-Variable xomlFile}     
        if(Test-Path variable:wflowAssociationFileExist){Remove-Variable wflowAssociationFileExist}
        if(Test-Path variable:wflowAssociation){Remove-Variable wflowAssociation}
        if(Test-Path variable:TaskListTitle){Remove-Variable TaskListTitle}
        if(Test-Path variable:HistoryListTitle){Remove-Variable HistoryListTitle}
        if(Test-Path variable:LastModifiedItemDate){Remove-Variable LastModifiedItemDate}
        if(Test-Path variable:HistoryListItemLastModified){Remove-Variable HistoryListItemLastModified}
        if(Test-Path variable:TaskListItemLastModified){Remove-Variable TaskListItemLastModified}
        
        # Optimize memory usage to write to CSV in batches if PS supports it
        if ($currentPSVersion -ge $batchPSVersion) {
            if(($Workflow2010Index % $batchExportThreshold) -eq 0){
                $WorkflowDetails2010 | Export-Csv -Path $outputWorkflowDetails2010 -NoTypeInformation -Append
                $itemArrayWorkflow = @()
                $WorkflowDetails2010 = @()
                Write-Host "      Batch-point detected and wrote results to disk."
            }
        }
    }

    # Write to CSV if PS Doesn't support batches
    if ($currentPSVersion -lt $batchPSVersion) {
        $WorkflowDetails2010 | Export-Csv -Path $outputWorkflowDetails2010 -NoTypeInformation
    } Elseif (($Workflow2010Index % $batchExportThreshold) -gt 0) {
        $WorkflowDetails2010 | Export-Csv -Path $outputWorkflowDetails2010 -NoTypeInformation -Append
        Write-Host "      Final batch-point detected and wrote results to disk."
    }
}
else {
	#Write-Host
	#throw ("The WF2010 Report file does not exist: {0}" -f $inputWF2010Report)	
    Write-Host ("{0} - WF2010 Report Missing: {1}" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"), $inputWF2010Report) -ForegroundColor Yellow
}


#endregion #>

#region 8.  Workflow 2013 Details
If (Test-Path ($inputWF2013Report)){
	$WF2013Report = Import-Csv $inputWF2013Report | Where-Object {$_.ListUrl -ne 'N/A'} | Select-Object 'WebApplicationUrl','SiteUrl','Scope','RunningInstances','WebUrl','ListTitle','ListUrl','IsReusable','WorkflowAssociationName','WorkflowName','WorkflowPublishedBy','WorkflowAssociationID'

    Write-Host
    Write-Host
    Write-Host ("{0} - Scanning 2013 Workflow: {1} workflows" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"),$WF2013Report.count) -ForegroundColor Magenta
    $WorkflowDetails2013 = @()
    $Workflow2013Index = 0

    foreach ($WF2013 in $WF2013Report) 
    {
        Write-Host ("   {0} - {1} of {2}" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"),$Workflow2013Index,$WF2013Report.Count)
        Write-Host ("      List: {0}" -f $WF2013.ListUrl) -ForegroundColor DarkGray
        Write-Host ("      Workflow: {0}" -f $WF2010.WorkflowName) -ForegroundColor DarkGray

        # Get the Site, Web, and List Objects - This is a very expensive call so we only want to do it if needed.
        If($oldListURL -ne $WF2013.ListUrl){
            
            # Dispose Objects (if they exist) since we are about to load a new one
            if($site){
                $site.Dispose()
                $web.Dispose()
            }
            # Create the site and web objects
            $site = New-Object Microsoft.SharePoint.SPSite($WF2013.ListUrl)
            $web = $site.OpenWeb()
        } 

        $parentListName = $WF2013.ListUrl.Replace($web.Url + "/","")
        $list = $web.lists | Where-Object {$parentListName -eq $_.RootFolder}
        #Write-Host " = List Item Count:" $list.itemcount -f Yellow

        # Get Last Modified from List Items (if there are any items there)
        if($list.ItemCount -gt 0){ 
            #Get the last modified List Items, need to save the date in a variable since it will fail on the Array if we refer to first item [0]
            $LastModifiedItems = $List.GetItems($queryLastItemModified)
            #Write-Host " = List ITEM Last Modified:" $LastModifiedItems[0]["Modified"] "(GOOD)" -f Yellow
            $LastModifiedItemDate = $LastModifiedItems[0]["Modified"]
        } 

        # Get Workflow Association for this specific 2013 workflow
        $wfm = New-object Microsoft.SharePoint.WorkflowServices.WorkflowServicesManager($web)
        $sub = $wfm.GetWorkflowSubscriptionService()
        $wflowAssociation = $sub.EnumerateSubscriptionsByList($list.ID) | Where-Object {$_.Id -eq $WF2013.WorkflowAssociationID}
        #Write-Host " = Found 2013 WF Association:" $wflowAssociation.name "|" $wflowAssociation.id  -f Yellow

        # Get the History List
        $HistoryListId = $wflowAssociation.GetProperty("HistoryListId")   
        $HistoryList = $web.Lists | Where-Object {$_.id -Eq $HistoryListId}
        #Write-Host " = History Title:" $HistoryList.Title -f Magenta
        #Write-Host " = History Root:" $HistoryList.RootFolder.Url -f Magenta
        #Write-Host " = History List Last Modified:" $HistoryList.LastItemModifiedDate "(BAD)" -f Magenta
        #Write-Host " = History Item Count:" $HistoryList.ItemCount -f Magenta

        if($HistoryList.ItemCount -gt 0){ 
            #Get the last modified History List Items, need to save the date in a variable since it will fail on the Array if we refer to first item [0]
            $HistoryListLastModifiedItems = $HistoryList.GetItems($queryLastItemModified)
            #Write-Host " = History List ITEM Last Modified:" $HistoryListLastModifiedItems[0]["Modified"] "(GOOD)" -f Magenta
            $HistoryListItemLastModified = $HistoryListLastModifiedItems[0]["Modified"]
        }

        # Get Task List
        $TaskListId = $wflowAssociation.GetProperty("TaskListId") 
        $TaskList = $web.Lists | Where-Object {$_.id -Eq $TaskListId}
        #Write-Host " = Task List Title:" $TaskList.Title -f Green
        #Write-Host " = Task List Root:" $TaskList.RootFolder.Url -f Green
        #Write-Host " = Task List List Last Modified:" $TaskList.LastItemModifiedDate "(BAD)" -f Green
        #Write-Host " = Task List Item Count:" $TaskList.ItemCount -f Green

        if($TaskList.ItemCount -gt 0){
            #Get the last modified Task List Items, need to save the date in a variable since it will fail on the Array if we refer to first item [0]
            $TaskListLastModifiedItems = $TaskList.GetItems($queryLastItemModified)
            #Write-Host " = Task List ITEM Last Modified:" $TaskListLastModifiedItems[0]["Modified"] "(GOOD)" -f Green
            $TaskListItemLastModified = $TaskListLastModifiedItems[0]["Modified"]
        }
                        
        # Get complexity stuff for 2013 workflow, as long as there are items in the list
        $WorkflowType = "Custom"
        
        if($list.ItemCount -gt 0){
            
            # Sometimes the Subscription Relative Path doesn't start with a '/' so we need to add that
            if($wflowAssociation.GetProperty("SharePointWorkflowContext.Subscription.Path").StartsWith('/')){
                $SubscriptionRelativePath = $wflowAssociation.GetProperty("SharePointWorkflowContext.Subscription.Path")
                #Write-Host "Subscription Relative Path looked good" -f Green
            }
            else{
                $SubscriptionRelativePath = "/" + $wflowAssociation.GetProperty("SharePointWorkflowContext.Subscription.Path")
                #Write-Host "Subscription Relative path needed fixin'." -f Green
            }
            
            $WFSubscriptionPath = $WF2013.WebApplicationUrl.Substring(0,$WF2013.WebApplicationUrl.Length-1) + $SubscriptionRelativePath + "/workflow.xaml"
            Write-Host ("      WF Subscription Path: {0}" -f $WFSubscriptionPath) -ForegroundColor DarkGray

            $filepath = $tempPath + "\" + $WF2013.WorkflowAssociationName + ".xaml"
            Write-Host $filepath -ForegroundColor DarkGray

            $XAMLfile=New-Object System.net.webclient
            $XAMLfile.UseDefaultCredentials = $true
            $XAMLfile.downloadfile($WFSubscriptionPath, $filepath)

            #Write-Host " = Found XAML and output it to here:" $filepath -f Green
            If (Test-Path ($filepath)){
                # Count number of Variables
                $XAMLVariables = Select-String -Path $filepath -pattern "<x:Property Name=" -AllMatches
                $XAMLVariableCount = $XAMLVariables.Matches.count
                #Write-Host " = Variable Count:" $XAMLVariableCount -f Cyan         

                # Count Activity: If/Else, need to get this number before Stages and Steps since it's used there.
                $XAMLActivityIf = Select-String -Path $filepath -pattern "</If.Then>" -AllMatches
                $XAMLActivityElse = Select-String -Path $filepath -pattern "</If.Else>" -AllMatches
                $XAMLActivityIfElse = $XAMLActivityIf.Matches.count
                $XAMLActivityIfElse += $XAMLActivityElse.Matches.count
                #Write-Host " = Found If/Else (final):" $XAMLActivityIfElse -f Cyan    

                # Count number of Stages & Steps, need to subtract all ElseIf since they also generage a "<Sequence DisplayName". 
                $XAMLSteps = Select-String -Path $filepath -pattern "<Sequence DisplayName" -AllMatches
                $XAMLStepsFinal = $XAMLSteps.Matches.count
                $XAMLStepsFinal -= $XAMLActivityIfElse
                #Write-Host " = Stages & Steps Count:" $XAMLSteps.Matches.count -f Cyan
                #Write-Host " = Stages & Steps Count (after subtracting ElseIf):" $XAMLStepsFinal -f Cyan
                # Broken flows return strange numbers so zeroing that out to avoid confustion.
                if($XAMLStepsFinal -lt 0){
                    $XAMLStepsFinal = "0"
                }

                # Count number of Loops
                $XAMLLoopsNTime = Select-String -Path $filepath -pattern "</local:LoopNTimes>" -AllMatches
                $XAMLLoopsCondition = Select-String -Path $filepath -pattern "</While.Condition>" -AllMatches        
                $XAMLActivityLoopsFinal = $XAMLLoopsNTime.Matches.count
                $XAMLActivityLoopsFinal += $XAMLLoopsCondition.Matches.count
                #Write-Host " = Found Loops (final):" $XAMLActivityLoopsFinal -f Cyan

                # Count Activity: StartWorkflow
                $XAMLActivityStartWF = Select-String -Path $filepath -pattern "<local:WorkflowInterop" -AllMatches
                #Write-Host " = Found StartWF:" $XAMLActivityStartWF.Matches.count -f Cyan

                # Count Activity: AddComment
                $XAMLActivityAddComment = Select-String -Path $filepath -pattern "<local:Comment" -AllMatches
                #Write-Host " = Found AddComment:" $XAMLActivityAddComment.Matches.count -f Cyan           
                
                # Count Activity: AddTime 
                $XAMLActivityAddTime = Select-String -Path $filepath -pattern "</p:AddToDate>" -AllMatches
                #Write-Host " = Found AddTime:" $XAMLActivityAddTime.Matches.count -f Cyan      
        
                # Count Activity: BuildDictionary
                $XAMLActivityBuildDictionary = Select-String -Path $filepath -pattern "</p:BuildDynamicValue>" -AllMatches
                #Write-Host " = Found BuildDictionary:" $XAMLActivityBuildDictionary.Matches.count -f Cyan

                # Count Activity: HTTPCall
                $XAMLActivityHTTPCall = Select-String -Path $filepath -pattern "</local:CallHTTPWebService>" -AllMatches
                #Write-Host " = Found HTTPCall:" $XAMLActivityHTTPCall.Matches.count -f Cyan

                # Count Activity: CountDictionaryItem
                $XAMLActivityCountDictionaryItem = Select-String -Path $filepath -pattern "</p:CountDynamicValueItems>" -AllMatches
                #Write-Host " = Found CountDictionaryItem:" $XAMLActivityCountDictionaryItem.Matches.count -f Cyan
        
                # Count Activity: Calculate
                $XAMLActivityCalculate = Select-String -Path $filepath -pattern "</local:Calc>" -AllMatches
                #Write-Host " = Found Calculate:" $XAMLActivityCalculate.Matches.count -f Cyan      

                # Count Activity: GetDictionaryItem
                $XAMLActivityGetDictionaryItem = Select-String -Path $filepath -pattern "<p:GetDynamicValueProperty" -AllMatches
                #Write-Host " = Found GetDictionaryItem:" $XAMLActivityGetDictionaryItem.Matches.count -f Cyan

                # Count Activity: PauseFor
                $XAMLActivityPauseFor = Select-String -Path $filepath -pattern "<local:DelayFor " -AllMatches
                #Write-Host " = Found PauseFor:" $XAMLActivityPauseFor.Matches.count -f Cyan      
        
                # Count Activity: PauseUntil
                $XAMLActivityPauseUntil = Select-String -Path $filepath -pattern "<local:DelayUntil " -AllMatches
                #Write-Host " = Found PauseUntil:" $XAMLActivityPauseUntil.Matches.count -f Cyan      
        
                # Count Activity: E-mail
                $XAMLActivityEmail = Select-String -Path $filepath -pattern "<local:Email " -AllMatches
                #Write-Host " = Found E-mail:" $XAMLActivityEmail.Matches.count -f Cyan      

                # Count Activity: SetTime
                $XAMLActivitySetTime = Select-String -Path $filepath -pattern "</local:SetTimeField>" -AllMatches
                #Write-Host " = Found SetTime:" $XAMLActivitySetTime.Matches.count -f Cyan      

                # Count Activity: SetVariable
                $XAMLActivitySetVariable = Select-String -Path $filepath -pattern "</Assign>" -AllMatches
                #Write-Host " = Found SetVariable:" $XAMLActivitySetVariable.Matches.count -f Cyan     

                # Count Activity: ListCheckInItem
                $XAMLActivityListCheckInItem = Select-String -Path $filepath -pattern "<local:CheckInItem " -AllMatches
                #Write-Host " = Found ListCheckInItem:" $XAMLActivityListCheckInItem.Matches.count -f Cyan      
        
                # Count Activity: ListCheckOutItem
                $XAMLActivityListCheckOutItem = Select-String -Path $filepath -pattern "<local:CheckOutItem " -AllMatches
                #Write-Host " = Found ListCheckOutItem:" $XAMLActivityListCheckOutItem.Matches.count -f Cyan      
        
                # Count Activity: ListCopyItem
                $XAMLActivityListCopyItem = Select-String -Path $filepath -pattern "<local:CopyItem " -AllMatches
                #Write-Host " = Found ListCopyItem:" $XAMLActivityListCopyItem.Matches.count -f Cyan      
        
                # Count Activity: ListCreateItem
                $XAMLActivityListCreateItem = Select-String -Path $filepath -pattern "</local:CreateListItem>" -AllMatches
                #Write-Host " = Found ListCreateItem:" $XAMLActivityListCreateItem.Matches.count -f Cyan      
        
                # Count Activity: ListDeleteItem
                $XAMLActivityListDeleteItem = Select-String -Path $filepath -pattern "<local:DeleteListItem " -AllMatches
                #Write-Host " = Found ListDeleteItem:" $XAMLActivityListDeleteItem.Matches.count -f Cyan      
        
                # Count Activity: ListDiscardCheckout
                $XAMLActivityListDiscardCheckout = Select-String -Path $filepath -pattern "<local:UndoCheckOutItem " -AllMatches
                #Write-Host " = Found ListDiscardCheckout:" $XAMLActivityListDiscardCheckout.Matches.count -f Cyan      
        
                # Count Activity: SetField
                $XAMLActivitySetField = Select-String -Path $filepath -pattern "<local:SetField " -AllMatches
                #Write-Host " = Found SetField:" $XAMLActivitySetField.Matches.count -f Cyan

                # Count Activity: TranslateDoc
                $XAMLActivityTranslateDoc = Select-String -Path $filepath -pattern "<local:SetField " -AllMatches
                #Write-Host " = Found TranslateDoc:" $XAMLActivityTranslateDoc.Matches.count -f Cyan

                # Count Activity: ListUpdateItem
                $XAMLActivityListUpdateItem = Select-String -Path $filepath -pattern "<local:UpdateListItem " -AllMatches
                #Write-Host " = Found ListUpdateItem:" $XAMLActivityListUpdateItem.Matches.count -f Cyan    
        
                # Count Activity: WaitForItem
                $XAMLActivityWaitForItem = Select-String -Path $filepath -pattern "</local:WaitForItemEvent>" -AllMatches
                #Write-Host " = Found WaitForItem:" $XAMLActivityWaitForItem.Matches.count -f Cyan           

                # Count Activity: WaitForField
                $XAMLActivityWaitForField = Select-String -Path $filepath -pattern "</local:WaitForFieldChange>" -AllMatches
                #Write-Host " = Found WaitForField:" $XAMLActivityWaitForField.Matches.count -f Cyan       
                
                # Count Activity: TaskAssign
                $XAMLActivityTaskAssign = Select-String -Path $filepath -pattern "</local:SingleTask>" -AllMatches
                #Write-Host " = Found TaskAssign:" $XAMLActivityTaskAssign.Matches.count -f Cyan    
        
                # Count Activity: TaskStartProcess
                $XAMLActivityTaskStartProcess = Select-String -Path $filepath -pattern "</local:CompositeTask>" -AllMatches
                #Write-Host " = Found TaskStartProcess:" $XAMLActivityTaskStartProcess.Matches.count -f Cyan    
            
                # Count Activity: StringFromEnd
                $XAMLActivityStringFromEnd= Select-String -Path $filepath -pattern "</local:ExtractSubstringFromEnd>" -AllMatches
                #Write-Host " = Found StringFromEnd:" $XAMLActivityStringFromEnd.Matches.count -f Cyan    

                # Count Activity: StringFromIndex
                $XAMLActivityStringFromIndex = Select-String -Path $filepath -pattern "</local:ExtractSubstringFromIndex>" -AllMatches 
                #Write-Host " = Found StringFromIndex:" $XAMLActivityStringFromIndex.Matches.count -f Cyan    
                    
                # Count Activity: StringFromStart
                $XAMLActivityStringFromStart = Select-String -Path $filepath -pattern "</local:ExtractSubstringFromStart>" -AllMatches
                #Write-Host " = Found StringFromStart:" $XAMLActivityStringFromStart.Matches.count -f Cyan    
            
                # Count Activity: StringFromIndexLenght
                $XAMLActivityStringFromIndexLenght = Select-String -Path $filepath -pattern "</local:ExtractSubstringFromIndexLength>" -AllMatches
                #Write-Host " = Found StringFromIndexLenght:" $XAMLActivityStringFromIndexLenght.Matches.count -f Cyan    
        
                # Count Activity: FindBetweenDates
                $XAMLActivityFindBetweenDates = Select-String -Path $filepath -pattern "</local:DateInterval>" -AllMatches
                #Write-Host " = Found FindBetweenDates:" $XAMLActivityFindBetweenDates.Matches.count -f Cyan    
        
                # Count Activity: FindInString
                $XAMLActivityFindInString = Select-String -Path $filepath -pattern "</p:IndexOfString>" -AllMatches
                #Write-Host " = Found FindInString:" $XAMLActivityFindInString.Matches.count -f Cyan    
                
                # Count Activity: ReplaceString
                $XAMLActivityReplaceString = Select-String -Path $filepath -pattern "</p:ReplaceString>" -AllMatches
                #Write-Host " = Found ReplaceString:" $XAMLActivityReplaceString.Matches.count -f Cyan    
                            
                # Count Activity: StringTrim
                $XAMLActivityStringTrim = Select-String -Path $filepath -pattern "</p:Trim>" -AllMatches
                #Write-Host " = Found StringTrim:" $XAMLActivityStringTrim.Matches.count -f Cyan    
        

                # Total Activity Count
                $XAMLActivityTotalCount += $XAMLActivityIfElse
                $XAMLActivityTotalCount += $XAMLActivityLoopsFinal
                $XAMLActivityTotalCount += $XAMLActivityStartWF.Matches.count
                $XAMLActivityTotalCount += $XAMLActivityAddComment.Matches.count
                $XAMLActivityTotalCount += $XAMLActivityAddTime.Matches.count
                $XAMLActivityTotalCount += $XAMLActivityBuildDictionary.Matches.count
                $XAMLActivityTotalCount += $XAMLActivityHTTPCall.Matches.count
                $XAMLActivityTotalCount += $XAMLActivityCountDictionaryItem.Matches.count
                $XAMLActivityTotalCount += $XAMLActivityCalculate.Matches.count
                $XAMLActivityTotalCount += $XAMLActivityGetDictionaryItem.Matches.count
                $XAMLActivityTotalCount += $XAMLActivityPauseFor.Matches.count
                $XAMLActivityTotalCount += $XAMLActivityPauseUntil.Matches.count
                $XAMLActivityTotalCount += $XAMLActivityEmail.Matches.count       
                $XAMLActivityTotalCount += $XAMLActivitySetTime.Matches.count
                $XAMLActivityTotalCount += $XAMLActivitySetVariable.Matches.count
                $XAMLActivityTotalCount += $XAMLActivityListCheckInItem.Matches.count
                $XAMLActivityTotalCount += $XAMLActivityListCheckOutItem.Matches.count
                $XAMLActivityTotalCount += $XAMLActivityListCopyItem.Matches.count
                $XAMLActivityTotalCount += $XAMLActivityListCreateItem.Matches.count
                $XAMLActivityTotalCount += $XAMLActivityListDeleteItem.Matches.count
                $XAMLActivityTotalCount += $XAMLActivityListDiscardCheckout.Matches.count
                $XAMLActivityTotalCount += $XAMLActivitySetField.Matches.count
                $XAMLActivityTotalCount += $XAMLActivityTranslateDoc.Matches.count
                $XAMLActivityTotalCount += $XAMLActivityListUpdateItem.Matches.count
                $XAMLActivityTotalCount += $XAMLActivityWaitForItem.Matches.count
                $XAMLActivityTotalCount += $XAMLActivityWaitForField.Matches.count 
                $XAMLActivityTotalCount += $XAMLActivityTaskAssign.Matches.count 
                $XAMLActivityTotalCount += $XAMLActivityTaskStartProcess.Matches.count
                $XAMLActivityTotalCount += $XAMLActivityStringFromEnd.Matches.count
                $XAMLActivityTotalCount += $XAMLActivityStringFromIndex.Matches.count
                $XAMLActivityTotalCount += $XAMLActivityStringFromStart.Matches.count
                $XAMLActivityTotalCount += $XAMLActivityStringFromIndexLenght.Matches.count
                $XAMLActivityTotalCount += $XAMLActivityFindBetweenDates.Matches.count
                $XAMLActivityTotalCount += $XAMLActivityFindInString.Matches.count
                $XAMLActivityTotalCount += $XAMLActivityReplaceString.Matches.count
                $XAMLActivityTotalCount += $XAMLActivityStringTrim.Matches.count
        
                Write-Host ("      Total Activity Count: {0}" -f $XAMLActivityTotalCount) -ForegroundColor DarkGray

                # Logic to delete the file once it has been analyzed   
                Remove-Item -Path $filepath -Force
            }
            else {
                Write-Host
                Write-Host ("      Unable to download XAML file: {0}" -f $WFSubscriptionPath) -ForegroundColor Red
            }
        }

        # Add to Array
        $itemArrayWorkflow = New-Object PSObject
        $itemArrayWorkflow | Add-Member -type NoteProperty -Name 'SiteURL' -Value $WF2013.SiteURL
        $itemArrayWorkflow | Add-Member -type NoteProperty -Name 'WebURL' -Value $WF2013.WebUrl
        $itemArrayWorkflow | Add-Member -type NoteProperty -Name 'ListURL' -Value $WF2013.ListUrl
        $itemArrayWorkflow | Add-Member -type NoteProperty -Name 'ListTitle' -Value ($WF2013.ListTitle -replace $specialCharacters,'')
        $itemArrayWorkflow | Add-Member -type NoteProperty -Name 'ListItemCount' -Value $list.ItemCount
        $itemArrayWorkflow | Add-Member -type NoteProperty -Name 'ListLastModified' -Value $list.LastItemModifiedDate
        $itemArrayWorkflow | Add-Member -type NoteProperty -Name 'ListItemLastModified' -Value $LastModifiedItemDate
        $itemArrayWorkflow | Add-Member -type NoteProperty -Name 'WorkflowName' -Value ($WF2013.WorkflowName -replace $specialCharacters,'')
        $itemArrayWorkflow | Add-Member -type NoteProperty -Name 'WorkflowAssociationName' -Value $WF2013.WorkflowAssociationName
        $itemArrayWorkflow | Add-Member -type NoteProperty -Name 'WorkflowModifiedBy' -Value $wflowAssociation.GetProperty("ModifiedBy")
        $itemArrayWorkflow | Add-Member -type NoteProperty -Name 'WorkflowModified' -Value $wflowAssociation.GetProperty("SharePointWorkflowContext.Subscription.ModifiedDate")
        $itemArrayWorkflow | Add-Member -type NoteProperty -Name 'Scope' -Value $WF2013.Scope
        $itemArrayWorkflow | Add-Member -type NoteProperty -Name 'RunningInstances' -Value $WF2013.RunningInstances
        $itemArrayWorkflow | Add-Member -type NoteProperty -Name 'IsReusable' -Value $WF2013.IsReusable
        $itemArrayWorkflow | Add-Member -type NoteProperty -Name 'ScanID' -Value $SMATScanID
        $itemArrayWorkflow | Add-Member -type NoteProperty -Name 'TaskListTitle' -Value ($TaskList.Title -replace $specialCharacters,'')
        $itemArrayWorkflow | Add-Member -type NoteProperty -Name 'TaskListRootFolder' -Value $TaskList.RootFolder.Url
        $itemArrayWorkflow | Add-Member -type NoteProperty -Name 'TaskListItemCount' -Value $TaskList.ItemCount
        $itemArrayWorkflow | Add-Member -type NoteProperty -Name 'TaskListLastModified' -Value $TaskList.LastItemModifiedDate
        $itemArrayWorkflow | Add-Member -type NoteProperty -Name 'TaskListItemLastModified' -Value $TaskListItemLastModified
        $itemArrayWorkflow | Add-Member -type NoteProperty -Name 'HistoryListTitle' -Value ($HistoryList.Title -replace $specialCharacters,'')
        $itemArrayWorkflow | Add-Member -type NoteProperty -Name 'HistoryListRootFolder' -Value $HistoryList.RootFolder.Url
        $itemArrayWorkflow | Add-Member -type NoteProperty -Name 'HistoryListItemCount' -Value $HistoryList.ItemCount
        $itemArrayWorkflow | Add-Member -type NoteProperty -Name 'HistoryListLastModified' -Value $HistoryList.LastItemModifiedDate
        $itemArrayWorkflow | Add-Member -type NoteProperty -Name 'HistoryListItemLastModified' -Value $HistoryListItemLastModified
        $itemArrayWorkflow | Add-Member -type NoteProperty -Name 'WorkflowType' -Value $WorkflowType
        $itemArrayWorkflow | Add-Member -type NoteProperty -Name 'XOMLVariables' -Value $XAMLVariableCount
        $itemArrayWorkflow | Add-Member -type NoteProperty -Name 'XOMLSteps' -Value $XAMLStepsFinal
        $itemArrayWorkflow | Add-Member -type NoteProperty -Name 'XOMLActivities' -Value $XAMLActivityTotalCount

        $WorkflowDetails2013 += $itemArrayWorkflow

        # Set the oldListURL so that we can compare up-top when we open the web and site object
        $oldListURL = $WF2013.ListURL

        # Increment index with 1
        $Workflow2013Index ++

        # Clear Variables
        if(Test-Path variable:XAMLVariableCount){Remove-Variable XAMLVariableCount}
        if(Test-Path variable:XAMLActivityTotalCount){Remove-Variable XAMLActivityTotalCount}
        if(Test-Path variable:XAMLActivityIfElse){Remove-Variable XAMLActivityIfElse}
        if(Test-Path variable:XAMLStepsFinal){Remove-Variable XAMLStepsFinal}
        if(Test-Path variable:TaskListTitle){Remove-Variable TaskListTitle}
        if(Test-Path variable:HistoryListTitle){Remove-Variable HistoryListTitle}
        if(Test-Path variable:LastModifiedItemDate){Remove-Variable LastModifiedItemDate}  
        if(Test-Path variable:TaskListItemLastModified){Remove-Variable TaskListItemLastModified}    
        if(Test-Path variable:HistoryListItemLastModified){Remove-Variable HistoryListItemLastModified}

        # Optimize memory usage to write to CSV in batches if PS supports it
        if ($currentPSVersion -ge $batchPSVersion) {
            if(($Workflow2013Index % $batchExportThreshold) -eq 0){
                $WorkflowDetails2013 | Export-Csv -Path $outputWorkflowDetails2013 -NoTypeInformation -Append
                $itemArrayWorkflow = @()
                $WorkflowDetails2013 = @()
                Write-Host "      Batch-point detected and wrote results to disk."
            }
        }
    }

    # Write to CSV if PS Doesn't support batches
    if ($currentPSVersion -lt $batchPSVersion) {
        $WorkflowDetails2013 | Export-Csv -Path $outputWorkflowDetails2013 -NoTypeInformation
    } Elseif (($Workflow2013Index % $batchExportThreshold) -gt 0) {
        $WorkflowDetails2013 | Export-Csv -Path $outputWorkflowDetails2013 -NoTypeInformation -Append
        Write-Host "      Final batch-point detected and wrote results to disk."
    }
}
else {
	Write-Host
    Write-Host ("{0} - WF2013 Report Missing: {1}" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"), $inputWF2013Report) -ForegroundColor Yellow
}
#endregion #>

#region 9.	Web Details
# Load Site Assessment file
If (Test-Path ($inputSiteAssessment)){
	$siteAssessment = Import-Csv $inputSiteAssessment

    # Loop through Site Collections
    Write-Host
    Write-Host ("{0} - Scanning Site Collections" -f (Get-Date -Format "yy-MM-dd HH:mm:ss")) -ForegroundColor Magenta
    $webDetails = @()
    $webIndex = 0

    foreach ($site in $siteAssessment) {
        Write-Host ("   {0} - Retrieving Webs: {1}" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"),$site.SiteUrl) -NoNewline
        $webs = New-Object System.Collections.Queue;
        (Get-SPSite $site.SiteUrl).AllWebs | ForEach-Object{$webs.Enqueue($_)} 
        If ($webs.Count -gt 0) {
            # Loop through Web Queue
            while($webs.count -ne 0) {
                $webs.Dequeue() | ForEach-Object{
                    $web = New-Object -TypeName PSObject -Property (@{ScanId=$SMATScanID;SiteId=$site.SiteId;SiteUrl=$site.SiteUrl;ID=$_.ID;Title=($_.Title -replace $specialCharacters,'');Url=$_.Url;Description=((($_.Description -replace $specialCharacters,'') -replace "`n",' ') -replace "`r",' ');
                        Created=$_.Created;WebTemplate=$_.WebTemplate;WebTemplateId=$_.WebTemplateId;Exists=$_.Exists;EmailInsertsEnabled=$_.EmailInsertsEnabled;EventHandlersEnabled=$_.EventHandlersEnabled;
                        IncludeSupportingFolders=$_.IncludeSupportingFolders;NoCrawl=$_.NoCrawl;SaveSiteAsTemplateEnabled=$_.SaveSiteAsTemplateEnabled;Name=$_.Name;IsMultilingual=$_.IsMultilingual;
                        OverwriteTranslationsOnChange=$_.OverwriteTranslationsOnChange;ParentWeb=$_.ParentWeb;Theme=$_.Theme;ThemeCssUrl=$_.ThemeCssUrl;ThemedCssFolderUrl=$_.ThemedCssFolderUrl;
                        AlternateCssUrl=$_.AlternateCssUrl;CustomJavaScriptFileUrl=$_.CustomJavaScriptFileUrl;AlternateHeader=$_.AlternateHeader;MasterUrl=$_.MasterUrl;CustomMasterUrl=$_.CustomMasterUrl;
                        SiteLogoUrl=$_.SiteLogoUrl;SiteLogoDescription=((($_.SiteLogoDescription -replace $specialCharacters,'') -replace "`n",' ') -replace "`r",' ');RequestAccessEmail=$_.RequestAccessEmail;RequestAccessEnabled=$_.RequestAccessEnabled;
                        SiteClientTag=$_.SiteClientTag;LastItemModifiedDate=$_.LastItemModifiedDate;DocumentLibraryCalloutOfficeWebAppPreviewersDisabled=$_.DocumentLibraryCalloutOfficeWebAppPreviewersDisabled;
                        RecycleBinEnabled=$_.RecycleBinEnabled;RequireDynamicCanary=$_.RequireDynamicCanary;Author=$_.Author;NonHostHeaderUrl=$_.NonHostHeaderUrl;ServerRelativeUrl=$_.ServerRelativeUrl;
                        IsRootWeb=$_.IsRootWeb;AppInstanceId=$_.AppInstanceId;AppDatabaseName=$_.AppDatabaseName;AppDatabaseServerReferenceId=$_.AppDatabaseServerReferenceId;
                        AppDatabaseTargetApplicationId=$_.AppDatabaseTargetApplicationId;AuthenticationMode=$_.AuthenticationMode;IsADAccountCreationMode=$_.IsADAccountCreationMode;
                        IsADEmailEnabled=$_.IsADEmailEnabled;MasterPageReferenceEnabled=$_.MasterPageReferenceEnabled;PresenceEnabled=$_.PresenceEnabled;ExcludeFromOfflineClient=$_.ExcludeFromOfflineClient;
                        EnableMinimalDownload=$_.EnableMinimalDownload;EffectivePresenceEnabled=$_.EffectivePresenceEnabled;SyndicationEnabled=$_.SyndicationEnabled;AllowRssFeeds=$_.AllowRssFeeds;
                        QuickLaunchEnabled=$_.QuickLaunchEnabled;HideSiteContentsLink=$_.HideSiteContentsLink;TreeViewEnabled=$_.TreeViewEnabled;ParserEnabled=$_.ParserEnabled;CacheAllSchema=$_.CacheAllSchema;
                        PortalMember=$_.PortalMember;HasExternalSecurityProvider=$_.HasExternalSecurityProvider;ExternalSecurityProviderSetting=$_.ExternalSecurityProviderSetting;Language=$_.Language;
                        UICulture=$_.UICulture;Configuration=$_.Configuration;Locale=$_.Locale;CurrencyLocaleID=$_.CurrencyLocaleID;Provisioned=$_.Provisioned;RegionalSettings=$_.RegionalSettings;
                        AllowUnsafeUpdates=$_.AllowUnsafeUpdates;AllWebTemplatesAllowed=$_.AllWebTemplatesAllowed;ExecuteUrl=$_.ExecuteUrl;PortalUrl=$_.PortalUrl;PortalName=$_.PortalName;
                        PortalSubscriptionUrl=$_.PortalSubscriptionUrl;CurrentChangeToken=$_.CurrentChangeToken;UIVersion=$_.UIVersion;UIVersionConfigurationEnabled=$_.UIVersionConfigurationEnabled;
                        ClientTag=$_.ClientTag;FileDialogPostProcessorId=$_.FileDialogPostProcessorId;CustomUploadPage=$_.CustomUploadPage;AllowDesignerForCurrentUser=$_.AllowDesignerForCurrentUser;
                        AllowRevertFromTemplateForCurrentUser=$_.AllowRevertFromTemplateForCurrentUser;AllowMasterPageEditingForCurrentUser=$_.AllowMasterPageEditingForCurrentUser;
                        ShowUrlStructureForCurrentUser=$_.ShowUrlStructureForCurrentUser;DesignerDownloadUrlForCurrentUser=$_.DesignerDownloadUrlForCurrentUser;AllowCreateDeclarativeWorkflowForCurrentUser=$_.AllowCreateDeclarativeWorkflowForCurrentUser;
                        AllowSavePublishDeclarativeWorkflowForCurrentUser=$_.AllowSavePublishDeclarativeWorkflowForCurrentUser;AllowSaveDeclarativeWorkflowAsTemplateForCurrentUser=$_.AllowSaveDeclarativeWorkflowAsTemplateForCurrentUser;
                        PushNotificationSubscribers=$_.PushNotificationSubscribers;IsAppWeb=$_.IsAppWeb;ParentWebId=$_.ParentWebId;SiteUserInfoList=$_.SiteUserInfoList;AllowAnonymousAccess=$_.AllowAnonymousAccess;
                        AnonymousState=$_.AnonymousState;HasUniqueRoleDefinitions=$_.HasUniqueRoleDefinitions;FirstUniqueRoleDefinitionWeb=$_.FirstUniqueRoleDefinitionWeb;FirstUniqueAncestorWeb=$_.FirstUniqueAncestorWeb;
                        AnonymousPermMask64=$_.AnonymousPermMask64;FirstUniqueAncestorSecurableObject=$_.FirstUniqueAncestorSecurableObject;AllowAutomaticASPXPageIndexing=$_.AllowAutomaticASPXPageIndexing;
                        ASPXPageIndexMode=$_.ASPXPageIndexMode;ASPXPageIndexed=$_.ASPXPageIndexed;AssociatedOwnerGroup=$_.AssociatedOwnerGroup;AsociatedMemberGroup=$_.AsociatedMemberGroup;
                        AssociatedVisitorGroup=$_.AssociatedVisitorGroup;HasUniquePerm=$_.HasUniquePerm;HasUniqueRoleAssignments=$_.HasUniqueRoleAssignments;FirstUniqueAncestor=$_.FirstUniqueAncestor})
                    $webDetails += $web
                    $webIndex ++
                    Remove-Variable web
                }
            }
            Write-Host " - Success" -ForegroundColor Green
        }
        Else {
            Write-Host " - No Webs Found" -ForegroundColor Yellow
        }

        # Cleaning up memory
        Remove-Variable webs

        # Optimize memory usage to write to CSV in batches if PS supports it
        if ($currentPSVersion -ge $batchPSVersion) {
            if(($webIndex % $batchExportThreshold) -eq 0){
                $webDetails | Export-Csv -Path $outputWebs -NoTypeInformation -Append
                $webDetails = @()
                Write-Host "      Batch-point detected and wrote results to disk."
            }
        }
    }

    # Write to CSV if PS Doesn't support batches
    if ($currentPSVersion -lt $batchPSVersion) {
        $webDetails | Export-Csv -Path $outputWebs -NoTypeInformation
    } Elseif (($webIndex % $batchExportThreshold) -gt 0) {
        $webDetails | Export-Csv -Path $outputWebs -NoTypeInformation -Append
        Write-Host "      Final batch-point detected and wrote results to disk."
    }

    # Cleaning up memory
    Remove-Variable siteAssessment
    Remove-Variable webDetails
}
else {
	Write-Host
    Write-Host ("{0} - Site Assessment Report Missing: {1}" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"), $inputSiteAssessment) -ForegroundColor Yellow
}
#endregion #>

#region 10.	Checked Out Documents
# Load Checked Out Files csv
If (Test-Path ($inputCheckedOutDocs)){
	$checkedOutItems = Import-Csv $inputCheckedOutDocs

    # Loop through Lists
    Write-Host
    Write-Host ("{0} - Scanning Checked Out Items" -f (Get-Date -Format "yy-MM-dd HH:mm:ss")) -ForegroundColor Magenta
    $checkedOutDetails = @()
    $checkedOutIndex = 0

    :checkedOutLoop foreach ($row in $checkedOutItems) {
        Write-Host ("   {0} - Retrieving item: {1}" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"),$row.File) -NoNewline
        try {
            $items = New-Object System.Collections.Queue
            $site = New-Object Microsoft.SharePoint.SPSite($row.File)
            $web = $site.OpenWeb()
            $item = $web.GetListItem($row.File)
            $item.File | ForEach-Object {$items.Enqueue($_)}
        }
        catch {
            Write-Host " - Item Not Found" -ForegroundColor Yellow
            Remove-Variable items
            Remove-Variable site
            Remove-Variable web
            Continue checkedOutLoop
        }
        If ($items.Count -gt 0) {
            while($items.count -ne 0) {
                $items.Dequeue() | ForEach-Object{
                    $obj = New-Object -TypeName PSObject -Property (@{ScanId=$SMATScanID;File=$row.File;ID=$item.ID;Folder=$item.Folder;HasUniqueRoleAssignments=$item.HasUniqueRoleAssignments;Name=($_.Name -replace $specialCharacters,'');
                        Title=($_.Title -replace $specialCharacters,'');TimeCreated=$_TimeCreated;TimeLastModified=$_.TimeLastModified;CheckOutStatus=$_.CheckOutStatus;CheckOutType=$_.CheckOutType;CheckedOutDate=$_.CheckedOutDate;
                        CheckOutExpires=$_.CheckOutExpires;CheckedOutBy=$_.CheckedOutBy;CheckedOutByUser=$_.CheckedOutByUser;CheckInComment=((($_.CheckInComment -replace $specialCharacters,'') -replace "`n",' ') -replace "`r",' ');EffectiveRawPermissions=$_.EffectiveRawPermissions;
                        EffectiveAuditMask=$_.EffectiveAuditMask;Exists=$_.Exists;Length=$_.Length;TotalLength=$_.TotalLength;LengthByUser=$_.LengthByUser;Author=$_.Author;ModifiedBy=$_.ModifiedBy;
                        CustomizedPageStatus=$_.CustomizedPageStatus;LockType=$_.LockType;LockId=$_.LockId;LockedDate=$_.LockedDate;LockExpires=$_.LockExpires;LockedByUser=$_.LockedByUser;Generator=$_.Generator;
                        CharSetName=$_.CharSetName;ProgID=$_.ProgID;UIVersion=$_.UIVersion;MajorVersion=$_.MajorVersion;MinorVersion=$_.MinorVersion;UIVersionLabel=$_.UIVersionLabel;ParentFolder=$_.ParentFolder;
                        Url=$_.Url;UniqueId=$_.UniqueId;ServerRelativeUrl=$_.ServerRelativeUrl;IconUrl=$_.IconUrl;RequiresCheckout=$_.RequiresCheckout;DocumentLibrary=$_.DocumentLibrary;Versions=$_.Versions;
                        InDocumentLibrary=$_.InDocumentLibrary;Level=$_.Level;IsConvertedFile=$_.IsConvertedFile;SourceLeafName=$_.SourceLeafName;SourceUIVersion=$_.SourceUIVersion;GeneratingConverterId=$_.GeneratingConverterId;
                        SourceFile=$_.SourceFile;Web=$_.Web;ServerRedirected=$_.ServerRedirected;IsIRMed=$_.IsIRMed;IsSharedAccessRequested=$_.IsSharedAccessRequested})
                    $checkedOutDetails += $obj
                    $checkedOutIndex ++
                    Remove-Variable obj
                }
            }
            Write-Host " - Success" -ForegroundColor Green
        }
        Else {
            Write-Host " - Item Not Found" -ForegroundColor Yellow
        }
        
        # Optimize memory usage to write to CSV in batches if PS supports it
        if ($currentPSVersion -ge $batchPSVersion) {
            if(($checkedOutIndex % $batchExportThreshold) -eq 0){
                $checkedOutDetails | Export-Csv -Path $outputCheckedOutDocs -NoTypeInformation -Append
                $checkedOutDetails = @()
                Write-Host "      Batch-point detected and wrote results to disk."
            }
        }

        # Clean Up Memory
        Remove-Variable items
        Remove-Variable site
        Remove-Variable web
        try {Remove-Variable item -ErrorAction Stop} catch {}
    }

    # Write to CSV if PS Doesn't support batches
    if ($currentPSVersion -lt $batchPSVersion) {
        $checkedOutDetails | Export-Csv -Path $outputCheckedOutDocs -NoTypeInformation
    } Elseif (($checkedOutIndex % $batchExportThreshold) -gt 0) {
        $checkedOutDetails | Export-Csv -Path $outputCheckedOutDocs -NoTypeInformation -Append
        Write-Host "      Final batch-point detected and wrote results to disk."
    }

    # Cleaning up memory
    Remove-Variable checkedOutItems
    Remove-Variable checkedOutDetails
}
else {
	Write-Host
    Write-Host ("{0} - Checked Out Files Report Missing: {1}" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"), $inputCheckedOutDocs) -ForegroundColor Yellow
}
#endregion #>

#region 11.	Scanning Lists
# Load Large Lists file
If (Test-Path ($inputLists)){	
	$lists = Import-Csv $inputLists | Where-Object ListUrl -NotMatch $skipListURLs

    # Looping through all lists
    Write-Host
    Write-Host ("{0} - Scanning Lists" -f (Get-Date -Format "yy-MM-dd HH:mm:ss")) -ForegroundColor Magenta
    $fieldDetails = @()
    $fieldIndex = 0
    $lastModifiedIndex = 0
    $lastModifiedDetails = @()
    foreach ($list in $lists) {
        # Retrieve List Object
        $web = Get-SPWeb -Identity $list.WebUrl
        $currentList = $web.Lists | where{$_.Title -eq $list.ListTitle}

        #region 11a.	Complex Columns
        Write-Host ("   {0} - Retrieving complex columns: {1}" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"),$list.ListUrl) -NoNewline
        $fields = New-Object System.Collections.Queue;	
        $currentList.Fields | where{($complexColumns -contains $_.Type) -and ($_.Hidden -ne $true)} | ForEach-Object{$fields.Enqueue($_)} 
        If ($fields.Count -gt 0) {
            while($fields.count -ne 0) {
                $fields.Dequeue() | ForEach-Object{
                    $field = New-Object -TypeName PSObject -Property (@{ScanId=$SMATScanID;ListTitle=($list.ListTitle -replace $specialCharacters,'');ListURL=$list.ListURL;Title=($_.Title -replace $specialCharacters,'');InternalName=$_.InternalName;Id=$_.Id;Type=$_.Type;TypeAsString=$_.TypeAsString;
                        Formula=$_.Formula;Currency=$_.Currency;CurrencyLocaleId=$_.CurrencyLocaleId;DisplayFormat=$_.DisplayFormat;OutputType=$_.OutputType;ShowAsPercentage=$_.ShowAsPercentage;DateFormat=$_.DateFormat;LookupWebId=$_.LookupWebId;
                        LookupList=$_.LookupList;LookupField=$_.LookupField;RelationshipDeleteBehavior=$_.RelationshipDeleteBehavior;IsRelationship=$_.IsRelationship;IsDependentLookup=$_.IsDependentLookup;AllowMultipleValues=$_.AllowMultipleValues;
                        UnlimitedLengthInDocumentLibrary=$_.UnlimitedLengthInDocumentLibrary;CountRelated=$_.CountRelated;PrependId=$_.PrependId;PrimaryFieldId=$_.PrimaryFieldId;JSLink=$_.JSLink;FieldTypeDefinition=$_.FieldTypeDefinition;TypeDisplayName=$_.TypeDisplayName;
                        XPath=$_.XPath;AggregationFunction=$_.AggregationFunction;PrimaryPITarget=$_.PrimaryPITarget;PrimaryPIAttribute=$_.PrimaryPIAttribute;PITarget=$_.PITarget;PIAttribute=$_.PIAttribute;TypeShortDescription=$_.TypeShortDescription;
                        FieldValueType=$_.FieldValueType;ParentList=$_.ParentList;DefaultListField=$_.DefaultListField;EntityPropertyName=$_.EntityPropertyName;NoCrawl=$_.NoCrawl;Indexable=$_.Indexable;CompositeIndexable=$_.CompositeIndexable;
                        EnforceUniqueValues=$_.EnforceUniqueValues;AuthoringInfo=$_.AuthoringInfo;TitleResource=$_.TitleResource;DescriptionResource=$_.DescriptionResource;Description=((($_.Description -replace $specialCharacters,'') -replace "`n",' ') -replace "`r",' ');RelatedField=$_.RelatedField;Group=$_.Group;
                        PreviewValueTyped=$_.PreviewValueTyped;DefaultValueTyped=$_.DefaultValueTyped;DefaultValue=$_.DefaultValue;DefaultFormula=$_.DefaultFormula;ValidationFormula=$_.ValidationFormula;ValidationMessage=$_.ValidationMessage;
                        ShowInDisplayForm=$_.ShowInDisplayForm;ShowInEditForm=$_.ShowInEditForm;ShowInListSettings=$_.ShowInListSettings;ShowInNewForm=$_.ShowInNewForm;ShowInViewForms=$_.ShowInViewForms;FromBaseType=$_.FromBaseType;Sealed=$_.Sealed;
                        CanBeDeleted=$_.CanBeDeleted;AllowDeletion=$_.AllowDeletion;CanToggleHidden=$_.CanToggleHidden;DisplaySize=$_.DisplaySize;Required=$_.Required;ReadOnlyField=$_.ReadOnlyField;Hidden=$_.Hidden;Direction=$_.Direction;JumpToField=$_.JumpToField;
                        IMEMode=$_.IMEMode;ShowInVersionHistory=$_.ShowInVersionHistory;Sortable=$_.Sortable;ListItemMenu=$_.ListItemMenu;ListItemMenuAllowed=$_.ListItemMenuAllowed;LinkToItem=$_.LinkToItem;LinkToItemAllowed=$_.LinkToItemAllowed;CalloutMenu=$_.CalloutMenu;
                        CalloutMenuAllowed=$_.CalloutMenuAllowed;Filterable=$_.Filterable;FilterableNoRecurrence=$_.FilterableNoRecurrence;Reorderable=$_.Reorderable;CanBeDisplayedInEditForm=$_.CanBeDisplayedInEditForm;PushChangesToLists=$_.PushChangesToLists;
                        Version=$_.Version;TranslationXml=$_.TranslationXml;SourceId=$_.SourceId;StaticName=$_.StaticName;UsedInWebContentTypes=$_.UsedInWebContentTypes;Scope=$_.Scope})
                    $fieldDetails += $field
                    $fieldIndex ++
                    Remove-Variable field
                }
            }
            Write-Host (" - {0} - Success" -f $fieldIndex) -ForegroundColor Green
        }
        Else {
            Write-Host " - No Complex Columns Found"
        }

        # Optimize memory usage to write to CSV in batches if PS supports it
        if ($currentPSVersion -ge $batchPSVersion) {
            if($fieldDetails.count -gt $batchExportThreshold){
                $fieldDetails | Export-Csv -Path $outputColumns -NoTypeInformation -Append
                $fieldDetails = @()
                Write-Host "      Complex Columns: Batch-point detected and wrote results to disk."
            }
        }

        # Clean up memory
        Remove-Variable fields
        #endregion #>

        #region 11b.	Last Item Modified
        Write-Host ("   {0} - Retrieving Last Modified Item: {1}" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"),$list.ListUrl) -NoNewline
        $lastModifiedItems = New-Object System.Collections.Queue;
        $currentList.GetItems($queryLastItemModified) | ForEach-Object {$lastModifiedItems.Enqueue($_)}
        If ($lastModifiedItems.Count -gt 0) {
            while($lastModifiedItems.count -ne 0) {
                $lastModifiedItems.Dequeue() | ForEach-Object{
                    $item = New-Object -TypeName PSObject -Property (@{ScanId=$SMATScanID;ListTitle=($list.ListTitle -replace $specialCharacters,'');ListUrl=$list.ListUrl;Filename=($_.Name -replace $specialCharacters,'');Title=($_.Title -replace $specialCharacters,'');URL=$_.Url;Modified=$_["Modified"];ModifiedBy=$_["Modified By"];
                        Created=$_["Created"];CreatedBy=$_["Created By"];DisplayName=$_.DisplayName;ParentList=$_.ParentList;UniqueId=$_.UniqueId;Web=$_.Web;ModerationInformation=$_.ModerationInformation;
                        File=$_.File;Folder=$_.Folder;ServerRedirected=$_.ServerRedirected;ServerRedirectedPreviewUrl=$_.ServerRedirectedPreviewUrl;ProgId=$_.ProgId;HasPublishedVersion=$_.HasPublishedVersion;
                        FileSystemObjectType=$_.FileSystemObjectType;SortType=$_.SortType;CopySource=$_.CopySource;HasUniqueRoleAssignments=$_.HasUniqueRoleAssignments;EffectiveBasePermissions=$_.EffectiveBasePermissions;
                        EffectiveBasePermissionsForUI=$_.EffectiveBasePermissionsForUI;Level=$_.Level;ContentTypeId=$_.ContentTypeId;MissingRequiredFields=$_.MissingRequiredFields})
                    $lastModifiedDetails += $item
                    $lastModifiedIndex ++
                    Remove-Variable item
                }
            }
            Write-Host (" - {0} - Success" -f $lastModifiedIndex) -ForegroundColor Green
        }
        Else {
            Write-Host " - No Item Returned" -ForegroundColor Yellow
        }

        # Optimize memory usage to write to CSV in batches if PS supports it
        if ($currentPSVersion -ge $batchPSVersion) {
            if(($lastModifiedIndex % $batchExportThreshold) -eq 0){
                $lastModifiedDetails | Export-Csv -Path $outputLastModifiedItem -NoTypeInformation -Append
                $lastModifiedDetails = @()
                Write-Host "      Last Modified: Batch-point detected and wrote results to disk."
            }
        }

        # Clean up memory
        Remove-Variable lastModifiedItems
        #endregion #>

        # Clean up memory
        Remove-Variable web
        Remove-Variable currentList
    }

    # Complex Columns: Write to CSV if PS Doesn't support batches
    if ($currentPSVersion -lt $batchPSVersion) {
        $fieldDetails | Export-Csv -Path $outputColumns -NoTypeInformation
    } Elseif ($fieldDetails.count -gt 0) {
        $fieldDetails | Export-Csv -Path $outputColumns -NoTypeInformation -Append
        Write-Host "      Complex Columns: Final batch-point detected and wrote results to disk."
    }

    # Last Modified: Write to CSV if PS Doesn't support batches
    if ($currentPSVersion -lt $batchPSVersion) {
        $lastModifiedDetails | Export-Csv -Path $outputLastModifiedItem -NoTypeInformation
    } Elseif (($lastModifiedIndex % $batchExportThreshold) -gt 0) {
        $lastModifiedDetails | Export-Csv -Path $outputLastModifiedItem -NoTypeInformation -Append
        Write-Host "      Last Modified: Final batch-point detected and wrote results to disk."
    }

    # Cleaning up memory
    Remove-Variable lists
    Remove-Variable fieldDetails
    Remove-Variable lastModifiedDetails
}
else {
	Write-Host
    Write-Host ("{0} - Large Lists Report Missing: {1}" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"), $inputLists) -ForegroundColor Yellow
}
#endregion #>

#endregion

#region Completed Script
# Output Script Info
Complete-CatScript $scriptStart[0]

# Stop logging
Stop-Transcript
#endregion
