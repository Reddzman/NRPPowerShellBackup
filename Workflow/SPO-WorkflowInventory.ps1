Clear-Host;

#$password = (Get-Credential).Password | ConvertFrom-SecureString;
#Write-Host $password;
#EXIT;

#Display script path
Write-Host $MyInvocation.MyCommand.Path

#Get script filename
$SPOScriptName = split-path $MyInvocation.MyCommand.Path -Leaf;
Write-Host $SPOScriptName;

#Get script folder
$SPOScriptFolder = split-path $MyInvocation.MyCommand.Path -Parent;
Write-Host $SPOScriptFolder;

#Get script start time
$ScriptStart = (Get-Date);
$startTime = $ScriptStart.ToString("-yyyyMMdd-hhmmss"); #(Get-Date).ToFileTimeUtc().ToString();

$global:settings = [xml]((Get-Content .\$($SPOScriptName.Replace('.ps1','.config'))) | Out-String);

$global:debug = ($global:settings.configuration.debug -eq "true");

function Get-PnPWorkflowInstance {
    param(
        [SharePointPnP.PowerShell.Commands.Base.PipeBinds.WebPipeBind]$Web, 
        [SharePointPnP.PowerShell.Commands.Base.PipeBinds.ListPipeBind]$List, 
        [SharePointPnP.PowerShell.Commands.Base.PipeBinds.ListItemPipeBind]$ListItem, 
        [SharePointPnP.PowerShell.Commands.Base.PipeBinds.WorkflowSubscriptionPipeBind]$Subscription, 
        $Offset
    )

    $getSiteInstances = ($NULL -eq $List);
    $getListInstances = ($NULL -ne $List -and $NULL -ne $ListItem);
    $useOffset = ($false -eq [String]::IsNullOrWhiteSpace($Offset) -and [Int32]::TryParse($Offset, [ref]$Offset));

    if ($getSiteInstances -or $useOffset) {
        $ctx = Get-PnPContext;

        if ($NULL -eq $Web) {
            $Web = Get-PnPWeb;
        }
    
        $workflowServicesManager = New-Object Microsoft.SharePoint.Client.WorkflowServices.WorkflowServicesManager($ctx, $Web.Web);
        $workflowInstanceService = $workflowServicesManager.GetWorkflowInstanceService();

        if ($useOffset) {
            if ($Offset -gt 0) {
                if ($getSiteInstances) {
                    $workflows = $workflowInstanceService.EnumerateInstancesForSiteWithOffset($Offset);
                }
                else{
                    if ($debug) { Write-Host "List: $($List.List.Id), Item: $($ListItem.Item.Id)"; }
                    $workflows = $workflowInstanceService.EnumerateInstancesForListItemWithOffset($List.List.Id, $ListItem.Item.Id, $Offset);
                }
                $ctx.Load($workflows);
                $ctx.ExecuteQuery();
            }
            else {
                $workflows = @();

                $instanceCount = 0;
                if ($NULL -eq $Subscription) {
                    $eventSourceId = $null;
                    if ($getSiteInstances) {
                        $eventSourceId = $Web.Web.Id;
                    }
                    else {
                        $eventSourceId = $List.List.Id;
                    }
                    (Get-PnPWorkflowSubscription | Where-Object -Property EventSourceId -eq $eventSourceId) | ForEach-Object {
                        $instanceCount += Get-WorkflowInstanceCount -Subscription $_;
                    }
                }
                else {
                    $Subscription | ForEach-Object {
                        $instanceCount += Get-WorkflowInstanceCount -Subscription $_;
                    }
                }

                Write-Host "";
                if ($getSiteInstances) {
                    Write-Host -BackgroundColor DARKGREEN -f WHITE "Total Site Workflow Instance Count: $($instanceCount) $(Get-Date))";
                }
                else {
                    Write-Host -BackgroundColor DARKGREEN -f WHITE "Total Workflow Instance Count: $($instanceCount) $(Get-Date))";
                }

                while ($workflows.Length -lt $instanceCount -or $offsetInstances.Length -gt 0) {
                    Write-Host -BackgroundColor DARKGREEN -f WHITE "`t... offset = $($workflows.Length)";
                    if ($getSiteInstances) {
                        $offsetInstances = $workflowInstanceService.EnumerateInstancesForSiteWithOffset($workflows.Length);
                    }
                    else{
                        $offsetInstances = $workflowInstanceService.EnumerateInstancesForListItemWithOffset($List.List.Id, $ListItem.Item.Id, $workflows.Length);
                    }

                    $ctx.Load($offsetInstances);
                    $ctx.ExecuteQuery();

                    $workflows += $offsetInstances;
                }             

                if ($getSiteInstances) {   
                    Write-Host -BackgroundColor DARKGREEN -f WHITE "** Site Workflow Instance Count (w/disassociations): $($workflows.Length) $(Get-Date) **";
                }
                else {
                    Write-Host -BackgroundColor DARKGREEN -f WHITE "** Workflow Instance Count (w/disassociations): $($workflows.Length) $(Get-Date) **";
                }
                Write-Host "";
            }
        }
        elseif ($getSiteInstances) {
            $workflows = $workflowInstanceService.EnumerateInstancesForSite();
            $ctx.Load($workflows);
            $ctx.ExecuteQuery();
        }

        return $workflows;
    }
    elseif ($getListInstances) {
        if ($NULL -eq $Web) {
            return SharePointPnPPowerShellOnline\Get-PnPWorkflowInstance -List $List -ListItem $ListItem;
        }
        else {
            return SharePointPnPPowerShellOnline\Get-PnPWorkflowInstance -Web $Web -List $List -ListItem $ListItem;
        }
    }
    return $();
}
function Get-WorkflowInstanceCount {
    param(
        [SharePointPnP.PowerShell.Commands.Base.PipeBinds.WebPipeBind]$Web, 
        [SharePointPnP.PowerShell.Commands.Base.PipeBinds.WorkflowSubscriptionPipeBind]$Subscription, 
        [String]$Status
    )

    if ($NULL -eq $Web) {
        $Web = Get-PnPWeb;
    }

    $ctx = Get-PnPContext;

    $workflowServicesManager = New-Object Microsoft.SharePoint.Client.WorkflowServices.WorkflowServicesManager($ctx, $Web.Web);
    $workflowInstanceService = $workflowServicesManager.GetWorkflowInstanceService();

    if ([String]::IsNullOrEmpty($Status)) {
        $count = $workflowInstanceService.CountInstances($Subscription.Subscription);
    }
    else {
        $count = $workflowInstanceService.CountInstancesWithStatus($Subscription.Subscription, $Status);
    }

    $ctx.ExecuteQuery();

    return $count.Value;
}
function Get-WorkflowTasks {
    param(
        [SharePointPnP.PowerShell.Commands.Base.PipeBinds.WorkflowInstancePipeBind]$Instance,
        [SharePointPnP.PowerShell.Commands.Base.PipeBinds.WorkflowSubscriptionPipeBind]$Subscription, 
        [SharePointPnP.PowerShell.Commands.Base.PipeBinds.WorkflowDefinitionPipeBind]$Definition
    )
<#
    if ($NULL -eq $Subscription) {
        #Get-PnPWorkflowSubscription 
    }

    if ($NULL -eq $Definition) {
        $Definition = $Subscription | ForEach-Object { Get-PnPWorkflowDefinition -Name $_.Subscription.Name }; #| Select-Object Id, DisplayName, Properties, Published;
    }
#>
    if (($NULL -eq $Instance) -or ($NULL -eq $Subscription) -or ($NULL -eq $Definition)) {
        return "";
    }

    $List = Get-PnPList -Identity ($Definition.Definition.Properties["SPDConfig.TaskListID"]);

    $query = $global:settings.configuration.output.task.InnerXml.Replace("{`$Instance.Id}", $Instance.Instance.Id);
    Write-Host "`t ... Executing $($List.Title) Query:";
    Write-Host -f YELLOW "$($query)";

    $pageSize = $global:settings.configuration.output.task.View.RowLimit;

    $wfTask = "";
    try {
        $taskItems = (Get-PnPListItem -List $List -Query $query -PageSize $pageSize);
        if ($taskItems.Length -eq 1) {
            $wfTask = $taskItems.Title;
        }
        elseif ($false -eq [Int32]::TryParse($pageSize, [ref]$pageSize) -or $taskItems.Length -lt $pageSize) {
            $taskItems | ForEach-Object {
                $wfTask += "$($_.Title)`r`n";
            }
        }
        else {
            $wfTask = $taskItems[$pageSize - 1].Title;
        }
    }
    catch {
        Write-Host "`t`t Error querying for History :-( `r`n";
        Write-Host -f RED $_.Exception.Message;
        Write-Host -f RED $_.Exception.ItemName;
    }

    return $wfTask;
}
function Get-WorkflowHistory {
    param(
        [SharePointPnP.PowerShell.Commands.Base.PipeBinds.WorkflowInstancePipeBind]$Instance,
        [SharePointPnP.PowerShell.Commands.Base.PipeBinds.WorkflowSubscriptionPipeBind]$Subscription, 
        [SharePointPnP.PowerShell.Commands.Base.PipeBinds.WorkflowDefinitionPipeBind]$Definition
    )
<#
    if ($NULL -eq $Subscription) {
        #Get-PnPWorkflowSubscription 
    }

    if ($NULL -eq $Definition) {
        $Definition = $Subscription | ForEach-Object { Get-PnPWorkflowDefinition -Name $_.Subscription.Name }; #| Select-Object Id, DisplayName, Properties, Published;
        #$Definition = ($Subscription | Get-PnPWorkflowDefinition -Name $_.Subscription.Name); #| Select-Object Id, DisplayName, Properties, Published;
    }
    else {
        
    }
#>
    if (($NULL -eq $Instance) -or ($NULL -eq $Subscription) -or ($NULL -eq $Definition)) {
        return "";
    }

    $List = Get-PnPList -Identity ($Definition.Definition.Properties["SPDConfig.HistoryListID"]);

    $query = $global:settings.configuration.output.history.InnerXml.Replace("{`$Instance.Id}", $Instance.Instance.Id);    
    Write-Host "";
    Write-Host -f CYAN "`t ... Executing $($List.Title) Query:";
    Write-Host -f YELLOW "$($query)";

    $pageSize = $global:settings.configuration.output.history.View.RowLimit;

    $wfHistory = "";
    try {
        $historyItems = (Get-PnPListItem -List $List -Query $query -PageSize $pageSize);
        if ($historyItems.Length -eq 1) {
            $wfHistory = $historyItems["Description"];
        }
        elseif ($false -eq [Int32]::TryParse($pageSize, [ref]$pageSize) -or $historyItems.Length -lt $pageSize) {
            $historyItems | ForEach-Object {
                $wfHistory += "$($_["Description"])`r`n";
            }
        }
        else {
            $wfHistory = $historyItems[$pageSize - 1]["Description"];
        }
        Write-Host "`t`t $($wfHistory)"
        Write-Host "";
    }
    catch {
        Write-Host "`t`t Error querying for History :-( `r`n";
        Write-Host -f RED $_.Exception.Message;
        Write-Host -f RED $_.Exception.ItemName;
    }
    return $wfHistory;
}
function Get-WorkflowDetails {
    param(
        [SharePointPnP.PowerShell.Commands.Base.PipeBinds.WebPipeBind]$Web,
        [SharePointPnP.PowerShell.Commands.Base.PipeBinds.WorkflowInstancePipeBind]$Instance,
        [SharePointPnP.PowerShell.Commands.Base.PipeBinds.WorkflowSubscriptionPipeBind]$Subscription,
        [SharePointPnP.PowerShell.Commands.Base.PipeBinds.WorkflowDefinitionPipeBind]$Definition
    )

    if ($NULL -eq $Web) {
        $Web = Get-PnPWeb;
    }

    $wf = New-Object PSObject;

    $wf | Add-Member -MemberType NoteProperty -name "Site" -value $Web.Web.Title;
    $wf | Add-Member -MemberType NoteProperty -name "SiteUrl" -value $Web.Web.URL;

    if ($NULL -ne $Subscription) {
        $wf | Add-Member -MemberType NoteProperty -name "WFName" -value $Subscription.Subscription.Name;
    }
    else {
        $wf | Add-Member -MemberType NoteProperty -name "WFName" -value "";
    }

    if ($NULL -ne $Instance) {
        $wf | Add-Member -MemberType NoteProperty -name "WFStatusUrl" -value "$($Web.Web.URL)/_layouts/15/wrkstat.aspx?WorkflowInstanceName=$($Instance.Instance.Id)";
        #$wf | Add-Member -MemberType NoteProperty -name "WFInstanceID" -value $Instance.Instance.Id;
        $wf | Add-Member -MemberType NoteProperty -name "WFInternalStatus" -value $Instance.Instance.Status;
        $wf | Add-Member -MemberType NoteProperty -name "WFUserStatus" -value $Instance.Instance.UserStatus;
        $wf | Add-Member -MemberType NoteProperty -name "WFStarted" -value $Instance.Instance.InstanceCreated;
        $wf | Add-Member -MemberType NoteProperty -name "WFLastUpdated" -value $Instance.Instance.LastUpdated;
        $wf | Add-Member -MemberType NoteProperty -name "WFError" -value $Instance.Instance.FaultInfo;
    }
    else {
        $wf | Add-Member -MemberType NoteProperty -name "WFStatusUrl" -value "";
        #$wf | Add-Member -MemberType NoteProperty -name "WFInstanceID" -value "";
        $wf | Add-Member -MemberType NoteProperty -name "WFInternalStatus" -value "";
        $wf | Add-Member -MemberType NoteProperty -name "WFUserStatus" -value "";
        $wf | Add-Member -MemberType NoteProperty -name "WFStarted" -value "";
        $wf | Add-Member -MemberType NoteProperty -name "WFLastUpdated" -value "";
        $wf | Add-Member -MemberType NoteProperty -name "WFError" -value "";
    }

    $wf | Add-Member -MemberType NoteProperty -name "WFLogicCheckWarning" -value (Get-WorkflowDetailsForStatus -Instance $Instance);
    $wf | Add-Member -MemberType NoteProperty -name "WFRecentHistory" -value (Get-WorkflowHistory -Instance $Instance -Subscription $Subscription -Definition $Definition);

    return $wf;
}
function Get-SiteWorkflowDetails {
    param(
        [SharePointPnP.PowerShell.Commands.Base.PipeBinds.WebPipeBind]$Web,
        [SharePointPnP.PowerShell.Commands.Base.PipeBinds.WorkflowInstancePipeBind]$Instance,
        [SharePointPnP.PowerShell.Commands.Base.PipeBinds.WorkflowSubscriptionPipeBind]$Subscription    
    )

    if ($NULL -eq $Web) {
        return (Get-WorkflowDetails -Instance $Instance -Subscription $Subscription);
    }
    else {
        return (Get-WorkflowDetails -Web $Web -Instance $Instance -Subscription $Subscription);
    }
}
function Get-ListWorkflowDetails {
    param(
        [SharePointPnP.PowerShell.Commands.Base.PipeBinds.WebPipeBind]$Web,
        [SharePointPnP.PowerShell.Commands.Base.PipeBinds.ListPipeBind]$List,
        [SharePointPnP.PowerShell.Commands.Base.PipeBinds.ListItemPipeBind]$ListItem,
        [SharePointPnP.PowerShell.Commands.Base.PipeBinds.WorkflowInstancePipeBind]$Instance,
        [SharePointPnP.PowerShell.Commands.Base.PipeBinds.WorkflowSubscriptionPipeBind]$Subscription,
        [SharePointPnP.PowerShell.Commands.Base.PipeBinds.WorkflowDefinitionPipeBind]$Definition
    )

    if ($NULL -eq $Web) {
        $wf = (Get-WorkflowDetails -List $List -Instance $Instance -Subscription $Subscription -Definition $Definition);
        $siteURI = [Uri](Get-PnPWeb).Url;
    }
    else {
        $wf = (Get-WorkflowDetails -Web $Web -List $List -Instance $Instance -Subscription $Subscription -Definition $Definition);
        $siteURI = [Uri]$Web.Web.Url;
    }

    $wf.WFStatusUrl += "&List=$($List.List.Id)";

    $wf | Add-Member -MemberType NoteProperty -name "List" -value $List.List.Title;
    $wf | Add-Member -MemberType NoteProperty -name "ItemGuid" -value $ListItem.Item["UniqueId"];
    $wf | Add-Member -MemberType NoteProperty -name "ItemID" -value $ListItem.Item.Id;
    $wf | Add-Member -MemberType NoteProperty -name "ItemUrl" -value "$($siteURI.scheme)://$($siteURI.host)$($List.List.RootFolder.ServerRelativeUrl)/DispForm.aspx?ID=$($ListItem.Item.Id)"
    $wf | Add-Member -MemberType NoteProperty -name "ItemTitle" -value $ListItem.Item.Title;
    $wf | Add-Member -MemberType NoteProperty -name "ItemModified" -value $ListItem.Item["Modified"];
    
    $liMsg = "";
    switch ($List.List.RootFolder.Name) {
        "ComplianceEventQueue" {
            $wf | Add-Member -MemberType NoteProperty -name "ItemLogicCheckWarning" -value "";
            $wf | Add-Member -MemberType NoteProperty -name "ItemLogicQueued" -value "";
        }
        "ComplianceWFReportSchedules" {
            $wf | Add-Member -MemberType NoteProperty -name "ItemLogicCheckWarning" -value "";
            $wf | Add-Member -MemberType NoteProperty -name "ItemLogicQueued" -value "";
        }
        default {
            $ComplianceDueDate = $ListItem.Item["HPComplianceDueDate"];
            $ComplianceStatus = $ListItem.Item["HPComplianceStatus"];
            $Recurrence = $ListItem.Item["fRecurrence"];
            $MilestoneStatus = $ListItem.Item["HPComplianceMilestones"] + "-" + $ListItem.Item["HPComplianceMilestoneStatus"];    

            if ($NULL -eq $ComplianceDueDate) {
                ## DO NOTHING ... ##
            }
            elseif (([datetime]$ComplianceDueDate).Date.AddDays(15).Date -lt [datetime]::Now.Date -and $ComplianceStatus -ne "Exception") {
                if ($ComplianceStatus -ne "Completed" -or $Recurrence -eq $true) {
                    $liMsg = "Compliance Due Date (recurring = $($Recurrence)) exception threshold has past, Status should be: Exception.";
                }
            }
            elseif (([datetime]$ComplianceDueDate).Date -lt [datetime]::Now.Date -and $ComplianceStatus -ne "Completed") {
                $liMsg = "Compliance Due Date has already past, and Status should be: Completed or Exception.";
            }
            elseif ($MilestoneStatus -ne "Event-End-Completed" -and $MilestoneStatus.EndsWith("-Completed")) {
                $liMsg = "Milestone has been completed, but not advanced to next milestone. Check the Compliance Queue.";
            }

            $wf | Add-Member -MemberType NoteProperty -name "ItemLogicCheckWarning" -value $liMsg;
            $wf | Add-Member -MemberType NoteProperty -name "ItemLogicQueued" -value "$($global:queuedWorkflows -contains $ListItem.Item["UniqueId"])";

            if ($ebug) { Write-Host "`t ... $($List.List.Title) #$($ListItem.Item.Id) [$($ListItem.Item["UniqueId"])] is queued: $($global:queuedWorkflows -contains $ListItem.Item["UniqueId"])" }
        }
    }

	$outputViewFields = $global:settings.configuration.output.viewfields;
	if ($null -ne $outputViewFields -and $null -ne $outputViewFields.fieldref) {
		$outputViewfields.fieldref.name | ForEach-Object {
			$wf | Add-Member -MemberType NoteProperty -name $_ -value $ListItem.Item[$_];
		}
	}
    return $wf;
}
function Get-WorkflowDetailsForStatus {
    param(
        [SharePointPnP.PowerShell.Commands.Base.PipeBinds.WorkflowInstancePipeBind]$Instance
    )

    $wfMsg = "";

    switch ($Instance.Instance.Status) {
        "Started" {
            $statusDate = $null;
        
            $dateTimeMatch = [regex]::Matches($Instance.Instance.UserStatus, "(\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}Z)");
            if ($dateTimeMatch.count -eq 0) {
                $dateTimeMatch = [regex]::Matches($Instance.Instance.UserStatus, "(\d{1,2}\/\d{1,2}\/\d{4} \d{2}:\d{2}:\d{2} [A|P]M)");
            }
    
            if ($dateTimeMatch.count -gt 0) {
                $statusDate = [datetime]::Parse($dateTimeMatch[0].Groups[1].Value).Date;
                if ($Instance.Instance.UserStatus -like "*Wait*" -or $Instance.Instance.UserStatus -like "*Paus*") {
                    if ($statusDate.Date -lt [datetime]::Now.Date) {
                        $wfMsg = "Workflow stalled: Date for wait/pause activity has already past; ";
                    }
                    else {
                        $wfMsg = "Workflow waiting/paused; ";
                    }
                }
            }
    
            if ($Instance.Instance.Status -eq "Started" -and $Instance.Instance.LastUpdated -lt [datetime]::Now.AddHours(-1)) {
                $wfMsg += "Workflow running: Last update over 1 hr. ago, MIGHT be stalled if not tasked/waiting/paused; "
            }    
        }

        "Suspended" {
            $wfMsg = "Workflow Error: Suspended, restart necessary!";
        }
    }
    return $wfMsg;
}

$logFile = $SPOScriptName.Replace(".ps1", $startTime);
$csvFile = $SPOScriptName.Replace(".ps1", $startTime);

#try { Stop-Transcript } catch { }

try {
    Start-Transcript -path .\$($logFile).log -append
    Write-Output "";

    if (Test-Path .\$($csvFile)-Workflows.csv) {
        Write-Output "Deleting file: '$($csvFile)-Workflows.csv' to prevent duplicate CSV entries from -append";
        Remove-Item .\$($csvFile)-Workflows.csv
        Write-Output "";
    }

    #Credentials
	if ([String]::IsNullOrWhiteSpace($global:settings.configuration.credentials.p)) {
		$creds = Get-Credential -credential $global:settings.configuration.credentials.u;
	}
	else {		
		$creds = New-Object System.Management.Automation.PSCredential(
			$global:settings.configuration.credentials.u, ($global:settings.configuration.credentials.p | ConvertTo-SecureString)
		);
	}

    $global:queuedWorkflows = @();

	$siteUrl = $global:settings.configuration.siteUrl;
	
	$csvSource = $false;
	$importCsvFileName = "ModernizationWorkflowScanResults";

	#Get-ChildItem -Filter *.csv | Select-Object -ExpandProperty FullName | Import-Csv | Export-Csv .\combinedcsvs.csv -NoTypeInformation -Append

	try {
		$importExport = Import-Csv .\$($importCsvFileName).csv | Select-Object @{ Expression={$_.'Site Url'.Replace($siteUrl, '')}; Name='serverRelativeUrl' }, @{ Expression={$_.'List Url'.Replace($_.'Site Url'.Replace($siteUrl, ''), '')}; Name='webRelativeUrl' } | Sort 'serverRelativeUrl', 'webRelativeUrl' -Unique;
		#ConvertTo-Xml -InputObject $importExport -As 'string' | Out-File -Width 80 .\$($importCsvFileName)-Sites-n-Lists.xml;	

		$importWebs = Import-Csv .\$($importCsvFileName).csv | Select-Object @{ Expression={$_.'Site Url'.Replace($siteUrl, '')}; Name='serverRelativeUrl' } | Sort 'serverRelativeUrl' -Unique;
		Write-Host -F MAGENTA "Found CSV of webs/lists!  Ignoring script's config XML file for webs/lists.";

		$csvSource = $true;
	}
	catch {	
		Write-Host -F MAGENTA "No CSV input of webs/lists.  Using script's config XML file.";
		$importWebs = $global:settings.configuration.web;	
	}
	
	<#	
	Write-Host -F YELLOW "Iterating $($importWebs.Count) Webs ...";
	$importWebs | ForEach-Object { 
		$xmlWeb = $_;

		if ($csvSource) {
			$importLists = $importExport | Where-Object { $_.serverRelativeUrl -eq $xmlWeb.serverRelativeUrl -and $_.webRelativeUrl -ne "" };
		}
		else {
			Write-Host -F MAGENTA "No Lists from CSV $($xmlWeb.serverRelativeUrl)/"	
			$importLists = $xmlWeb.List;
		}

		Write-Host;
		Write-Host -F YELLOW "... Interating $($importLists.Count) lists in $($xmlWeb.serverRelativeUrl)/";
		$importLists | ForEach-Object {
			$xmlList = $_;
			Write-Host -F GREEN "`t $($xmlList.webRelativeUrl)";
		}
	}
	EXIT;
	#>

	#Connect to Webs + Lists
	Write-Host -F YELLOW "Iterating $($importWebs.Count) Webs ...";
    $importWebs | ForEach-Object {
        $xmlWeb = $_;

        $webUrl = [Uri]$NULL;

        if ([Uri]::TryCreate([Uri]$global:settings.configuration.siteUrl, $xmlWeb.serverRelativeUrl, [ref]$webUrl)) {
            Write-Host -f GREEN "Connecting to site '$($webUrl)' ..."

            Connect-PnPOnline -Url $webUrl -Credentials $creds;
            
            $WFDefinitions = @{ };
            $WFSubscriptions = @{ };

            Write-Host "... Get workflow subscriptions";
            Get-PnPWorkflowSubscription | ForEach-Object {
                $subscription = $_;

                $WFSubscriptions[$subscription.Id] = $subscription;
                $WFDefinitions[$subscription.Id] = (Get-PnPWorkflowDefinition -Name $subscription.Name); #| Select-Object Id, DisplayName, Properties, Published);

                Write-Host -f YELLOW "`t $(Get-WorkflowInstanceCount -Subscription $subscription) $($subscription.Name) instance(s)";

                @("Started","Suspended","NotStarted","Canceling","Canceled","Terminated","Completed","NotSpecified","Invalid") | ForEach-Object {
                    $highlight = "WHITE";

                    $count = (Get-WorkflowInstanceCount -Subscription $subscription -Status $_);
                    if ($count -gt 0) {
                        switch ($_) {
                            "Suspended" { $highlight = "RED" }
                            "Started" { $highlight = "GREEN" }
                            #"Terminated" { $highlight = "DARKRED" }
                        }
                    }
                    Write-Host -f $highlight "`t`t $($count) [$($_)]";
                }
            }
	
			if ($csvSource) {
				$importLists = $importExport | Where-Object { $_.serverRelativeUrl -eq $xmlWeb.serverRelativeUrl -and $_.webRelativeUrl -ne "" };
			}
			else {
				$importLists = $xmlWeb.List;
			}
			
			#if ($null -eq $importLists) {
			#	$importLists = Get-PnPList;
			#}
			
			Write-Host -F YELLOW "... Interating $($importLists.Count) lists in $($xmlWeb.serverRelativeUrl)/";
            $importLists | ForEach-Object {
                $xmlList = $_;

                Write-Host "... Getting list " -NoNewline;
                Get-PnPList -Identity $xmlList.webRelativeUrl -Includes RootFolder | ForEach-Object {
                    $list = $_;
                    Write-Host -f YELLOW "$($list.Title)";

                    if ([String]::IsNullOrWhiteSpace($xmlList.View)) {
                        $listItems = Get-PnPListItem -List $list -Fields ($global:settings.configuration.output.viewfields.fieldref.name);
                    }
                    else {
                        $listItems = Get-PnPListItem -List $list -query $xmlList.View.OuterXml;
                    }
                    
                    $listWorkflows = @();

                    $listItems | ForEach-Object {
                        $item = $_;
                        Write-Host -f GREEN "`t Get workflow instance(s) for $($list.Title) #$($item.Id)";

                        switch ($list.RootFolder.Name) {
                            "ComplianceEventQueue" {
                                $global:queuedWorkflows += $item["HPComplianceEventGUID"];
                                    #if ($debug) { Write-Host "`t ... Field - GUID: $($item["GUID"]), `r`n`t ... Field - UniqueId: $($item["UniqueId"])" }
                                    #if ($debug) { Write-Host "`t ... Property - GUID: $($item.GUID), `r`n`t ... Property - UniqueId: $($item.UniqueId)" }
                                    if ($debug) { Write-Host "`t ... $($item["HPComplianceEventsListName"]) #$($item["HPComplianceEventID"]) [$($item["HPComplianceEventGUID"])] is #$($global:queuedWorkflows.Length) in the queue" }
                            }
                        }

                        $workflowInstances = (Get-PnPWorkflowInstance -List $list -ListItem $item); 

                        if ($workflowInstances.Length -eq 0) {
                            $listWorkflows += (Get-ListWorkflowDetails -List $list -ListItem $item);
                        }
                        else {
                            $workflowInstances | Where-Object { $global:settings.configuration.workflow.instance.status -contains $_.Status } | ForEach-Object { 
                                $listWorkflows += (Get-ListWorkflowDetails -List $list -ListItem $item -Instance $_ -Subscription $WFSubscriptions[$_.WorkflowSubscriptionId] -Definition $WFDefinitions[$_.WorkflowSubscriptionId]);
                            }
                        }
                    }

                    if ($listWorkflows.Length -gt 0) {
                        $listWorkflows | Export-Csv .\$($csvFile)-Workflows.csv -notype -append -force;
                        Write-Host "";    
                    }
                    
                    switch ($list.RootFolder.Name) {
                        "ComplianceEventQueue" {
                            Write-Host "";
                            Write-Host -F MAGENTA -BackgroundColor WHITE "** $($global:queuedWorkflows.Length) Compliance Events have been queued **";
                            Write-Host "";
                        }
                    }
                }
            }

			$instanceStatii = $global:settings.configuration.workflow.instance.status;

			#$MaxThreads = 5
			#$RunspacePool = [runspacefactory]::CreateRunspacePool(1, $MaxThreads)
			#$PowerShell = [powershell]::Create()
			#$PowerShell.RunspacePool = $RunspacePool
			#$RunspacePool.Open()
			#$PowerShell.AddScript({
			#	param (
			#		$instanceStatii,
			#		$WFSubscriptions,
			#		$WFDefinitions,
			#		$csvFile,
			#	)
				Write-Host "... Get site workflow instances";
				$siteInstances = (Get-PnPWorkflowInstance -Offset 0 | Where-Object { $instanceStatii -contains $_.Status });
				#($siteInstances | Format-List);

				Write-Host "`t $($siteInstances.Length) site workflow(s) with status of '$($instanceStatus -Join ''' or ''')' ...";

				$siteWorkflows = @();
				$siteInstances | ForEach-Object {
					$siteWorkflows += (Get-SiteWorkflowDetails -Instance $_ -Subscription $WFSubscriptions[$_.WorkflowSubscriptionId] -Definition $WFDefinitions[$_.WorkflowSubscriptionId]);
				}

				if ($siteWorkflows.Length -gt 0) {
					$siteWorkflows | Export-Csv .\$($csvFile)-SiteWorkflows.csv -notype -append -force;
					Write-Host "";    
				}
			#});
			#
			#$PowerShell.AddArgument($instanceStatus);
			#$PowerShell.AddArgument($WFSubscriptions);
			#$PowerShell.AddArgument($WFDefinitions);
			#$PowerShell.AddArgument($csvFile);
			#
			#$PowerShell.BeginInvoke();
        }
    }
	
    Write-Host "Uploading file .\$($csvFile)-Workflow.csv to $($global:settings.configuration.siteUrl) ...";
    Connect-PnPOnline -Url $global:settings.configuration.siteUrl -Credentials $creds;

    $SPFilename = $global:settings.configuration.output.fileName;
    if ([String]::IsNullOrEmpty($SPFilename)) {
        Add-PnPFile -Path .\$($csvFile)-Workflows.csv -Folder $global:settings.configuration.output.webRelativeUrl;
    }
    else {
        $SPFilename += (Get-Date).ToString("-yyyyMMdd");
        
        Copy-Item -Path .\$($csvFile)-Workflows.csv -Destination .\$($SPFilename).csv
        Add-PnPFile -Path .\$($SPFilename).csv -Folder $global:settings.configuration.output.webRelativeUrl;
        Remove-Item -Path .\$($SPFilename).csv;
    }

    Write-Host "... done!";

    $ScriptEnd = (Get-Date);#.ToString("-yyyyMMdd-hhmmss");

    Write-Host -f GREEN "Started: $($ScriptStart) `r`n  Ended: $($ScriptEnd)";
    Write-Output "";
}
finally {
    Stop-Transcript
}
