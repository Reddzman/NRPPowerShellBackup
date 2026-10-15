<#  Get-CatPostMigItemInfo (PnP)
    This function reports migration related data for a items in the web. (Long Running)

    Get-CatPostMigItemInfo -ScanURL $VarScanURL -ScanCreds $VarScanCred -ReportURL $VarReportUrl -ReportCred $VarReportCred -WCount $VarWebCount -WTotal $VarWebTTotal

    Parameters
        Env - Name of the Environment being scanned
        ScanURL - URL of Site to be scanned
        ScanCreds - Credentials for access to site to be scanned
		ReportURL - URL of Reporting Site
        ReportCreds - Credentials for access to Reporting site
        WCount - Current Web index
        WTotal - Total Number of Webs
#>
function global:Get-CatPostMigItemInfo
{
    param (
        [parameter(Mandatory=$true, Position=0, ValueFromPipeline=$true)]
        [String]$Env,
        
        [parameter(Mandatory=$true, Position=0, ValueFromPipeline=$true)]
        [String]$ScanUrl,
        
        [parameter(Mandatory=$true, Position=1, ValueFromPipeline=$true)]
        [System.Management.Automation.PSCredential]$ScanCreds,

		[parameter(Mandatory=$true, Position=0, ValueFromPipeline=$true)]
        [String]$ReportUrl,
        
        [parameter(Mandatory=$true, Position=1, ValueFromPipeline=$true)]
        [System.Management.Automation.PSCredential]$ReportCreds,
        
        [parameter(Mandatory=$true, Position=2, ValueFromPipeline=$true)]
        [Int]$WCount,
        
        [parameter(Mandatory=$true, Position=2, ValueFromPipeline=$true)]
        [Int]$WTotal)
    
    # Connect to SharePoint 2013 Web
    Connect-PnPOnline -Url $ScanUrl -Credentials $ScanCreds

    # Get web object with additional properties
    $CurWeb = Get-PnPWeb -Includes Lists, HasUniqueRoleAssignments, RoleAssignments
    $CurSite = Get-PnPSite -Includes RootWeb
    Write-Host ("      Scanning Web: {0} - {1} out of {2}" -f $CurWeb.Title, $WCount, $WTotal) -ForegroundColor DarkGray
                
    # Loop through all lists and libraries in a sub site
    $ListCount = 0
    ForEach ($List in $CurWeb.Lists)
    {
        # Confirm Not a hidden list
        If (-not $List.Hidden)
        {
            # Get list object with additional properties
            $ListCount++
            $CurList = Get-PnPList -Identity $List.Id -Includes ItemCount, HasUniqueRoleAssignments, RoleAssignments, RootFolder.ServerRelativeUrl
            Write-Host ("         Scanning List: {0} - {1} out of {2} (Minus Hidden Lists)" -f $CurList.Title, $ListCount, $CurWeb.Lists.Count)

            # Get all items in current list
            $Items = Get-PnPListItem -List $CurList -PageSize 2000

            # Loop through all items in list
            $ItemCount = 0
            ForEach ($Item in $Items)
            {
                # Get Permissions for this item
                $ItemCount++
                $CurItem = Get-PnPListItem -List $List.Id -Id $Item.Id
                Get-PnPProperty -ClientObject $CurItem -Property HasUniqueRoleAssignments, RoleAssignments

                # Determine if this is a list or a library
                If ($List.BaseType -eq "DocumentLibrary")
                {
                    Write-Host ("            Adding Version Record: {0} - {1} out of {2}" -f $Item["FileLeafRef"], $ItemCount, $CurList.ItemCount)
                    $ItemValues = @{"MigEnvName"=$Env;
                                "MigSCName"=$CurSite.RootWeb.Title;"MigSCURL"=$CurSite.Url;
                                "MigWebName"=$CurWeb.Title;"MigWebURL"=$CurWeb.Url;
                                "MigListName"=$CurList.Title;"MigListURL"=$CurList.RootFolder.ServerRelativeUrl;
                                "MigItemTitle"=$CurItem["FileLeafRef"];"MigItemURL"=$CurItem["FileRef"];
                                "MigItemVersion"=$CurItem["_UIVersionString"];}
                    $Temp = Add-CatMigrationRecord -Url $ReportURL -Creds $ReportCreds -L "Migration Item Version Counts" -ContentType "Migration Version Item" -Values $ItemValues
                    Connect-PnPOnline -Url $ScanUrl -Credentials $ScanCreds

                    if($CurItem.HasUniqueRoleAssignments)
                    {
                        Write-Host ("               Adding Permissions Record: {0} - {1} out of {2}" -f $Item["FileLeafRef"], $ItemCount, $CurList.ItemCount)
                        $ItemValues = @{"MigEnvName"=$Env;
                                    "MigSCName"=$CurSite.RootWeb.Title;"MigSCURL"=$CurSite.Url;
                                    "MigWebName"=$CurWeb.Title;"MigWebURL"=$CurWeb.Url;
                                    "MigListName"=$CurList.Title;"MigListURL"=$CurList.RootFolder.ServerRelativeUrl;
                                    "MigItemTitle"=$CurItem["FileLeafRef"];"MigItemURL"=$CurItem["FileRef"];
                                    "MigRoleCount"=$CurItem.RoleAssignments.Count;}
                        $Temp = Add-CatMigrationRecord -Url $ReportURL -Creds $ReportCreds -L "Migration Permission Counts" -ContentType "Migration Permission Item" -Values $ItemValues
                        Connect-PnPOnline -Url $ScanUrl -Credentials $ScanCreds
                    } # if($CurItem.HasUniqueRoleAssignments)

                } # If ($List.BaseType -eq "DocumentLibrary")
                Else
                {
                    Write-Host ("            Adding Version Record: {0} - {1} out of {2}" -f $Item["Title"], $ItemCount, $CurList.Items.Count)
                    $ItemValues = @{"MigEnvName"=$Env;
                                "MigSCName"=$CurSite.RootWeb.Title;"MigSCURL"=$CurSite.Url;
                                "MigWebName"=$CurWeb.Title;"MigWebURL"=$CurWeb.Url;
                                "MigListName"=$CurList.Title;"MigListURL"=$CurList.RootFolder.ServerRelativeUrl;
                                "MigItemTitle"=$CurItem["Title"];"MigItemURL"=$CurItem["FileRef"];
                                "MigItemVersion"=$CurItem["_UIVersionString"];}
                    $Temp = Add-CatMigrationRecord -Url $ReportURL -Creds $ReportCreds -L "Migration Item Version Counts" -ContentType "Migration Version Item" -Values $ItemValues
                    Connect-PnPOnline -Url $ScanUrl -Credentials $ScanCreds

                    if($CurItem.HasUniqueRoleAssignments)
                    {
                        Write-Host ("               Adding Permission Record: {0} - {1} out of {2}" -f $Item["Title"], $ItemCount, $CurList.Items.Count)
                        $ItemValues = @{"MigEnvName"=$Env;
                                    "MigSCName"=$CurSite.RootWeb.Title;"MigSCURL"=$CurSite.Url;
                                    "MigWebName"=$CurWeb.Title;"MigWebURL"=$CurWeb.Url;
                                    "MigListName"=$CurList.Title;"MigListURL"=$CurList.RootFolder.ServerRelativeUrl;
                                    "MigItemTitle"=$CurItem["Title"];"MigItemURL"=$CurItem["FileRef"];
                                    "MigRoleCount"=$CurItem.RoleAssignments.Count;}
                        $Temp = Add-CatMigrationRecord -Url $ReportURL -Creds $ReportCreds -L "Migration Permission Counts" -ContentType "Migration Permission Item" -Values $ItemValues
                        Connect-PnPOnline -Url $ScanUrl -Credentials $ScanCreds
                    } # if($CurItem.HasUniqueRoleAssignments -eq $true)
                }
            } # ForEach ($Item in $Items)		
        } # If (-not $List.Hidden)
    } # ForEach ($List in $CurWeb.Lists)

    # Screen Output for Web
    Write-Host ("         Item Scan Complete: {0}" -f $CurWeb.Title) -ForegroundColor Green
}