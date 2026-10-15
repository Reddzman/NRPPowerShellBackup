<#  Get-CatPostMigWebListInfo (PnP)
    This function reports migration related data for a web and all lists in the web.

    Get-CatPostMigWebListInfo -ScanURL $VarScanURL -ScanCreds $VarScanCred -ReportURL $VarReportUrl -ReportCred $VarReportCred -WCount $VarWebCount -WTotal $VarWebTTotal

    Parameters
        Env - Name of the Environment being scanned
        ScanURL - URL of Site to be scanned
        ScanCreds - Credentials for access to site to be scanned
		ReportURL - URL of Reporting Site
        ReportCreds - Credentials for access to Reporting site
        WCount - Current Web index
        WTotal - Total Number of Webs
#>
function global:Get-CatPostMigWebListInfo
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
    $CurSite = Get-PnPSite -Includes RootWeb
    $CurWeb = Get-PnPWeb -Includes Lists, HasUniqueRoleAssignments, RoleAssignments
    Write-Host ("      Scanning Web: {0} - {1} out of {2}" -f $CurWeb.Title, $WCount, $WTotal) -ForegroundColor DarkGray

    # Output permissions for current web to CSV file if inheritance is broken
    if($CurWeb.HasUniqueRoleAssignments)
    {
        Write-Host ("         Adding Permission Record: {0}" -f $CurWeb.Title)
        $ItemValues = @{"MigEnvName"=$Env;
                    "MigSCName"=$CurSite.RootWeb.Title;"MigSCURL"=$CurSite.Url;
                    "MigWebName"=$CurWeb.Title;"MigWebURL"=$CurWeb.Url;
					"MigRoleCount"=$CurWeb.RoleAssignments.Count;}
        $Temp = Add-CatMigrationRecord -Url $ReportURL -Creds $ReportCreds -L "Migration Permission Counts" -ContentType "Migration Permission Item" -Values $ItemValues
        Connect-PnPOnline -Url $ScanUrl -Credentials $ScanCreds
    } # if($CurWeb.HasUniqueRoleAssignments)\
    
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
            Write-Host ("         Adding Item Count Record: {0} - {1} out of {2} (Minus Hidden Lists)" -f $CurList.Title, $ListCount, $CurWeb.Lists.Count)
            $ListUrl = ("https://{0}{1}" -f $Host.Host, $CurList.RootFolder.ServerRelativeUrl)

            # Adding Item Count Information to Reporting Site
            $ItemValues = @{"MigEnvName"=$Env;
                        "MigSCName"=$CurSite.RootWeb.Title;"MigSCURL"=$CurSite.Url;
                        "MigWebName"=$CurWeb.Title;"MigWebURL"=$CurWeb.Url;
                        "MigListName"=$CurList.Title;"MigListURL"=$CurList.RootFolder.ServerRelativeUrl;
                        "MigListItemCount"=$CurList.ItemCount;}
            $Temp = Add-CatMigrationRecord -Url $ReportURL -Creds $ReportCreds -L "Migration List Item Counts" -ContentType "Migration Item Count Item" -Values $ItemValues
            Connect-PnPOnline -Url $ScanUrl -Credentials $ScanCreds

            # Output permissions for the curent list to CSV file if inheritance is broken
            if($CurList.HasUniqueRoleAssignments)
            {
                Write-Host ("            Adding Permission Record: {0}" -f $CurList.Title)
                $ItemValues = @{"MigEnvName"=$Env;
                    "MigSCName"=$CurSite.RootWeb.Title;"MigSCURL"=$CurSite.Url;
                    "MigWebName"=$CurWeb.Title;"MigWebURL"=$CurWeb.Url;
                    "MigListName"=$CurList.Title;"MigListURL"=$ListUrl;
					"MigRoleCount"=$CurList.RoleAssignments.Count;}
                $Temp = Add-CatMigrationRecord -Url $ReportURL -Creds $ReportCreds -L "Migration Permission Counts" -ContentType "Migration Permission Item" -Values $ItemValues                
                Connect-PnPOnline -Url $ScanUrl -Credentials $ScanCreds
            } # if($CurList.HasUniqueRoleAssignments)			
        } # If (-not $List.Hidden)
    } # ForEach ($List in $CurWeb.Lists)

    # Screen Output for Web
    Write-Host ("         Web & List Scan Complete: {0}" -f $CurWeb.Title) -ForegroundColor Green
}