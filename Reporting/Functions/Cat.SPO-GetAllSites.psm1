<#  Cat.SPO-GetAllSites.psm1
This function uses SPO PowerShell Modules to get all site collections in a tenant. It assumes there is an 
active PowerShell connection to the SharePoint Online Admin Center.

Command: Get-CatSPOAllSites -personal $true

Parameters
    personal - Should it include personal Sites: $true or $false
#>

function global:Get-CatSPOAllSites
{ 
	param (
            [bool] $personal
    )
    # Get site collection objects
    If ($personal) {
        $sites = Get-SPOSite -Limit All -IncludePersonalSite $personal
    }
    Else {
        $sites = Get-SPOSite -Limit All -IncludePersonalSite $personal | Where{$_.URL -notlike "*-my.sharepoint.com*"}
    }

    $sites
}