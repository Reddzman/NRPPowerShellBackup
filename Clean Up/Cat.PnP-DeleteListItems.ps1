<# Cat.PnP-PurgeListItems.ps1
This script deletes all items from the lists in the "$ListNames" array.  Need to validate if items go to the
recycle bin or not.
	
Dependencies
    -	PnP PowerShell https://msdn.microsoft.com/en-us/pnp_powershell/pnp-powershell-overview
    
Table of Contents
	1.	Variables
	2.	Main Script
#>

# region 1. Variables
$ListNames = "Upgrade - InfoPath", 
             "Upgrade - Large Lists", 
             "Upgrade - Complex Columns"


$SPOCred = Get-Credential -Message "Enter your credentials for SharePoint Online:" 
$SPOUrl = "https://terumobct.sharepoint.com/sites/spupgrade"
Write-Host
Connect-PnPOnline -Url $SPOUrl -Credentials $SPOCred
$Site = Get-PnPSite -Includes RootWeb
$SPOCtx = Get-PnPContext
$Web = Get-PnPWeb 
#endregion

#region 2.	Main Script
foreach($ListName in $ListNames)
{
    write-host "Processing $ListName"
    
    # Get the List and the Items in the list
    $List = Get-PnPList -Identity $ListName 
    $Items = Get-PnPListItem -List $ListName -Fields Title, "ListName", "URL"
    foreach($Item in $Items)
    {
        Remove-PnPListItem -List $ListName -Identity $Item.Id -Force
    }
} #foreach($ListName in $ListNames)
#endregion
write-host "Done" -ForegroundColor Green