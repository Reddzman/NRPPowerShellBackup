<# Cat.PnP-GetAllWebs.psm1

This function uses SharePoint PnP PowerShell to get all web objects with a site collection.

Sample Call: 
    Get-CatPnPAllWebs -Creds $spCred -URL $spURL

#>
function Get-CatPnPAllWebs{
	param(
        [System.Management.Automation.PSCredential]$Creds,
		[string]$URL
    )

    # Create PnP Context
    $tempConnection = Connect-PnPOnline -Credentials $Creds -Url $URL

    # Create array for web objects
    $Ws = @()

    # Add root web to array
    $Ws += Get-PnPWeb

    # Add sub webs to array
    $Ws += Get-PnPSubWebs -Recurse

    # Output web array
    $Ws
}