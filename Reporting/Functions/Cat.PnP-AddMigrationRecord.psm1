<#  Add-CatMigrationRecord (PnP)
    This function create a site record in a reporting site for migration discovery.  The site, Lists,
    Content Types, and libraries must exist in advance.

    Commands:
        $Values = @{"ColumnName"=$VarColValue;"ColumnName"=$VarColValue;"ColumnName"=$VarColValue;}
        $Item = Add-CatMigrationRecord -Url $VarURL -Creds $VarCred -List $VarListName -ContentType $VarContentTypeName -Values $Values

	Parameters
		URL - URL of Site
        Creds - Credentials for access to site
        List - Name of list
        ContentType - Name of content type
        Values - Array of column & value mappings
#>
function global:Add-CatMigrationRecord
{
    param (
        [parameter(Mandatory=$true, Position=0, ValueFromPipeline=$true)]
        [String]$Url,
        
        [parameter(Mandatory=$true, Position=1, ValueFromPipeline=$true)]
        [System.Management.Automation.PSCredential]$Creds,

		[parameter(Mandatory=$true, Position=2, ValueFromPipeline=$true)]
        [String]$L,

        [parameter(Mandatory=$false, Position=3, ValueFromPipeline=$true)]
        [String]$ContentType,
        
		[parameter(Mandatory=$false, Position=4, ValueFromPipeline=$true)]
		[System.Collections.Hashtable]$Values)
    
    # Connecting to Reporting Site
    Connect-PnPOnline -Url $Url -Credentials $Creds
    
    # Adding List Item to Sites List
    Add-PnPListItem -List $L -ContentType $ContentType -Values $Values
} 