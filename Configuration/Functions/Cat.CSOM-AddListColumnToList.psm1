<#  Add-ListColumnToListCT (PnP & CSOM)
This function adds a single list column to a content type on a list. This function expects a client context
as a parameter.

Command: Add-ListColumnToListCT -ClientContext $VarCtx -ContentType $VarCT -ListColumn $VarField

Parameters
    ClientContext - PnP client context
    ContentType - Content type object from list
    ListColumn - Field object from list
#>

function global:Add-ListColumnToListCT
{ 
	param 
        ( 
            [parameter(Mandatory=$true, Position=0, ValueFromPipeline=$true)]
            [Microsoft.SharePoint.Client.ClientContext]$ClientContext,
            
            [parameter(Mandatory=$true, Position=0, ValueFromPipeline=$true)]
            [Microsoft.SharePoint.Client.ContentType]$ContentType,

            [Parameter(Mandatory=$true,Position=1)] 
            [Microsoft.SharePoint.Client.Field]$ListColumn
        )
    try 
    {
        # Creating Field Link
        $FieldLink = New-Object Microsoft.SharePoint.Client.FieldLinkCreationInformation
        $FieldLink.Field = $ListColumn

        # Get ContentType Fields
        
        
        # Updating Content type
        [Void]$ContentType.FieldLinks.Add($FieldLink)
        $ContentType.Update($false)
        $ClientContext.ExecuteQuery()

        # Successfull Output
        Write-host ("               {0} has been added to {1}" -f $ListColumn.Title, $ContentType.Name)
    }
    catch 
    {
        Write-host ("               Failed to add {0} to {1}" -f $ListColumn.Title, $ContentType.Name) -ForegroundColor Red
    }
    
}