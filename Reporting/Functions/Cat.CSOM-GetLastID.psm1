<#  Get-LastID 
This function uses CSOM to get the last SharePoint List ID used on a list. PnP PowerShell currently has limitations
related to RowLimit.

Command: Get-LastID -context $context -list $csomList

Parameters
    context - PnP client context
    listID - Name of the list to query
#>

function global:Get-LastID
{ 
	param ( 
        [Microsoft.SharePoint.Client.ClientContext]$context,
        [string]$listID
    )
    
    #Define the CAML Query
    $camlQuery = New-Object Microsoft.SharePoint.Client.CamlQuery
    <# This query includes folders but only gets the top level.  It avoids List View Threshold errors.
    # You may need to upload a temporary document to get the right ID.
    $camlQuery.ViewXml = "@<View Scope='RecursiveAll'>
            <Query>
                <OrderBy>
                    <FieldRef Name='ID' Ascending='False'/>
                </OrderBy>
            </Query>
            <RowLimit>1</RowLimit>
        </View>" #>
    # This query ignores folders and is the better of the two queries, but can throw a List view threshold 
    # error on large lists. Column indexing can potentially help resolve these errors.
    $camlQuery.ViewXml = "@<View Scope='RecursiveAll'>
            <Query>
                <Where>
                    <Eq>
                        <FieldRef Name='FSObjType' />
                        <Value Type='Integer'>0</Value>
                    </Eq>
                </Where>
                <OrderBy>
                    <FieldRef Name='ID' Ascending='False'/>
                </OrderBy>
            </Query>
            <RowLimit>1</RowLimit>
        </View>" #>
    
    try 
    {
        #Get the List
        $csomList = $context.Web.lists.GetById($listID)

        #Get All List Items matching the query
        $csomItem = $csomList.GetItems($camlQuery)
        $context.Load($csomItem)
        $context.ExecuteQuery()

        $csomItem.Id
    }
    catch 
    {
        Write-host ("Failed to query list: {0}" -f $csomList.Title) -ForegroundColor Red
        Write-host ("Message: {0}" -f $_.Exception.Message) -ForegroundColor Red
    }    
}