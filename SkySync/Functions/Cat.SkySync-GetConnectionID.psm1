<# GetConnectionID

This function get the id for a specifc Conneciton based on name.

Sample Call: 
    GetConnectionID $Name
#>
function GetConnectionID($connectionName)
{    
    $results = skysync-cli connections list --search=$connectionName --limit=100 --output-json=true | ConvertFrom-Json
    $result = ""
    if ($results.Count -gt 0)
    {

        if ($results.Count -gt 1)
        {
            $msg = "The search for the connection (" + $connectionName + ") yielded multiple results.  Please enter the name completely or provide the connection ID instead"
            throw $msg
        }
        else
        {
            $result = $results[0].id
        }
    }
    else
    {
        $connection = skysync-cli connections show $connectionName --output-json=true | ConvertFrom-Json

        if (IsNull $connection -eq $true)
        {
            $msg = "Unable to locate connection (" + $connectionName + ")."
            throw $msg
        }
        else
        {
            $result = $connection.id
        }
        
    }

    return $result

}