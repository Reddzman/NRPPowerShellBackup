<# connectionExists

This function validates if a connection already exists with a specific name.

Sample Call: 
    connctionExists $Name

#>
function connectionExists{
    param(
        [String] $findName
    )

    $result = skysync-cli connections list --search=$findName --limit=100 --output-json=true | ConvertFrom-Json

    if ($result.Count -gt 0)
    {
        return $true
    }
    else
    {
        return $false
    }   
} # function connectionExists