<# JobExists

This function validates if a job already exists with a specific name.

Sample Call: 
    JobExists $Name

#>
function JobExists($jobName){
    $result = skysync-cli jobs list --search $jobName --json | ConvertFrom-Json

    if ($result.Count -gt 0)
    {
        return $true
    }
    else
    {
        return $false
    }
}