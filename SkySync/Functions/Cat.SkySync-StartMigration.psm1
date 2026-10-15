<# Start-SkySyncMigration

This function converst and schedules simulation jobs to regular jobs.
    -   Disables Simulation Mode
    -   Sets Schedule based JSON
    -   Starts Job

Sample Call: 
    Start-SkySyncMigration -jobs $updateJobs -restURL $baseURL -restHeader $authHeader -schedJSON $scheduleJSON -ct  $contentType

#>
function Start-SkySyncMigration {
    param (
        $jobs,
        [String] $restURL,
        $restHeader,
        $schedJSON,
        [String] $ct
    )
    <# Output Variables for Troubleshooting
    Write-Host "Script Variables" -ForegroundColor Yellow
    Write-Host ("   Schedule JSON: {0}" -f $schedJSON) -ForegroundColor Yellow
    Write-Host ("   Rest URL: {0}" -f $restURL) -ForegroundColor Yellow
    Write-Host ("   Content Type: {0}" -f $ct) -ForegroundColor Yellow 
    #>

    # Looping through Jobs
    foreach ($job in $jobs) {
        Write-Host
        Write-Host ("   Configuring Job: {0}" -f $job.Name) -ForegroundColor DarkGray

        # Disabling Simulation Mode
        Write-Host "      Disabling Simulation Mode"
        $simEndpoint = ("{0}/jobs/{1}?reset=simulation" -f $restURL,$job.ID)
        Try{
            # REST Web Request
            $response = Invoke-RestMethod $simEndpoint -Method PATCH -Headers $restHeader -ContentType $ct
        }
        Catch {
            Write-Host
            Write-Host "!! EXCEPTION OCCURRED !!" -ForegroundColor Red
            Write-Host "!! Exception Type: $($_.Exception.GetType().FullName)" -ForegroundColor Red
            Write-Host "!! Exception Message: $($_.Exception.Message)" -ForegroundColor Red
        }
        
        # Update Job Schedule
        Write-Host "      Setting Schedule"
        $scheduleEndpoint = ("{0}/jobs/{1}?fields=all" -f $restURL,$job.ID)
        Try{
            # REST Web Request
            $response = Invoke-RestMethod $scheduleEndpoint -Method PATCH -Body $schedJSON -Headers $restHeader -ContentType $ct
        }
        Catch {
            Write-Host
            Write-Host "!! EXCEPTION OCCURRED !!" -ForegroundColor Red
            Write-Host "!! Exception Type: $($_.Exception.GetType().FullName)" -ForegroundColor Red
            Write-Host "!! Exception Message: $($_.Exception.Message)" -ForegroundColor Red
        }

        # Starting Job
        Write-Host "      Starting Job"
        $startJob = skysync-cli jobs start --search=$job.Name --output-json=true | ConvertFrom-Json
    }
}