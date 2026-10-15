<# Wait-BackgroundJobs

This function monitors the server to wait for a free powershell session.

#>
function global:Wait-BackgroundJobs{
	$Jobs = Get-Job | ?{$_.State -eq "Running"}
	
	Write-Host ("   Waiting on in progress jobs: {0}..." -f $Jobs.Count) -NoNewline

    while ($Jobs.Count -gt 0)
    {
        Write-Host "." -NoNewLine
        Start-Sleep -Seconds 15
		$Jobs = Get-Job | Where-Object{$_.State -eq "Running"}
    }
    Write-Host  "Finished waiting for jobs"
} # Wait-BackgroundJobs