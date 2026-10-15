<# Wait-PSSession

This function monitors the server to wait for a free powershell session.

#>
function global:Wait-PSSession{
	param(
		[parameter(Mandatory=$true, Position=0, ValueFromPipeline=$true)][int]$MaxSessions
	)
	
	$OpenSessions = (Get-Process | ? {$_.ProcessName -eq "powershell"}).Count

	Write-Host ("      Waiting on current sessions: {0}..." -f $OpenSessions) -NoNewline

    while ($OpenSessions -ge $MaxSessions)
    {
        Write-Host -NoNewLine "."
        Start-Sleep -Seconds 15
		$OpenSessions = (Get-Process | ? {$_.ProcessName -eq "powershell"}).Count
    }
    Write-Host  "Finished waiting for sessions"
} # Wait-PSSession