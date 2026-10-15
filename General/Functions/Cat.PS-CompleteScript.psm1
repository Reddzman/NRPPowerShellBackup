<# Complete-CatScript

This function uses the start time array to calculate and output duration at the end of a script.

Sample Call: 
    Complete-CatScript $scriptStart[0]

#>
function Complete-CatScript{
	param(
        [parameter(Mandatory=$true, Position=0, ValueFromPipeline=$true)][datetime]$STime
    )
	
	# Finish Time & Time Taken
	[datetime]$FTime = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
	$TTime = New-TimeSpan -Start $STime -End $FTime

	# Screen Output
	Write-Host
	Write-Host "-----------------------------------------------------------------------------------------------------" -ForegroundColor Cyan
	Write-Host ("Started: {0}   Finished: {1}   Time Taken: {2}" -f $STime, $FTime, $TTime) -ForegroundColor Cyan
	Write-Host "-----------------------------------------------------------------------------------------------------" -ForegroundColor Cyan
	Write-Host

	# Stopping Transcript
	#Stop-Transcript
} # Complete-CatScript