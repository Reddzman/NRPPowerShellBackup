<# Wait-SystemResourcesFree

This function monitors the server every 15 seconds to determine if enough system resources are free.

#>
function global:Wait-SystemResourcesFree{
	$Counter = Get-Counter -Counter "\Processor(_Total)\% Processor Time" -MaxSamples 1
	$MemoryCounter = Get-Counter -Counter "\Memory\Available MBytes"

	Write-Host ("      CPU: {0}  Memory: {1} - Waiting for resources..." -f $Counter.CounterSamples.CookedValue, $MemoryCounter.CounterSamples.CookedValue) -NoNewline

	while (($Counter.CounterSamples.CookedValue -ge 70) -or ($MemoryCounter.CounterSamples.CookedValue -le 2048))
	{
		Write-Host"." -NoNewline
		Start-Sleep 15
		$Counter = Get-Counter -Counter "\Processor(_Total)\% Processor Time" -MaxSamples 1
		$MemoryCounter = Get-Counter -Counter "\Memory\Available MBytes"
	}		
	Write-Host  "Finished waiting for resources"
} # Wait-SystemResourcesFree