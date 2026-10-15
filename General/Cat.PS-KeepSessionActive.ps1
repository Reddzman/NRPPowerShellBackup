#region Documentation
<#  Cat.PS-KeepSessionActive.ps1

This script toggles the scroll lock on a perticulary duration of time in order to keep a windows session active. This has been
usefull for trying to use long running scripts through VDI sessions.

It has been tested with the VMWare Horizon client.

Call script with parameters
	.\Cat.PS-KeepSessionActive.ps1  -sleep 240 -intervals 10

Table of Contents
	1.	Variables
	2.	Main Script
#>
#endregion

#region Script Parameters
param(	$sleep = 300, # second
		$intervals = 12 # loops
)
#endregion

#region Start Date / Time
# Remove Existing Catapult Modules to ensure PS session has most up to date functions
Get-Module Cat.* | Remove-Module

# Load General Functions
Import-Module "..\General\Cat.GeneralFunctions.psm1"

# Get Start Date / Time
$scriptStart = Start-CatScript
#endregion

#region 1.	Variables
# Output Parameters & Variables
Write-Host
Write-Host "Script Parameters" -ForegroundColor Cyan
Write-Host ("   Sleep: {0}" -f $sleep)
Write-Host ("   Interval: {0}" -f $intervals)
#endregion
    
#region 2.	Main Script
Write-Host
Write-Host "Keeping Windows Session active." -ForegroundColor Magenta
Write-Host
Write-Host ("   Start time: {0}" -f $scriptStart[0])
Write-Host "   ." -NoNewline

# Creating Shell Object
$WShell = New-Object -com "Wscript.Shell"

# Create Stop watch
# Some environments don't support invocation of this method.
$stopwatch
try {
    $stopwatch = [system.diagnostics.stopwatch]::StartNew()
} catch {
   Write-Host "   Couldn't start the stopwatch." -ForegroundColor Red
}

# Infinte loop to keep session open.  Must use cntr-C to stop script
$index = 0
while ($true){
    
	# Toggle Scroll Lock
	Write-Host "." -NoNewLine
    $WShell.sendkeys("{SCROLLLOCK}")

	# Waiting on Scroll lock
    Start-Sleep -Milliseconds 200

	# Toggle Scroll lock again
	Write-Host "." -NoNewLine
    $WShell.sendkeys("{SCROLLLOCK}")

	# Sleep script
    Start-Sleep -Seconds $sleep

    # Output date time & elapsed time
    if ( $stopwatch.IsRunning -and (++$index % $intervals) -eq 0 )
    {
        Write-Host
		Write-Host
		Write-Host ("   Current Time: {0} Elapsed time: {1}" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"), $stopwatch.Elapsed.ToString('dd\.hh\:mm\:ss'))
		Write-Host "   ." -NoNewline
    }
} 
#endregion