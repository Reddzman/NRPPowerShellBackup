<# Start-CatScript

This function returns the script start time as an array.  The [0] member of the array is normal date time
format for use in the script and the [1] member is formated for file names.

Sample Call:
    $scriptStart = Start-CatScript

#>
function Start-CatScript{
	# Collecting Start Time
	$SDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
	$FDate = Get-Date -Format "yyMMddHHmm"
	
	# Return Start Time
	$DateArray = @()
	$DateArray += $SDate
	$DateArray += $FDate
	return $DateArray
} # Start-CatScript