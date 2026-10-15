<# Wait-SiteCreation

This function monitors the SharePoint Online to wait for a site to be created.

Parameters
    Url - URL of site being created
    Cred - Credentials for accessing site
    Sleep - # of seconds to wait between checking.

#>
function global:Wait-SiteCreation {
	param(	
        [String] $Url,
        [pscredential]$Cred,
        [int]$Sleep
	)
	# Loop until able to connect to stie
	$Complete = $false
	While (-not $Complete)
	{
		Try
		{
			Connect-PnPOnline -Url $Url -Credentials $Cred
			$Complete = $true
		} 
		Catch
		{
			$Complete = $false
			Write-Host "." -NoNewline
			#Write-Host ". Exception Message: $($_.Exception.Message)" # use to troubleshoot connection issues
			Start-Sleep -Seconds $Sleep
		} 
	}
	Write-Host
} #Wait-SiteCreation