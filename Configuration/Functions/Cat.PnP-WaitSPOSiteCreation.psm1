<#  Cat.PnP-WaitSPOSiteCreation.psm1
This function waits for a site collection creation in SharePoint Online to finish by attempting to connect
via PnP PowerShell

Command: Wait-CatSPOSiteCreation -Creds $VarCred -URL $VarURL

Parameters
    Creds - SPO Credentials that should be used to connect
    URL - URL of Site Collection being created
#>

function global:Wait-CatSPOSiteCreation
{ 
	param 
	( 
		[parameter(Mandatory=$true, Position=0, ValueFromPipeline=$true)]
		[System.Management.Automation.PSCredential]$Creds,

		[Parameter(Mandatory=$true,Position=1)] 
		[string]$URL
    )
    
    [boolean]$SiteComplete = $false
    [int]$SiteCreateWait = 30   # How long to wait between connection attempts on the new site collection

    While (-not $SiteComplete)
    {
        Try
        {
            Connect-PnPOnline -Url $URL -Credentials $Creds
            $SiteComplete = $true
        } # Try Connect-PnPOnline
        Catch
        {
            $SiteComplete = $false
            Write-Host "." -NoNewline -ForegroundColor Yellow
            Start-Sleep -Seconds $SiteCreateWait
        } # Catch Connect-PnPOnline
    } # While ($ConnectError -ne $null)
    Write-Host "."
} # Wait-CatSPOSiteCreation