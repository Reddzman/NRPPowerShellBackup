<#  Cat.PnP-TestSiteExists.psm1

This function confirms if a site collection or Web already exists.

Dependencies:
    -   SharePoint PnP PowerShell - https://docs.microsoft.com/en-us/powershell/sharepoint/sharepoint-pnp/sharepoint-pnp-cmdlets?view=sharepoint-ps

Command: Test-CatPnPSiteExists -url $siteURL

Parameters
    url - URL to the desired site
    type - Either "Site" for site collection or "web" for sub site
#>

function global:Test-CatPnPSiteExists
{ 
	param (
        [string] $url,
        [pscredential] $cred
    )
    # Check typ of site
        try {
            $tempConnect = Connect-PnPOnline -Url $url -Credentials $cred
            return $true
        }
        catch {
            return $false
        }
}

function global:Test-CatPnPSiteExistsAppReg
{ 
	param (
        [string] $fnUrl,
        [string] $fnAppID,
        [string] $fnTenant,
        [string] $fnCertPath,
        [System.Security.SecureString] $fnCertPass
    )
    # Check type of site
        try {            
            #for some reason this does not fail when connecting to 
            #an invalid site when using app registrations?
            $fnCert = Get-PnPAzureCertificate -CertificatePath $fnCertPath -CertificatePassword $fnCertPass
            $fnConnect = Connect-PnPOnline -ClientId $fnAppID -Url $fnUrl -Tenant $fnTenant -Thumbprint $fnCert.Thumbprint
            $fnSite = Get-PnPSite -Includes RootWeb -Connection $tempConnect -ErrorAction Stop
            return $true
            Disconnect-PnPOnline -Connection $fnConnect
        }
        catch {

            return $false
        }
        finally {
            # Disconnecting PnP
            
        }
}