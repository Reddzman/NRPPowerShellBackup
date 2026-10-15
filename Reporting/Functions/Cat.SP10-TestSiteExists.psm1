<#  Cat.SP10-TestSiteExists.psm1

This function confirms if a site collection already exists.

Dependencies:
    -   SharePoint 2010 Powershell Modules

Command: Test-CatSP10SiteExists -url $siteURL -type "Site/Web"

Parameters
    url - URL to the desired site
    type - Either "Site" for site collection or "Web" for sub site
#>

function global:Test-CatSP10SiteExists
{ 
	param (
            [string]$url,
            [String]$type
    )
    # Check typ of site
    switch ($type) {
        Site {
            try {
                $temp = Get-SPSite $url -ErrorAction Stop
                Return $true
            }
            catch {
                return $false
            }
        }
        Web {
            try {
                $temp = Get-SPWeb $url -ErrorAction Stop
                Return $true
            }
            catch {
                return $false
            }
        }
        Default {return $false}
    }
}