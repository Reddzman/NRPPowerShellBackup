<#
Cat.PnP-FindListColumnByName.ps1

This script is used to find a list column by internal name.  It will iterate through all sites by the selected
list template and write to the screen if it finds the column. This was originally created to locate a crawled property
that was being used in search, but we couldn't find it in the tenant.  

Dependencies
	-	PnP PowerShell https://msdn.microsoft.com/en-us/pnp_powershell/pnp-powershell-overview
    -   This script assumes SharePoint admin credentials. 
#>

#Parameters
$TenantAdminURL = "https://[tenant]-admin.sharepoint.com"
$ColumnName = "[InternalColumnName]"

#Connect to Tenant Admin Site
$creds = Get-Credential
Connect-SPOService -url $TenantAdminURL -Credential $creds
 
#Uncomment the site types that you want to look in. 
#Get All Communication Sites
$sites = Get-SPOSite -Template "SITEPAGEPUBLISHING#0" -Limit ALL

#Get All Team Sites
#$sites = Get-SPOSite -Template "GROUP#0" -Limit ALL

foreach ($site in $sites) {
  Write-Host "Site URL: ", $site.Url
    Try {
    Connect-PnPOnline -Url $site.Url -Credentials $creds
    $lists = Get-PnPList | Where { $_.Hidden -eq $false }
    $siteTitle = $site.Title
    foreach ($list in $lists) {
      #Write-Host "List Name: ", $list.Title
      $title = $list.Title    
      $fields = Get-PnPField -List $list.Title
      foreach ($field in $fields) {
        if ($field.InternalName -eq $ColumnName) {
          Write-Host "Found it! It's in the list $title in site $siteTitle."
        }
      }
    }
  }
  Catch {
    Write-Host "Cannot access site."
  }
}
