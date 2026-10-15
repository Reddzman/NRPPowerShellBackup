$siteURL = "https://prologis.sharepoint.com/teams/transactions_JP_2018_007"
$libraryName = "Documents"
$termGroup = "80ec1e02-23c6-4b42-a18b-cf58aab2204d"
$termSet = "937656a9-8691-4643-9f97-31471bf514fc"
$contentType = "0x0120"
$cred = Get-Credential

# Connect to Site
Connect-PnPOnline $siteURL -Credentials $cred

# Get Document Library
$library = Get-PnPList -Identity $libraryName

# Get child managed metadata terms
$parentTerms = Get-PnPTerm -TermGroup $termGroup -TermSet $termSet -IncludeChildTerms -Recursive

# Iterate through Parent Terms\
Write-Host "-------------------------------------------"
Write-Host "Creating Folders"
Write-Host "-------------------------------------------"
ForEach ($pTerm in $parentTerms){
	$folderPath = ("{0}/{1}" -f $library.RootFolder.Name,$pTerm.Name)
	$temp = Resolve-PnPFolder -SiteRelativePath $folderPath
	Write-Host $folderPath

	# Iterate through Child Terms
	if ($pTerm.TermsCount -gt 0) {
		ForEach ($cTerm in $pTerm.Terms){
			$folderPath = ("{0}/{1}/{2}" -f $library.RootFolder.Name,$pTerm.Name,$cTerm.Name)
			$Temp = Resolve-PnPFolder -SiteRelativePath $folderPath
			Write-Host $folderPath
		}
	}
}
