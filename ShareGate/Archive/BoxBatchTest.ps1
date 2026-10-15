<#  We are trying to migrate a specific path from box. The code provided by ShareGate didn't do that so
    I grabbed the example that included "-SourceFolder" from the ShareGate documentation.
    https://documentation.sharegate.com/hc/en-us/articles/115000343993-Import-from-Box-com #>

# Variables
$boxAdmin = "Box.Admin@SailPoint.com"
$inputPath = "C:\GitHub\CatPowerShell\_Input Files\Samples\ShareGate\Cat.SG-BoxFolderMigrations-Test.csv"
$rootSiteURL = "https://sailpoint.sharepoint.com"

# CSV Data import
$inputCSV = Import-Csv $inputPath -Encoding Unicode

# Tenant Authentication
$box = Connect-Box -Email $boxAdmin -Admin
$m365 = Connect-Site -Url $rootSiteURL -Browser -DisableSSO

#Loop CSV
foreach ($job in $inputCSV) {

    # Connect to Destination site/library
    $dstSite = $dstSite = Connect-Site -Url $job.dstURL -UseCredentialsFrom $m365
    $dstList = Get-List -Name $job.dstLibrary -Site $dstSite

    # Run ShareGate Job with or without a folder
    If ($job.dstFolder) {
        Import-BoxDocument -Box $box -SourceFolder $job.srcPath -DestinationList $dstList -DestinationFolder $job.dstFolder
	} else {
        Import-BoxDocument -Box $box -SourceFolder $job.srcPath -DestinationList $dstList
	}
}