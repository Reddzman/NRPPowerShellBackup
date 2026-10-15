# .\Cat.SG-SPOFolderSize.ps1 -jobFile "Samples\ShareGate\Cat.SG-SPOFolderSize-test-csv" -clientTenant "nrpdev2" -appID "caf6b11a-cf03-416c-b0f0-67a09a8e30ee" -certPath "C:\GitHub\NRPDev\AppRegistration\PowerShell App Registration.pfx"

Param(
    [String] $jobFile = "Cat.SG-SPOnPremListMigrations-Client.csv",
    [string] $clientTenant = "sailpoint",
    [String] $appID = "{Azure App Registration ID}",
    [String] $certPath = "C:\certs\PnPAppRegPowerShell.pfx"
)

# Load PnP Modules
Get-Module -ListAvailable | Where-Object{($_.Name -like "*SharePoint*") -or ($_.Name -like "*PnP*") } | Remove-Module
Import-Module PnP.PowerShell -WarningAction SilentlyContinue

# Input CSV File
$inPath = ("..\_Input Files\{0}" -f $jobFile)
If (Test-Path ($inPath)){
	$inFile = Import-Csv $inPath -Encoding Unicode
}
else {
	Write-Host
	throw ("The jobs file does not exist: {0}" -f $inPath)	
}

# Output file
$output = ("..\_Output Files\{0}.SG-SPOFolderSize.csv" -f $clientTenant)

#Variables
$domain = ("{0}.onmicrosoft.com" -f $clientTenant)
$certPass = ConvertTo-SecureString -String (Read-Host -Prompt "   Enter certificate password for the app registration") -AsPlainText -Force
$cert = Get-PnPAzureCertificate -Path $certPath -Password $certPass

# Looping CSV File
foreach ($job in $inFile) {
    # Build Folder URL
    if ($job.folderName) {
        $folderURL = ($job.siteURL + "/" + $job.libaryPath + "/" + $job.folderName)
    } Else {
        $folderURL = ($job.siteURL + "/" + $job.libaryPath)
    }

    # Content to Site
    Write-Host ("Collecting data on {0}: {1}" -f $job.jobTitle, $folderURL)
    Connect-PnPOnline -ClientId $appID -Url $job.siteURL -Tenant $domain -Thumbprint $cert.Thumbprint

    # Get Folder Object
    $folder = Get-PnPFolder -Url $folderURL -Includes StorageMetrics

    # Create CSV row
    $row = $job
    $row | Add-Member -MemberType NoteProperty -Name "SPOItemCOunt" -Value $folder.ItemCount
    $row | Add-Member -MemberType NoteProperty -Name "SPOSizeBytes" -Value $folder.StorageMetrics.TotalSize

    # Output to CSV
    $row | Export-Csv -Path $output -NoTypeInformation -Append

    # Cliean up variables
    Remove-Variable folderURL
    Remove-Variable folder
    Remove-Variable row

    # Disconnecting from Site
    Disconnect-PnPOnline    
}