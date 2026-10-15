<# SPO-SyncUserProfileProperties.ps1
The purpose of this script is set user profile properties in SharePoint Online

-----------------------
Nick Ortiz, Catapult Systems. 
2020-12-22: Rebuilt script to use PNP vs. On-Prem Powershell
2021-08-10: 

####
## Required Modules:
# Install-Module -Name SharePointPnPPowerShellOnline (Required to update SharePoint Online, this is the older module and is best for interactive logins)
# Uninstall-Module -Name SharePointPnPPowerShellOnline -AllVersions -Force (If you need to remove the old module)
#
# Install-Module -Name PnP.PowerShell (Required to update SharePoint Online, this Module requires extra configuration)
# Install-WindowsFeature RSAT-AD-PowerShell (Required to read Active Directory)

#####
## Create App Registration and Certificate (Needed if you want to use cert or client secret authentication)
# Initialize-PnPPowerShellAuthentication -ApplicationName "PNPPRD" -Tenant "contoso.onmicrosoft.com" -OutPath c:\temp

** It is currently not possible to modify user profile properties when authenticating with a certificate. **

## Add API Permissions:
# - SharePoint: Sites.FullControl.All

#####
## Add App Permissions
# https://tenantName-admin.sharepoint.com/_layouts/15/appinv.aspx
#
# <AppPermissionRequests AllowAppOnlyPolicy="true">
#   <AppPermissionRequest Scope="http://sharepoint/content/tenant" Right="FullControl" />
#   <AppPermissionRequest Scope="http://sharepoint/social/tenant" Right="FullControl" />
# </AppPermissionRequests>

#####
## See internal attribute names
# $u = Get-AzureADUser -Filter "UserPrincipalName eq 'user@contoso.com'" | Select ExtensionProperty
# $u.ExtensionProperty

You should run this script atleast once in a console to generate your password file.

Examples: 

    .\SPO-SyncUserProfileProperties.ps1 -tenantName "YourCompany" -dirSource "AAD" -mapLocationLocal

Options:
    -tenantName (string) : Tenent name for your SPO site i.e. the 'yourcompany' part of your SharePoint Online Url - https://yourcompany.sharepoint.com
    -dataMapPath (string) : Local or Network path to data file.
    -username (string) : user account that will be authenticating.
    
#>

param
(
    [Parameter(Mandatory=$false)] [string] $tenantName,
    [Parameter(Mandatory=$true,HelpMessage="Choose a Directory Source. Enter 'AD' for Active Directory or 'AAD' for AzureAD")] [string] $dirSource,
    [Parameter(Mandatory=$true)] [string] $dataMapPath,
    [Parameter(Mandatory=$false)] [string] $username,
    [Parameter(Mandatory=$false)] [string] $authMethod,
    [Parameter(Mandatory=$false)] [switch] $mapLocationNetwork,
    [Parameter(Mandatory=$false)] [switch] $mapLocationSPO,
    [Parameter(Mandatory=$false)] [switch] $mapLocationLocal
)
if(!(Get-Module SharePointPnPPowerShellOnline)) { Import-module SharePointPnPPowerShellOnline }


Write-Output "Starting Transcription"
$StartDateTime = Get-Date -Format "yyyyMMddHHmmss"
Start-Transcript -Path "$PSScriptRoot\logs\$startDateTime-ProfilePropertyImport.log"

## Configuration Settings
$useAzureAuth = $false
$usePWFile = $true

## Authentication Method:
# web, file and secret are supported, cert can not update UPS properties currently.
$authMethod = "web" 

$azureAuthAcc = "AutomationPSCredential"

## Authentication using a Cert or Client Secret Settings
$AppId = ""
$ClientSecret = ""

## Authentication using a Cert Settings
$certThumb = ""

## Authentication using a file Settings
if(!$username) {
    $username = ""
}
$PWFilePath = ".\creds"
$AESKeyFileName = "Cat.SPO-SyncUserProfileProperties.aes"
$PWFileName = "Cat.SPO-SyncUserProfileProperties.pw"

## SharePoint Tenant Url       
if(!$tenantName){     
    $tenantName = "contosos"
}

## Check for maplocation
if(!$mapLocationNetwork -and !$mapLocationSPO -and !$mapLocationLocal){
    Write-Host "You must specify if the map file is locatied locally (-mapLocationLocal), on a newtwork share (-mapLocationNetwork) or in a SharePoint Online library (-mapLocationSPO)." -ForegroundColor "Red"
    Exit
    Stop-Transcript
}

## Check for datamap path
if(!$dataMapPath) {
    Write-Host "You must specify patch and ti you map file. i.e. \\servername\share\folder or \SyncLibrary for a SharePoint library" -ForegroundColor "Red"
    Exit
    Stop-Transcript
}

## End Configuration Settings

# DataMapFile that contains a 1 to 1 mapping for AAD to SPO properties.
$dataMapFile = "Cat.SPO-SyncUserProfileProperties.map.csv"

## Tenant Urls
$siteUrl = "https://$tenantName.sharepoint.com"
$tenantAdminUrl = "https://$tenantName-admin.sharepoint.com"

if($authMethod -eq "file") {
   
    ## Check for and Create password file
    if(!(Test-Path $PWFilePath)) {
        $newPWFileDir = New-Item -Path $PWFilePath -ItemType "directory"
    }

    if(!(Test-Path  "$PWFilePath\$PWFileName")) {

        ## Check for and Create AES Key
        if(!(Test-Path  "$PWFilePath\$AESKeyFileName")) {
            write-host "AES Key file not found, generating ..." -ForegroundColor "Yellow"
            $AESKeyFilePath = "$PWFilePath\$AESKeyFileName"
            $AESKey = New-Object Byte[] 32
            [Security.Cryptography.RNGCryptoServiceProvider]::Create().GetBytes($AESKey)
            
            Set-Content $AESKeyFilePath $AESKey

        }

        $passwordfilePath = "$PWFilePath\$PWFileName"
        
        write-host "Password file not found, Please enter the password for '$username'" -ForegroundColor "Yellow"
        read-host -assecurestring -Prompt 'Enter Password' | convertfrom-securestring -Key $AESKey | out-file $passwordfilePath
    } else {
        Write-Verbose "Password File found ..."
    }

    ## Credential using Password file
    $AESKeyFile = Get-Content $PWFilePath\$AESKeyFileName
    $SPOPass = Get-Content $PWFilePath\$PWFileName | ConvertTo-SecureString -Key $AESKeyFile

    $credential = new-object System.Management.Automation.PSCredential ($username,$SPOPass)
} elseif($authMethod -eq "azureAuth") {

    ## Credential Azure Store Credntial
    $credential = Get-AutomationPSCredential -Name $azureAuthAcc
    
} elseif($authMethod -eq "web" -or $authMethod -eq "secret") {

    ## Credential Azure Store Credntial
    
} else {

    ## Credential User Prompted
    $credential = (Get-Credential)
}

## Connect to AzureAd

if($dirSource -eq 'AAD'){
    Write-Verbose "Connecting to AzureAD"
    $connAz = Connect-AzureAd
}

Write-Verbose "Connecting to $siteUrl"
## Connect to SPO
if($authMethod -eq 'web') {
    $connAdm = Connect-PnPOnline -Url $tenantAdminUrl -UseWebLogin -ReturnConnection -WarningAction ignore
    $connSP = Connect-PnPOnline -Url $siteUrl -UseWebLogin -ReturnConnection -WarningAction ignore
} elseif($authMethod -eq 'secret') {
    $connAdm = Connect-PnPOnline -Url $tenantAdminUrl -ClientID $AppId -ClientSecret $ClientSecret -WarningAction ignore -ReturnConnection 
    $connSP = Connect-PnPOnline -Url $siteUrl -ClientID $AppId -ClientSecret $ClientSecret -WarningAction ignore  -ReturnConnection 
} elseif($authMethod -eq 'cert') {
    $connAdm = Connect-PnPOnline -Url $tenantAdminUrl -ClientID $AppId -Tenant "$tenantName.onmicrosoft.com" -Thumbprint $certThumb  -WarningAction ignore -ReturnConnection 
    $connSP = Connect-PnPOnline -Url $siteUrl -ClientID $AppId -Tenant "$tenantName.onmicrosoft.com" -Thumbprint $certThumb -WarningAction ignore -ReturnConnection 
} elseif($authMethod -eq 'file') {
    $connAdm = Connect-PnPOnline -Url $tenantAdminUrl -Credentials $credential -ReturnConnection -WarningAction ignore
    $connSP = Connect-PnPOnline -Url $siteUrl -Credentials $credential -ReturnConnection -WarningAction ignore
}

# $connAz = Connect-AzureAd -Credential $credential
# $connSP = Connect-PnPOnline -url $siteUrl -Credentials $credential -ReturnConnection
# $connAdm = Connect-PnPOnline -url $tenantAdminUrl -Credentials $credential -ReturnConnection

## Load CSV data from local/network file
if($mapLocationNetwork) {
    
    Write-Verbose "Loading data file '$dataMapPath\$dataMapFile' ..."
    
    ## Load CSV data from local/network file
    $newDrive = New-PSDrive -Name "Q" -PSProvider FileSystem -Root $dataFilePath -Credential $credential
    $filter="*.csv"
    $latest = Get-ChildItem -Path "Q:\" -Filter $filter | Sort-Object CreationTime -Descending | Select-Object -First 1
    $latest.name
    $CSVPath = "Q:\" + $latest.Name
    # write-host $CSVPath

    $CSV = Import-CSV -Path $CSVPath
    
}

if($mapLocationLocal) {

    Write-Verbose "Loading data file '$dataMapPath\$dataMapFile' ..."

    ## Load CSV data from local/network file
    $CSV = Import-CSV -Path $dataMapPath\$dataMapFile
    
}

## Download data file from SharePoint Online
if($mapLocationSPO) {
    
    Write-Verbose "Loading data file '$siteUrl$csvLibPath/$dataMapFile' ..."
    
    ## Save data from data file store in SharePoint Online
    $latest = Get-PnpFile -Url $csvLibPath/$dataMapFile -AsString -Connection $connSP

    ## Convert Streamed data file to CSV
    $CSV = ConvertFrom-Csv $latest
}

Write-Verbose "Loading users ..."

# Load users from Active Directory or Azure Active Directory
if($dirSource -eq "AD") {
    $users = Get-ADUser -filter * -Properties *
} elseif($dirSource -eq "AAD") {
    $users = Get-AzureADUser -All:$true
}

Write-Verbose "Processing users ..."

if($users){
    # $bulkUpdateJSON = '{
    # "value": ['
    Write-Verbose "$($users.Count) users loaded..."

    foreach($user in $users) {

        if($null -ne $user.UserPrincipalName){
            $UPN = "i:0#.f|membership|"+$user.UserPrincipalName

            Write-Verbose "Loading User Properties for '$UPN' ..."
            $upsProps = Get-PnPUserProfileProperty -Account $UPN -Connection $connAdm

            if($null -ne $upsProps.DisplayName) {
                # $userUpdateJSON = '
                # {
                #     "IdName": "'+$($user.UserPrincipalName)+'",
                #     '
        
                ForEach($row in $CSV) {
                    $srcProperty = $row.AzureADProperty
                    $dstProperty = $row.SharePointProperty

                    if($dirSource -eq 'AAD') {
                        $srcPropertyVal = $user.ExtensionProperty.$srcProperty
                    } else {
                        $srcPropertyVal = $user.$srcProperty
                    }

                    if($null -ne $srcPropertyVal) {
                        # $userUpdateJSON += '"'+$($dstProperty)+'": "'+$($srcPropertyVal)+'",'
                        
                        Write-Verbose "- Updating '$dstProperty' for user '$($user.UserPrincipalName)' to '$srcPropertyVal'"
                        # Write-Verbose "Adding '$dstProperty' for user $($user.UserPrincipalName) to bulk import"
                        Set-PnPUserProfileProperty -Account $UPN -PropertyName $dstProperty -Value $srcPropertyVal -Connection $connAdm
                    } else {
                        Write-Verbose "- No value for '$srcProperty' for user '$($user.UserPrincipalName)'"
                    }

                } 

                # $userUpdateJSON += '
                # },'
            } else {
                Write-Verbose "- Unable to load profile for '$UPN' ..."
            }
        }
        # $bulkUpdateJSON += $userUpdateJSON

    }

    # $bulkUpdateJSON += ']
    # }'

    # $bulkUpdateJSON > param.json
    # Write-Verbose "Starting Bulk Import"
    # New-PnPUPABulkImportJob -Folder "synclib" -Path param.json -IdProperty "IdName" -UserProfilePropertyMapping @{"SPS-HireDate"="SPS-HireDate"} -Connection $connSP 

} else {
    Write-Verbose "No Users Found!"
}

Write-Verbose "Data import complete"

## Disconnect from Mapped Drive, Azure and SPO
if($mapLocationNetwork) {
    Remove-PSDrive -Name "Q" -Credential $credential
}
Write-Verbose "Disconnecting from Azure, SPO and SPO-Admin"
if($dirSource -eq 'AAD'){
    Disconnect-AzureAD
}
Disconnect-PnPOnline -Connection $connSP
Disconnect-PnPOnline -Connection $connAdm

Stop-Transcript