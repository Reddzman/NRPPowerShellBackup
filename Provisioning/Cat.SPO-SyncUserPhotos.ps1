<# SPO-SyncUserPhotosAIO.ps1
This script will export user photos from Active Directory, Azure Active Directory or Exchange, create multiple sizes of the image,
set the Exchange Online photo (optionally) and push the images into SharePoint User Profile Service.

-----------------------
Nick Ortiz, Quisitive. 
2021-04-29: Initial version, rebuilt as All in one script.
2021-05-06: Add ability to use an AD group to target users
2023-10-04: Remove Exchange Module support and replaced with MS Graph due to EXO module depreciation

####
## Required Modules:
# Remove the Old SharePointPNPPowerShellOnline Module if it's installed.
# Uninstall-Module -Name SharePointPnPPowerShellOnline -AllVersions -Force (If you need to remove the old module)
#
# Install-Module -Name PnP.PowerShell -requiredVersion 1.12.0 (Required to update SharePoint Online, this Module requires extra configuration)
# Install-WindowsFeature RSAT-AD-PowerShell (Required to read Active Directory)
# Install-Module AzureAD (Required to read Azure Active Directory)
# Install-Module Microsoft.Graph (Required to set the O365 User Photo or Use graph as the photo source)

#####
## Create App Registration and Certificate
# Initialize-PnPPowerShellAuthentication -ApplicationName "PNPPRD" -Tenant "contoso.onmicrosoft.com" -OutPath c:\temp

## Add API Permissions:
# - Role: Directory Reader (if using AAD as dirSource or photoSource)
# - Microsoft Graph: User.ReadWrite.All
# - SharePoint: Sites.FullControl.All


## In order to update SharePoint Photos and UPS Profiles, permissions but be added in SharePoint.
#####
## Add App Permissions for UPS
# https://tenantName-admin.sharepoint.com/_layouts/15/appinv.aspx
# Register the AppID from the applicatin created with the Initialize-PnPPowerShellAuthentication command.
# Grants manage permissions to User Profile services.

<AppPermissionRequests AllowAppOnlyPolicy="true">
   <AppPermissionRequest Scope="http://sharepoint/social/tenant" Right="Manage" />
</AppPermissionRequests>

## Add App Permissions for My Site
# https://tenantName-my.sharepoint.com/_layouts/15/appinv.aspx
# Register the AppID from the applicatin created with the Initialize-PnPPowerShellAuthentication command.
# This will grant read access to the root site collection and full control over the top site

<AppPermissionRequests AllowAppOnlyPolicy="true">
   <AppPermissionRequest Scope="http://sharepoint/content/sitecollection/" Right="Read"/>
   <AppPermissionRequest Scope="http://sharepoint/content/sitecollection/web" Right="FullControl"/>
</AppPermissionRequests>
#>

###############################################
##      Start Configuration Settings         ##
###############################################

###############################################
##      Tenant and Directory Settings        ##
###############################################

## O365 Tenant Name
$tenantName = "contoso"
$tenantID = "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"

## Directory Options
# You can use this to specify a specific Domain OU or apply additional user filters
$SearchFilterAD = "*"
$SearchFilterAAD = ""
$SearchBase = "DC=contoso,DC=com"
$MemberOfGroup = ""

## dirSource (AAD/AD)
# Users can be pulled from On-Prem Active Directory (AD) or AzureAD (AAD)
$dirSource = "AD"

###############################################
##          Photo Import Settings            ##
###############################################

## Photo Source (AD/AAD/Graph)
$photoSource = "AD"

## Tempory Photo Storage Page
$tempPhotoStorage = "c:\photos\" 

# This will recreate all phots in your temporary storage.
# If you have already downloaded the photos, you can set to false to skip this step
$generatePhotos = $true

# This will automatically delete generate photos once they've been uploaded.
# If you have already downloaded the photos, you can set to false to skip this step if you're testing.
$cleanUpPhotos = $true

# setMSGraphPhoto: If set to true, photos will be uploaded to Exchange Online
$setMSGraphPhoto = $true

# setSPOPhoto: If set to true, photo will be uploaded to SharePoint Online and the
# user profile property will be updated.
$setSPOPhoto = $true

###############################################
##       Authentication Settings             ##
###############################################

## Authentication Method (web/file/secret/cert):
# web, file and secret are supported, cert can not update UPS photos currently.
$authMethod = "secret" 

## Authentication using a Cert or Client Secret Settings
$AppId = ""
### Client Secret is now considered legacy authentication
$ClientSecret = ""

# Authentication using a Cert Settings
$certThumb = ""

## Authentication using a file Settings
$username = "user@contoso.com"
$PWFilePath = ".\creds"
$AESKeyFileName = "SPO-SyncUserPhotosv2.aes"
$PWFileName = "SPO-SyncUserPhotosv2.pw"

###############################################
##            Alert Settings                 ##
###############################################

$sendEmailAlert = $false

$smtpServerHost = ""
$smtpServerPort = 25
$smtpServerSsl  = $false

$emailFrom      = ''
$emailAddress   = ""
$emailCCAddress = ""
$emailAdmin	= ""
$emailSubject    = "ADPhoto Import REPORT"
# $emailFiles      = "C:\AllScripts\ADPhotos\Logs\UserImagesProcessed.txt"
$emailbody =  "Here is your scheduled $emailSubject "

###############################################
##             Log Settings                  ##
###############################################

# If $true, a single log file will be used otherwise each
# log will be time stamped.

$SingleLog = $true

###############################################
##       End Configuration Settings          ##
###############################################


$timeStamp = (Get-Date).ToString('yyyyMMdd-HHmm')
$ScriptPath = $MyInvocation.MyCommand.Path
$ScriptDir  = Split-Path -Parent $ScriptPath

## Set Log Variables
if($SingleLog) {
    $LogFile = "$ScriptDir\Logs\SPO-SyncUserPhotos.log"
} else {
    $LogFile = "$ScriptDir\Logs\SPO-SyncUserPhotos_$timeStamp.log"
}

## Create Log Directory
if(!(Test-Path $ScriptDir\Logs\)) {
    $newLogsDir = New-Item -Path $ScriptDir\Logs\ -ItemType "directory"
}

## Check Log file size and start a new one if its larger than 20 MB
if((Test-Path $LogFile)) {
    $curLog = Get-ChildItem -Path $LogFile
    $curSize = ($curLog.Length)/1024/1024
    $curLogNewName = $($LogFile.Split('.')[0]) + "$timeStamp." + $($LogFile.Split('.')[1])

    ## Create new log is the current one is greater that 20 MB
    if($curSize -ge "20"){

        Add-content $Logfile -value "---------------------- DevOps Provisioner log has grown too large, creating a new file: $(Get-date -format 'MM/dd/yyy hh:mm:ss tt') -------------------"
        Move-Item -Path $LogFile -Destination $curLogNewName

    }
}

Add-content $Logfile -value "---------------------- Photo Sync Script Started: $(Get-date -format 'dd/MM/yyy hh:mm:ss tt') -------------------"
Add-content $Logfile -value "  Parameters:"
Add-content $Logfile -value "    tenantName: $tenantName"
Add-content $Logfile -value "    tenantID: $tenantID"
Add-content $Logfile -value "    SearchFilterAD: $SearchFilterAD"
Add-content $Logfile -value "    SearchFilterAAD: $AuthMethod"
Add-content $Logfile -value "    SearchBase: $SearchBase"
Add-content $Logfile -value "    MemberOfGroup: $MemberOfGroup"
Add-content $Logfile -value "    dirSource: $dirSource"
Add-content $Logfile -value "    photoSource: $photoSource"
Add-content $Logfile -value "    tempPhotoStorage: $tempPhotoStorage"
Add-content $Logfile -value "    setMSGraphPhoto: $setMSGraphPhoto"
Add-content $Logfile -value "    setSPOPhoto: $setSPOPhoto"
Add-content $Logfile -value "    authMethod: $authMethod"
Add-content $Logfile -value "    AppId: $AppId"
Add-content $Logfile -value "    username: $username"
Add-content $Logfile -value "    PWFilePath: $PWFilePath"
Add-content $Logfile -value "    AESKeyFileName: $AESKeyFileName"
Add-content $Logfile -value "    PWFileName: $PWFileName"

function Resize-Image {
    <#
    .SYNOPSIS
        Resize-Image resizes an image file

    .DESCRIPTION
        This function uses the native .NET API to resize an image file, and optionally save it to a file or display it on the screen. You can specify a scale or a new resolution for the new image.
        
        It supports the following image formats: BMP, GIF, JPEG, PNG, TIFF 
 
    .EXAMPLE
        Resize-Image -InputFile "C:\kitten.jpg" -Display

        Resize the image by 50% and display it on the screen.

    .EXAMPLE
        Resize-Image -InputFile "C:\kitten.jpg" -Width 200 -Height 400 -Display

        Resize the image to a specific size and display it on the screen.

    .EXAMPLE
        Resize-Image -InputFile "C:\kitten.jpg" -Scale 30 -OutputFile "C:\kitten2.jpg"

        Resize the image to 30% of its original size and save it to a new file.

    .LINK
        Author: Patrick Lambert - http://dendory.net
    #>
    Param([Parameter(Mandatory = $true)][string]$InputFile, [string]$OutputFile, [int32]$Width, [int32]$Height, [int32]$Scale, [Switch]$Display)

    # Add System.Drawing assembly
    Add-Type -AssemblyName System.Drawing

    # Open image file
    $img = [System.Drawing.Image]::FromFile((Get-Item $InputFile))

    # Define new resolution
    if ($Width -gt 0) { [int32]$new_width = $Width }
    elseif ($Scale -gt 0) { [int32]$new_width = $img.Width * ($Scale / 100) }
    else { [int32]$new_width = $img.Width / 2 }
    if ($Height -gt 0) { [int32]$new_height = $Height }
    elseif ($Scale -gt 0) { [int32]$new_height = $img.Height * ($Scale / 100) }
    else { [int32]$new_height = $img.Height / 2 }

    # Create empty canvas for the new image
    $img2 = New-Object System.Drawing.Bitmap($new_width, $new_height)

    # Draw new image on the empty canvas
    $graph = [System.Drawing.Graphics]::FromImage($img2)
    $graph.DrawImage($img, 0, 0, $new_width, $new_height)

    # Create window to display the new image
    if ($Display) {
        Add-Type -AssemblyName System.Windows.Forms
        $win = New-Object Windows.Forms.Form
        $box = New-Object Windows.Forms.PictureBox
        $box.Width = $new_width
        $box.Height = $new_height
        $box.Image = $img2
        $win.Controls.Add($box)
        $win.AutoSize = $true
        $win.ShowDialog()
    }

    # Save the image
    if ($OutputFile -ne "") {
        $img2.Save($OutputFile);
    }
}

## Check if photoSource is set to Graph and SetGraphPhoto is true, we can use the source to set the destination
If($photoSource -eq "Graph" -and $setMSGraphPhoto -eq $true) {

    Write-Host "Error! You can not set Graph Photos and use Graph Photos as the photo source." -F Red
    Add-Content $LogFile "Error! You can not set Graph Photos and use Graph Photos as the photo source."
    Exit
}

if($authMethod -eq "file") {
   
    ## Check for and Create password file if needed
    if(!(Test-Path $PWFilePath)) {
        $newPWFileDir = New-Item -Path $PWFilePath -ItemType "directory"
    }

    ## Check for Passworf file
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
} 

## Load PS Modules
if($dirSource -eq "AAD"){
    ## Load Azure AD Modules
    Import-Module AzureAd
} else {
    ## Load On-Prem Active Directory Module.
    Import-Module ActiveDirectory
}

## Load PNP for SharePoint ONline Access
Import-Module PnP.PowerShell

## It is not necesary to Import the Microsoft.Graph 
## PowerShell, if you try the script will become unresponsive

## Set SharePoint Tenant Urls
$adminSiteUrl = "https://$tenantName-admin.sharepoint.com/"
$mySiteUrl =  "https://$tenantName-my.sharepoint.com/"

## Connect to all the things based on authMethod
Switch($authMethod) {
	"web" {
        ## Connect using an Interactive Web Login
		if($dirSource -eq "AAD" -or $photoSource -eq "AAD"){
			Write-Host "Connecting to Azure AD ... " -F Cyan
			Add-Content $Logfile "Connecting to Azure AD ... "
			$aad = Connect-AzureAD
		}
		if($photoSource -eq "Graph" -or $setMSGraphPhoto -eq $true) {
			Write-Host "Connecting to Graph ... " -F Cyan
			Add-Content $Logfile "Connecting to Graph ... "

            try {
			    $graph = Connect-MgGraph
                write-host "done" -ForegroundColor green
            }
            catch {
                Write-Host "failed: $($_.exception.message)" -F Red
                Add-content $Logfile -value " ... failed: $($_.exception.message)"
            }
		}
		Write-Host "Connecting to SharePoint Online ... " -F Cyan -NoNewline
        try {
		    $sptAdmin =  Connect-PnPOnline -Url $adminSiteUrl -UseWebLogin -ReturnConnection -WarningAction ignore
		    $sptMy = Connect-PnPOnline -Url $mySiteUrl  -UseWebLogin -ReturnConnection -WarningAction ignore
            write-host "done" -ForegroundColor green
        }
        catch {
            Write-Host "failed: $($_.exception.message)" -F Red
            Add-content $Logfile -value " ... failed: $($_.exception.message)"
        }
	} 
	"secret" {
        ## Connect using an Application Id / Client Secret. Additional configuration is required for this.
        ## Connecting to AAD with an App ID Requires a certficiate.
		if($dirSource -eq "AAD" -or $photoSource -eq "AAD"){
			Write-Host "Connecting to Azure AD ... " -F Cyan -NoNewline
			Add-Content $Logfile "Connecting to Azure AD ... " -NoNewline
            try {
                $aad = Connect-AzureAD -TenantDomain "$tenantName.onmicrosoft.com" -ApplicationId $AppId -CertificateThumbprint $certThumb
                write-host "done" -ForegroundColor green
            }
            catch{
                Write-Host "failed: $($_.exception.message)" -F Red
                Add-content $Logfile -value " ... failed: $($_.exception.message)"
            }
		}
		if($photoSource -eq "Graph" -or $setMSGraphPhoto -eq $true) {
			Write-Host "Connecting to Graph ... " -F Cyan -NoNewline
			Add-Content $Logfile "Connecting to Graph ... "
			try {
                $ClientSecretConvert = ConvertTo-SecureString -String $ClientSecret -AsPlainText -Force
                $ClientSecretCredential = new-object System.Management.Automation.PSCredential ($AppId,$ClientSecretConvert)
                            
                $graph = Connect-MgGraph -TenantId $tenantID -ClientSecretCredential $ClientSecretCredential 
                write-host "done" -ForegroundColor green
            }
            catch {
                Write-Host "failed: $($_.exception.message)" -F Red
                Add-content $Logfile -value " ... failed: $($_.exception.message)"
            }
		}
		Write-Host "Connecting to SharePoint Online ... " -F Cyan -NoNewline
        try {
		    $sptAdmin = Connect-PnPOnline -Url $adminSiteUrl -ClientID $AppId -ClientSecret $ClientSecret -WarningAction ignore -ReturnConnection 
		    $sptMy = Connect-PnPOnline -Url $mySiteUrl  -ClientID $AppId -ClientSecret $ClientSecret -WarningAction ignore  -ReturnConnection 
            write-host "done" -ForegroundColor green
        }
        catch {
            Write-Host "failed: $($_.exception.message)" -F Red
            Add-content $Logfile -value " ... failed: $($_.exception.message)"
        }
	}
	"cert" {
        ## Connect using an Application Id / Client Certificate. Additional configuration is required for this.
		if($dirSource -eq "AAD" -or $photoSource -eq "AAD"){
			Write-Host "Connecting to Azure AD ... " -F Cyan -NoNewline
			Add-Content $Logfile "Connecting to Azure AD ... "
			try {
			    $aad = Connect-AzureAD -TenantDomain "$tenantName.onmicrosoft.com" -ApplicationId  $AppId -CertificateThumbprint $certThumb
                write-host "done" -ForegroundColor green
            }
            catch {
                Write-Host "failed: $($_.exception.message)" -F Red
                Add-content $Logfile -value " ... failed: $($_.exception.message)"
            }
		}
		if($photoSource -eq "Graph" -or $setMSGraphPhoto -eq $true) {
			Write-Host "Connecting to Graph ... " -F Cyan -NoNewline
			Add-Content $Logfile "Connecting to Graph ... "
			try {
			    $graph = Connect-MgGraph -ClientId $AppId -TenantId $tenantID -CertificateThumbprint $certThumb
                write-host "done" -ForegroundColor green
            }
            catch {
                Write-Host "failed: $($_.exception.message)" -F Red
                Add-content $Logfile -value " ... failed: $($_.exception.message)"
            }
		}
		Write-Host "Connecting to SharePoint Online ..." -F Cyan -NoNewline
        try {
		    $sptAdmin = Connect-PnPOnline -Url $adminSiteUrl -ClientID $AppId -Tenant "$tenantName.onmicrosoft.com" -Thumbprint $certThumb  -WarningAction ignore -ReturnConnection 
		    $sptMy = Connect-PnPOnline -Url $mySiteUrl  -ClientID $AppId -Tenant "$tenantName.onmicrosoft.com" -Thumbprint $certThumb -WarningAction ignore -ReturnConnection 
            write-host "done" -ForegroundColor green
        }
        catch {
            Write-Host "failed: $($_.exception.message)" -F Red
            Add-content $Logfile -value " ... failed: $($_.exception.message)"
        }
	}
	"file" {
        ## Connect using a traditional user name / password. This method is depreciated and is considered Legacy Auth.
		if($dirSource -eq "AAD" -or $photoSource -eq "AAD"){
			Write-Host "Connecting to Azure AD ... " -F Cyan -NoNewline
			Add-Content $Logfile "Connecting to Azure AD ... "
            try { 
			    $aad = Connect-AzureAD -Credentials $credential 
                write-host "done" -ForegroundColor green
            }
            catch {
                Write-Host "failed: $($_.exception.message)" -F Red
                Add-content $Logfile -value " ... failed: $($_.exception.message)"
            }
		}
        if($photoSource -eq "Graph" -or $setMSGraphPhoto -eq $true) {
			Write-Host "Connecting to Graph ... " -F Cyan
			Add-Content $Logfile "Connecting to Graph ... "
			try { 					
                $graph = Connect-MgGraph -TenantId $tenantID -ClientSecretCredential $credential 
                write-host "done" -ForegroundColor green
            }
            catch {
                Write-Host "failed: $($_.exception.message)" -F Red
                Add-content $Logfile -value " ... failed: $($_.exception.message)"
            }
		}
		Write-Host "Connecting to SharePoint Online ..." -F Cyan -NoNewline
        try {
		    $sptAdmin = Connect-PnPOnline -Url $adminSiteUrl -Credentials $credential -ReturnConnection -WarningAction ignore
		    $sptMy = Connect-PnPOnline -Url $mySiteUrl  -Credentials $credential -ReturnConnection -WarningAction ignore
            write-host "done" -ForegroundColor green
        }
        catch {
            Write-Host "failed: $($_.exception.message)" -F Red
            Add-content $Logfile -value " ... failed: $($_.exception.message)"
        }
	}
}

## Get all the employees with AD Pictures
Add-content $Logfile -value "Loading users from $dirSource, this may take a moment ..."
Write-Host "Loading users from $dirSource, this may take a moment ... "  -ForegroundColor Cyan -noNewLine
if($dirSource -eq "AAD" ) {
    ## Load users from Azure AD
	try {
        if($SearchFilterAAD) {
		    $users = Get-AzureADUser -All:$true -Filter $SearchFilterAAD
        } else  {
            $users = Get-AzureADUser -All:$true
        }
        
		write-host "done" -ForegroundColor green
            
    }
    catch {
        Write-Host "failed: $($_.exception.message)" -F Red
        Add-content $Logfile -value " ... failed: $($_.exception.message)"
	}
} else {
    If($MemberOfGroup) {
        ## Load users from On-Prem Active Directory using a specific User Group
		try{
			$users = Get-ADGroupMember $MemberOfGroup | Get-ADUser -Properties * | Select UserPrincipalName, thumbnailPhoto, DisplayName | Where-Object { $_.thumbnailPhoto -ne $NULL }
			write-host "done" -ForegroundColor green
		}
		catch{
		    Write-Host "failed: $($_.exception.message)" -F Red
			Add-content $Logfile -value " ... failed: $($_.exception.message)"
		}
    } else {
        ## Load users from On-Prem Active Directory
		try{ 
			$users = Get-ADUser -filter $SearchFilterAD -Properties * -SearchBase $SearchBase | Select UserPrincipalName, thumbnailPhoto, DisplayName | Where-Object { $_.thumbnailPhoto -ne $NULL }  
			write-host "done" -ForegroundColor green
		}
		catch{
			Write-Host " ... failed: $($_.exception.message)" -F Red
			Add-content $Logfile -value " ... failed: $($_.exception.message)"		
		}
    }

}

## Check for and create temp photo storage folder
if(!(Test-Path $tempPhotoStorage)) {
    $newTempPhotoStorage = New-Item -Path $tempPhotoStorage -ItemType "directory"
}

## Export each users AD Pictures
Add-content $Logfile -value "Exporting User photos from '$photoSource' ..." 
Write-Host "Exporting User photos from '$photoSource' ...`n"  -ForegroundColor Cyan
foreach ($user in $users) {

    ## Reset Photo properties
    $photo = $null
    $photoInfo = $null
    $FileName = $null
    $thumbnail = $null
    $aadPhotoExport = $null
    $photoType = $null

    ## Save the User's UPN
    $upn = $user.UserPrincipalName

    ## Load Photos based on PhotoSource
    Add-content $Logfile -value "Checking '$photoSource' for user photo for '$upn'"
    Write-Host "Checking '$photoSource' for user photo for '$upn'" -ForegroundColor Cyan
    Switch($photoSource) {
		"AD" {
			$adPhoto = $user.thumbnailPhoto
			$FileName = "$tempPhotoStorage" + $upn + ".jpg"

            [System.Io.File]::WriteAllBytes($Filename, $adPhoto)

            $photo = Get-ChildItem -Path $FileName
		}
		"AAD" {

            if($dirSource -eq "AD") {
                $user = Get-AzureADUser -ObjectId $upn
            }

			$thumbnail = $user.ExtensionProperty."thumbnailPhoto@odata.mediaEditLink"
            $photoType = ($user.ExtensionProperty."thumbnailPhoto@odata.mediaContentType")

			if($null -ne $thumbnail -and $null -ne $photoType) {
				
                try{ 
                    $aadPhotoExport = Get-AzureADUserThumbnailPhoto -ObjectId $user.ObjectId -FileName $upn -FilePath $tempPhotoStorage
                    
                    $aadPhotoExport = "$tempPhotoStorage" + "$upn" + ".*"
                    $photo = Get-ChildItem -Path $aadPhotoExport
                    
                    $FileName = $photo.FullName
                }
                catch {
                    Write-Host "User photo failed to export: $($_.exception.message)" -F Red
                    Add-content $Logfile -value "User photo failed to export: $($_.exception.message)"
                    
                }

			} else {
				$photo = $null
			}
		} 
		"Graph" {
            try{
                ## Select which photo size to export (48x48, 64x64, 96x96, 120x120, 240x240,360x360, 432x432, 504x504, 648x648)
                ## Users may not have all file options available to them so aim for something in the middle.
                $photoSizeId = "240x240"

                ## Check if the user has a photo in Graph
                # $photoInfo = Get-MgUserPhoto -UserId $upn -ProfilePhotoId $photoSizeId -ErrorAction SilentlyContinue
                $photoInfo = Get-MgUserPhoto -UserId $upn -ProfilePhotoId $photoSizeId
          
                if($null -ne $photoInfo) {
                    
                    ## Determine the file type of the photo
                    $photoType = $photoInfo.AdditionalProperties.Item("@odata.mediaContentType")
                    
                    switch($photoType) {
                        "image/jpeg" {
                            $fileExt = "jpg" 
                        }
                        "image/png" {
                            $fileExt = "png" 
                        }
                        "image/tiff" {
                            $fileExt = "tiff" 
                        }
                        "image/bmp" {
                            $fileExt = "bmp" 
                        }
                        "image/gif" {
                            $fileExt = "gif" 
                        }
                    }

                    ## Create Photo export name + path
                    $FileName = $tempPhotoStorage + $upn + "." + $fileExt 
                    
                    ## Export file to temp stroage
                    try{
                        $photoExport = Get-MgUserPhotoContent -UserId $upn -ProfilePhotoId $photoSizeId -OutFile $FileName 
                        $photo = Get-ChildItem -Path $FileName
                    }
                    catch {
                        Write-Host "User photo failed to export: $($_.exception.message)" -F Red
                        Add-content $Logfile -value "User photo failed to export: $($_.exception.message)"
                        
                    }
                }
            }
            catch {
                Write-Host "User Photo Not found: $($_.exception.message)" -F Red
                Add-content $Logfile -value "User Photo Not found: $($_.exception.message)"
                
            }
		}
	}
	
    If($null -ne $photo){
        Add-content $Logfile -value "Photo found for '$upn'"
        Write-Host "Photo found for '$upn'" -ForegroundColor Green
        if($generatePhotos -eq $true) {

            # Create and Resize Photos
            Add-content $Logfile -value "Processing '$upn'"
            Write-Host "Processing '$upn'" -ForegroundColor Green

			Add-content $Logfile -value "Exporting AD Picture to '$FileName'"
            Write-Host "Exporting '$dirSource' Picture to '$FileName'" -ForegroundColor Yellow

            Add-content $Logfile -value "Creating SharePoint Photo Folder '$($tempPhotoStorage)$($upn)'" 
            Write-Host "Creating SharePoint Photo Folder '$($tempPhotoStorage)$($upn)'" -ForegroundColor Yellow

            New-Item -Path "$tempPhotoStorage" -Name $upn -ItemType "directory" -Force | Out-Null
            $FormattedName = $upn -replace "\.", "_" -replace "@", "_"

            Add-content $Logfile -value "Resizing Image for SharePoint"
            Write-Host "Resizing Image for SharePoint" -ForegroundColor Yellow

            $NewFilePath = "$tempPhotoStorage" + $upn + "\" + $FormattedName + "_SThumb.jpg"
            Resize-Image -InputFile $FileName -Width 48 -Height 48 -Outputfile $NewFilePath

            Add-content $Logfile -value " - Small Thumbnail Created"
            Write-Host " - Small Thumbnail Created" -ForegroundColor Cyan

            $NewFilePath = "$tempPhotoStorage" + $upn + "\" + $FormattedName + "_MThumb.jpg"
            Resize-Image -InputFile $FileName -Width 72 -Height 72 -Outputfile $NewFilePath

            Add-content $Logfile -value " - Medium Thumbnail Created"
            Write-Host " - Medium Thumbnail Created" -ForegroundColor Cyan

            $NewFilePath = "$tempPhotoStorage" + $upn + "\" + $FormattedName + "_LThumb.jpg"
            Resize-Image -InputFile $FileName -Width 96 -Height 96 -Outputfile $NewFilePath

            Add-content $Logfile -value " - Large Thumbnail Created"
            Write-Host " - Large Thumbnail Created" -ForegroundColor Cyan
            
        }

        # Set UserPhoto to set O365 Photos
        # Write-Host "Processing '$upn'" -ForegroundColor Green
        if($setMSGraphPhoto -eq $true) {
            Add-content $Logfile -value "Setting 0365 User Photo for '$upn' ..."
            Write-Host "Setting 0365 User Photo for $upn ..." -ForegroundColor yellow -NoNewline
            
            try {
                              
                $newMgUserPhoto = Set-MgUserPhotoContent -UserId $upn -InFile $FileName -ErrorVariable $MSGraphPhotoError
                write-host " done. " -ForegroundColor green
            
            }
            catch {
                # Add-content $Logfile -value " ... failed: $($_.exception.message)"
                Add-content $Logfile -value " ... failed: <error below>"
                Add-content $Logfile -value $MSGraphPhotoError

                Write-Host " failed: $($_.exception.message)" -F Red
            }
             
		}
        # Upload file to SPO
        if($setSPOPhoto -eq $true) {
            Add-content $Logfile -value "Uploading pictures to SharePoint"
            Write-Host "Uploading pictures to SharePoint" -ForegroundColor Yellow
			
            $Folder = "$tempPhotoStorage" + $upn
			
            Add-content $Logfile -value " Source Picture Path is '$Folder'"
            Write-Host " Source Picture Path is '$Folder'" -ForegroundColor Yellow

            Foreach ($File in (Get-ChildItem $Folder)) {  
                Add-content $Logfile -value " - Uploading image: $($File.FullName)"
                Write-Host " - Uploading image:" $File.FullName -noNewLine
				try {
					$FileStream = New-Object IO.FileStream($File.FullName, [System.IO.FileMode]::Open)
					$uploadFile = Add-PnPFile -FileName $File.Name -Folder "/User Photos/Profile Pictures/" -Stream $FileStream -Connection $sptMy
					$FileStream.Close()
					Write-Host " ... done" -F green
				}
				catch {
					Write-Host " ... failed: $($_.exception.message)" -F Red
					Add-content $Logfile -value " ... failed: $($_.exception.message)"
				}
            }

            # Set User Photo Url in SharePoint UPS
            $FormattedName = $upn -replace "\.", "_" -replace "@", "_"
			
            Add-content $Logfile -value "SharePoint Picture Upload Complete"
            Write-Host "SharePoint Picture Upload Complete" -ForegroundColor yellow
			
            Add-content $Logfile -value "Setting SharePoint User Photo" 
            Write-Host "Setting SharePoint User Photo" -ForegroundColor yellow -noNewLine
			try {
				Set-PnPUserProfileProperty -Account $upn -PropertyName "PictureURL" -Value "$($mySiteUrl)/User%20Photos/Profile%20Pictures/$($FormattedName)_LThumb.jpg" -Connection $sptAdmin -ErrorAction Stop
				
				Write-Host " ... done" -F green
				Add-content $Logfile -value " ... done"
			}
			catch {
				Write-Host " ... failed: $($_.exception.message)" -F Red
				Add-content $Logfile -value " ... failed: $($_.exception.message)"
			}
			
			
            # Clean up phots after they've been uploaded and if option is enabled.
            if($cleanUpPhotos -eq $true){
                Add-content $Logfile -value "Removing Temp images" 
                Write-Host "Removing Temp images ... " -ForegroundColor yellow -noNewLine
				
				## Remove temp user exports
				try {
					
					Remove-Item $tempPhotoStorage\$upn -Force -Recurse
					Remove-Item $FileName
					
					write-host "done" -F Green
				}
				catch {
					Write-Host " ... failed: $($_.exception.message)" -F Red
					Add-content $Logfile -value " ... failed: $($_.exception.message)"
				}					
            }
        }
        Add-content $Logfile -value "Photo Sync for '$upn' Complete"
        Write-Host "Photo Sync for '$upn' Complete`n" -ForegroundColor Cyan

        ## Reset Photo Variable to be safe.
        $photo = $null
        
    } else{
        Add-content $Logfile -value " - No Photo found for '$upn'"
        Write-Host " - No Photo found for '$upn'" -ForegroundColor Yellow
    }

    
}

Write-Host "Done, Exiting ..." -F Cyan

# Clean up Sessions
if($setMSGraphPhoto -eq $true) {
   $disconnectGraph = Disconnect-MgGraph
}
if($dirSource -eq "AAD") {
    $disconnectAzureAD = Disconnect-AzureAD
}

$sptAdmin = $null
$sptMy = $null

# Disconnect-PnPOnline

Add-content $Logfile -value "----------------------  Photo Sync Script Script Completed: $(Get-date -format 'dd/MM/yyy hh:mm:ss tt') -----------------"