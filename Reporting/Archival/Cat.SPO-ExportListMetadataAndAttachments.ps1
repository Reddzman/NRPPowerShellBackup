<#
Cat.SPO-ExportListMetadataAndAttachments.ps1
The purpose of this script is download all list items, the item's metadata and attachments.

-----------------------
Nick Ortiz, Catapult Systems. 
2020-05-15: Initial version.
2020-06-11: Update Destination to create list sub folders.
            Update CSV file name generation to include list name.
            Add logic to support MFA for SPO connection.
2020-06-16: Add logic to properly pull lookup column value.
2020-06-18: Add logic to allow for custom csv delimiters.

Requirements:

This script required PNPONline:

Install-Module SharePointPnPPowerShellOnline -force

Examples: 
	--Scan mode--
    .\SPO-ExportListMetadataAndAttachments.ps1 -url https://yourtenant.sharepoint.com -listName "SomeList"
    
	--Download--
	.\SPO-ExportListMetadataAndAttachments.ps1 -url https://yourtenant.sharepoint.com -listName "SomeList" -dest c:\tmp

Parameter Options:
    -url (string) : Tenant URL
    -listName (string) : List Name
	-dest (string): Folder to save the downloaded items

#>     
Param(
    [Parameter(Mandatory=$true,HelpMessage="Site URL")][string]$url,
    [Parameter(Mandatory=$true,HelpMessage="List Name")][string]$listName,
    [Parameter(Mandatory=$False,HelpMessage="Path to save files to")][string]$dest,
    [Parameter(Mandatory=$False,HelpMessage="Defaults to Comma")][string]$delimeter,
    [Parameter(Mandatory=$False,HelpMessage="connection requires multi-factor auth")][switch]$mfa
) 

# General Variables
$timeStamp = (Get-Date).ToString('yyyyMMdd-HHmm')
$outPutFileItemInventory = "SP-itemInventory_"+$listName.Replace(" ", "_")+"_"+$timeStamp+".csv"

if(!$delimeter) {
    $delimeter = ','
}

# Arrays and Variables
$itemInventory = @()
$itemCountDownloaded = 0

# Connect to SPO
if(!$mfa) {
    # Get Credentials
    $creds = (Get-Credential)
    $conn = Connect-PnPOnline -Url $url -Credentials $creds
} else {
    # Log in with MFA
    $conn = Connect-PnPOnline -Url $url -UseWebLogin
}

# Reset screen and show message
Clear-Host
Write-Host "Downloading, Please stand by ..." -ForegroundColor "Green"

# Get List Items
$items = Get-PnPListItem -List $listName -Connection $conn

# Build Object for Array
function buildTempObject ($fieldValues, $itemDetails)  {

    # Exclude system columns that aren't helpful.
    $excludedColumns = @("SyncClientId","SortBehavior","SMTotalSize",
    "SMTotalFileCount","File_x0020_Type","ComplianceAssetId","owshiddenversion",
    "FileRef","FileDirRef","ProgId","ScopeId","MetaInfo","ItemChildCount","FolderChildCount","Restricted","NoExecute",
    "ContentVersion","AccessPolicy","AppAuthor","AppEditor","SMLastModifiedDate","SMTotalFileStreamSize","Attachments",
    "InstanceID","Order","WorkflowInstanceID","WorkflowVersion","UniqueId","FSObjType","FileLeafRef","OriginatorId","GUID",
    "ContentTypeId")

    # Start Temp Object for Array
    $tempOBJ = New-Object system.object

    # Load Fields and Values into Object
    foreach($value in $fieldValues) {

        # Set field value variables.
        $valueKey = $value.Key
        $valueValue = $value.Value
        
        # Process Item Fields
        if(!($valueKey.StartsWith('_')) -and !($valueKey -in $excludedColumns)) {
                
                # Pull in CreatedBy, LastModifiedBy, and LookupField Details or just create field.
                if($valueKey -eq 'Author') {
                    $tempOBJ | add-member -type Noteproperty -name CreatedBy -value $itemDetails.FieldValues.Author.LookupValue
                    #write-host $valueKey" : "$itemDetails.FieldValues.Author.LookupValue
                } elseif($valueKey -eq 'Editor') {
                    $tempOBJ | add-member -type Noteproperty -name ModifiedBy -value $itemDetails.FieldValues.Author.LookupValue
                } elseif($valueKey -eq 'Last_x0020_Modified') {
                    $tempOBJ | add-member -type Noteproperty -name LastModified -value $itemDetails.FieldValues.Last_x0020_Modified
                } elseif($valueKey -eq 'Created_x0020_Date') {
                    $tempOBJ | add-member -type Noteproperty -name CreatedDate -value $itemDetails.FieldValues.Created_x0020_Date
                } elseif($valueValue -like "*FieldLookupValue*"){
                    $tempOBJ | add-member -type Noteproperty -name $valueKey -value $itemDetails.FieldValues.$valueKey.LookupValue
                    #write-host $valueKey" : "$itemDetails.FieldValues.$valueKey.LookupValue
                }else {
                    $tempOBJ | add-member -type Noteproperty -name $valueKey -value $valueValue
                }

        }

    }

    # Return Object
    $tempOBJ
}

# Process Items
foreach($item in $items) {

    # $itemId = $item.Id

    # Load Item details
    $itemDetails = Get-PnPListItem -List $listName -Id $item.Id

    # Load Field Values for Item
    $fieldValues = $itemDetails.FieldValues.GetEnumerator()

    # Load Attachment details
    $attachments = ForEach-Object{Get-PnPProperty -ClientObject $item -Property "AttachmentFiles"}  
    
    foreach($attachment in $attachments) {
        
        # Set Attachment Filename
        $fileName = $attachment.FileName 

        # Get built object for array.
        $tempOBJ = buildTempObject $fieldValues $itemDetails
        
        # Download files if destination has been set.
        if($dest) {
            
            # Test if destination is there, create it if not.
            if (!(Test-Path -path $dest))
            {
                New-Item $dest -type directory | Out-Null
            }
            
            # Test if listname destination is there, create if not.
            if (!(Test-Path -path $dest"\"$listName))
            {
                New-Item $dest"\"$listName -type directory | Out-Null
            }

            # Check for duplicate file names and append number
            if (Test-Path -Path $dest"\"$listName"\"$fileName)
            {
                $i = 0
                while (Test-Path -Path $dest"\"$listName"\"$fileName) {

                    $i++

                    $leaf = $fileName.split('.')[0]
                    $ext = $fileName.split('.')[-1]

                    $fileName = $leaf+"-"+$i+"."+$ext 
                    
                }
            }

            # Download the file
            Get-PnPFile -Url $attachment.ServerRelativeUrl -FileName $fileName -Path $dest"\"$listName -AsFile -Connection $conn

            # Increate downloaded item counter.
            $itemCountDownloaded++

        }

                # Execute if Item object returned with data
                if($tempOBJ) {

                    # Append File name to array object
                    $tempOBJ | add-member -type Noteproperty -name fileName -value $fileName
        
                    # Add Item to Array
                    $itemInventory += $tempobj
                }
    }

}

# Clean up our Connection
Disconnect-PnPOnline -Connection $conn

# Create CSV
$itemInventory | Export-Csv -Delimiter $delimeter -Path $dest"\"$listName"\"$outPutFileItemInventory -NoTypeInformation
Write-Host "Items Downloaded: $itemCountDownloaded" -ForegroundColor "Green"
Write-Host "CSV Saved to: $dest"\"$listName"\"$outPutFileItemInventory" -ForegroundColor "Green"