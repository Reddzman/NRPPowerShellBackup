<# Cat.PnP-ConvertHtmlToModernPage.ps1

This script will take html files from an input folder, clean and slightly adjust the HTML, and write it to a new Modern Page.

Dependencies
	-	PnP PowerShell https://msdn.microsoft.com/en-us/pnp_powershell/pnp-powershell-overview

Command
	.\Cat.PnP-ConvertHtmlToModernPage.ps1

CSV Format
    Script expects a csv with 3 columns:
    1. Source (The name of the HTML file in the input folder)
    2. Title (The title of the modern page in SharePoint)
    3. NewFile (The file name for the new modern page in SharePoint)

Example:
Source          Title       NewFile
-------         ------      ----------   
pagexyz.html    Page XYZ    page-xyz
#>

$csv = "C:\temp\htmlInput\filesToConvert.csv" # CSV containing the name and title of the html file
$site = "https://rdicicco.sharepoint.com/sites/guestsite"
$inFolder = "C:\temp\htmlInput\" # location of the html files

function CreatePageAddHtml($srcFile,$newFile,$title) { 
    $path = $inFolder + $srcFile      
    $content = Get-Content -Path "$path"  -Raw		
    
    #Table wrappers (otherwise SharePoint will strip table tags):
    $html = $content -replace '<TABLE','<div class="canvasRteResponsiveTable"><div class="tableWrapper"><TABLE'
    $html = $html -replace '</TABLE>','</TABLE></div></div>'	
    
    #Backwards replace of all the headings (so stupid, but SharePoint uses html H2 for what they call Heading 1 in the text editor)
    $html = $html -replace '<h4','<h5'
    $html = $html -replace '<h3','<h4'
    $html = $html -replace '<h2','<h3'
    $html = $html -replace '<h1','<h2'
    #Backwards replace of all the headings (close tags)
    $html = $html -replace '</h4','</h5'
    $html = $html -replace '</h3','</h4'
    $html = $html -replace '</h2','</h3'
    $html = $html -replace '</h1','</h2'

    # Add Page - may need to adjust header layout and comments settings if needed.  Reference:
    # https://docs.microsoft.com/en-us/powershell/module/sharepoint-pnp/add-pnpclientsidepage?view=sharepoint-ps
    $pg = Add-PnPClientSidePage -Name "$newFile" -HeaderLayoutType NoImage -CommentsEnabled:$false
    
    # Add web parts   
    Add-PnPClientSideText -Page $pg -Text $html   

    # Set Page Title
    $set = Set-PnPClientSidePage -Identity $pg -Title "$title" -Publish
}

#MAIN SCRIPT START
$table = Import-Csv $csv

Connect-PnPOnline -Url $site

foreach ($row in $table) {
    $srcFile = $row.Source
    $title = $row.Title
    $newFile = $row.NewFile
    Write-Host "Creating page for $title"    
    CreatePageAddHtml $srcFile $newFile $title
}