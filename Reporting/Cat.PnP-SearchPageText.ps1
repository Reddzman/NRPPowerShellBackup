<# uses a .csv file to crawl through all site pages libraries, find any text web parts with text containing the $filterText variable, 
and replace the text of the web part with $newText variable #>

$newText = '<span class="fontSizeSmall"><strong>PUT YOUR NEW TEXT HERE<strong></strong></span>'

$filterText = "*USE THIS TEXT TO FILTER ONLY CERTAIN TEXT WEB PARTS*"

Import-Csv -Path $PSScriptRoot\YOUREXCELFILE.csv |ForEach-Object{
    Write-Host "$($_.Url)";
    $web= $($_.Url);


    Connect-PnPOnline $web -Interactive

    $sitePagesLib = Get-PnPList -Identity "Site Pages"

    $pages = Get-PnPListItem -List $sitePagesLib -Fields FileLeafRef

    foreach($page in $pages){
        $p = Get-PnPPage -Identity $page["FileLeafRef"]
        $comps = Get-PnPPageComponent -Page $p
        foreach($comp in $comps){
            $textWp = Get-PnPPageComponent -Page $p -InstanceId $comp.InstanceId
            If($textWp.Text -like $filterText){
              Set-PnPPageTextPart -Page $p -InstanceId $textWp.InstanceId -Text $newText
            }
        }
    }
}

