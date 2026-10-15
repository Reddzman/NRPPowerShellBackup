# Set the directory containing the CSV files
$directory = "E:\Demo250915\2509150952\permissions"

# Set the text to find and replace
#$findText = "d4c0"
#$findText = "D4C"
#$findText = "Dentistry for Children"
$findText = "Dentistry 4 Children"
#$replaceText = "client"
$replaceText = "Client"

# Get all CSV files in the directory
$csvFiles = Get-ChildItem -Path $directory -Filter *.csv

foreach ($file in $csvFiles) {
    $filePath = $file.FullName

    # Read the content of the file
    $content = Get-Content -Path $filePath

    # Replace the text
    $newContent = $content -replace $findText, $replaceText

    # Write the updated content back to the file
    Set-Content -Path $filePath -Value $newContent

    # Clearing variable Cache
    Remove-Variable content
    Remove-Variable newContent
    Remove-Variable filePath
}

Write-Host "Find and replace completed for all CSV files in $directory"
Remove-Variable csvFiles