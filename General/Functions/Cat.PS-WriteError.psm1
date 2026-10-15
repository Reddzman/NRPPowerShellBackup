<#  Write-CatErrorCSV
This function creates a log file for errors that happen in scripts that are iterating a CSV file.

Command: Write-CatErrorCSV -funcCSVItem $item -funcAction "Creating Site" -funcException $_.Exception -funcDetails $_.ErrorDetails -funcStackTrace $_.ScriptStackTrace -funcFilePath $filePath

Parameters
    funcCSVItem = CSV Item information
    funcAction = Action that resulted in error
    funcException = Error Exception
    funcDetails = Error Details
    funcStackTrace = Error Script Stack Trace
    funcFilePath = Path to log file #>
function Write-CatErrorCSV
{ 
	param ( 
        $funcCSVItem,
        $funcAction,
        $funcException,
        $funcDetails,
        $funcStackTrace,
        $funcFilePath
    )
    $funcSpecialCharacters = "[~``!@#$%^&*()_-]"
    Write-Host
    Write-Host ("{0} ERROR - {1}: {2}" -f (Get-Date -Format "yy-MM-dd HH:mm:ss"),$funcAction, ((($funcException -replace $funcSpecialCharacters,'') -replace "`n",' ') -replace "`r",' ')) -ForegroundColor Red
    $err = $funcCSVItem
    $err | Add-Member -MemberType NoteProperty -Name "errorAction" -Value $funcAction
    $err | Add-Member -MemberType NoteProperty -Name "errorException" -Value ((($funcException -replace $funcSpecialCharacters,'') -replace "`n",' ') -replace "`r",' ')
    $err | Add-Member -MemberType NoteProperty -Name "errorDetails" -Value ((($funcDetails -replace $funcSpecialCharacters,'') -replace "`n",' ') -replace "`r",' ')
    $err | Add-Member -MemberType NoteProperty -Name "errorStackTrace" -Value ((($funcStackTrace -replace $funcSpecialCharacters,'') -replace "`n",' ') -replace "`r",' ')
    $err | Export-Csv -Path $funcFilePath -NoTypeInformation -Append
    Remove-Variable err
    Write-Host
}