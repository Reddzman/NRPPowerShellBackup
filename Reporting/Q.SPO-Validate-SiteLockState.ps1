<#
.SYNOPSIS
    Validates SharePoint Online site lock states and generates a validation report.
.DESCRIPTION
    Connects to the SharePoint Online Admin Center, reads a CSV containing site URLs,
    retrieves each site's lock state, compares it with the expected state, and
    generates a CSV validation report with PASS/FAIL results.
.INPUTS
    CSV file with required column: URL
.OUTPUTS
    CSV report containing SiteUrl, ExpectedState, CurrentState,
    ValidationStatus, Timestamp, and ErrorMsg.
.PARAMETER jobFile
    Path to CSV input file.
.PARAMETER adminURL
    SharePoint Online Admin Center URL.
.PARAMETER lockState
    Expected lock state. Valid values: ReadOnly, Unlock, NoAccess.
.PARAMETER outputFile
    Output CSV report path.
.EXAMPLE
    .\Q.SPO-Validate-SiteLockState.ps1 -jobFile .\Sites.csv -adminURL https://contoso-admin.sharepoint.com -lockState ReadOnly
    Validates that sites are configured as ReadOnly.
.EXAMPLE
    .\Q.SPO-Validate-SiteLockState.ps1 -jobFile .\Sites.csv -adminURL https://contoso-admin.sharepoint.com -lockState Unlock
    Validates that sites have been returned to normal operation after migration.
.EXAMPLE
    .\Q.SPO-Validate-SiteLockState.ps1 -jobFile .\Sites.csv -adminURL https://contoso-admin.sharepoint.com -lockState NoAccess
    Validates that archived sites are inaccessible.
.NOTES
    Requires Microsoft.Online.SharePoint.PowerShell module and SharePoint Admin permissions.
#>

[CmdletBinding()]
param(
    [string]$jobFile = 'C:\Quisitive\Powershell\_Input Files\Marshfield\Jobs\TestReadOnly.csv',
    [string]$adminURL = 'https://mfldclin-admin.sharepoint.com',
    [ValidateSet('ReadOnly','Unlock','NoAccess')]
    [string]$lockState = 'ReadOnly',
    [string]$outputFile = "C:\Quisitive\Powershell\_Output Files\SiteLockStateValidationReport-$(Get-Date -Format 'yyyyMMdd-HHmmss').csv",
    [int]$RetryCount = 3
)

$logFile = [System.IO.Path]::ChangeExtension($outputFile,'log')

function Write-Log {
    param([string]$Message,[string]$Level='INFO')
    $line = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') [$Level] $Message"
    Add-Content -Path $logFile -Value $line -Encoding UTF8
    Write-Host $line
}

Import-Module Microsoft.Online.SharePoint.PowerShell -ErrorAction Stop

Write-Log "Connecting to SharePoint Admin Center: $adminURL"
Connect-SPOService -Url $adminURL

$rows = Import-Csv $jobFile
$total = $rows.Count
$counter = 0

$results = foreach ($row in $rows) {
    $counter++
    Write-Progress -Activity 'Validating Site Lock States' -Status "$counter of $total" -PercentComplete (($counter / $total) * 100)

    try {
        $site = $null
        for ($i = 1; $i -le $RetryCount; $i++) {
            try {
                $site = Get-SPOSite -Identity $row.URL -ErrorAction Stop
                break
            }
            catch {
                if ($i -eq $RetryCount) { throw }
                Start-Sleep -Seconds 5
            }
        }

        $status = if ($site.LockState -eq $lockState) { 'PASS' } else { 'FAIL' }

        [PSCustomObject]@{
            SiteUrl          = $site.Url
            ExpectedState    = $lockState
            CurrentState     = $site.LockState
            ValidationStatus = $status
            Timestamp        = Get-Date
            ErrorMsg         = ''
        }
    }
    catch {
        [PSCustomObject]@{
            SiteUrl          = $row.URL
            ExpectedState    = $lockState
            CurrentState     = 'ERROR'
            ValidationStatus = 'ERROR'
            Timestamp        = Get-Date
            ErrorMsg         = $_.Exception.Message
        }
    }
}

$results | Export-Csv -Path $outputFile -NoTypeInformation -Encoding UTF8

$pass = @($results | Where-Object ValidationStatus -eq 'PASS').Count
$fail = @($results | Where-Object ValidationStatus -eq 'FAIL').Count
$error = @($results | Where-Object ValidationStatus -eq 'ERROR').Count

Write-Log "Sites Scanned: $($results.Count)"
Write-Log "PASS: $pass"
Write-Log "FAIL: $fail"
Write-Log "ERROR: $error"
Write-Log "Report Generated: $outputFile"
