<#
.SYNOPSIS
    Applies a SharePoint Online lock state to one or more sites from a CSV.

.DESCRIPTION
    Connects to a SharePoint Online Admin Center and applies the specified
    lock state to each site listed in the input CSV.

    Supported lock states:

        Unlock
            Users have normal access.

        ReadOnly
            Users can view content but cannot modify content.

        NoAccess
            Users are blocked from accessing the site.

    The script generates an audit report showing:

        - Site URL
        - Requested lock state
        - Previous lock state
        - New lock state
        - Status
        - Timestamp
        - Error Messages

    Common use cases:

        - Migration Cutover
        - Migration Release
        - Site Archiving
        - Governance Enforcement
        - Emergency Change Freeze

.INPUTS

    CSV File

    Required Column:

        URL

    Example:

        URL
        https://tenant.sharepoint.com/sites/Finance
        https://tenant.sharepoint.com/sites/HR
        https://tenant.sharepoint.com/sites/Projects

.OUTPUTS

    CSV Report

        SiteUrl
        RequestedState
        PreviousState
        NewState
        Status
        Timestamp
        ErrorMessage

    Log File

        SiteLockState.log

.PARAMETER JobFile

    CSV containing site URLs.

.PARAMETER AdminUrl

    SharePoint Online Admin Center URL.

.PARAMETER LockState

    Lock state to apply.

    Valid values:

        ReadOnly
        Unlock
        NoAccess

.PARAMETER OutputFile

    Output audit report.

.PARAMETER RetryCount

    Number of retry attempts.

.PARAMETER RetryDelaySeconds

    Delay between retries.

.PARAMETER Force

    Suppresses confirmation prompts.

.PARAMETER WhatIfMode

    Reports actions without modifying sites.

.EXAMPLE

    # Migration cutover

    .\Q.SPO-SetSiteLock.ps1 `
        -JobFile .\Batch01.csv `
        -AdminUrl https://contoso-admin.sharepoint.com `
        -LockState ReadOnly

    Places all sites into ReadOnly mode prior to migration.

.EXAMPLE

    # Migration release

    .\Q.SPO-SetSiteLock.ps1 `
        -JobFile .\Batch01.csv `
        -AdminUrl https://contoso-admin.sharepoint.com `
        -LockState Unlock

    Restores user access after migration validation.

.EXAMPLE

    # Archive sites

    .\Q.SPO-SetSiteLock.ps1 `
        -JobFile .\ArchiveSites.csv `
        -AdminUrl https://contoso-admin.sharepoint.com `
        -LockState NoAccess

    Prevents user access to archived sites.

.EXAMPLE

    # Dry run

    .\Q.SPO-SetSiteLock.ps1 `
        -JobFile .\Batch01.csv `
        -AdminUrl https://contoso-admin.sharepoint.com `
        -LockState ReadOnly `
        -WhatIfMode

    Displays the actions that would be performed without making changes.

.EXAMPLE

    # Marshfield migration wave

    .\Q.SPO-SetSiteLock.ps1 `
        -JobFile .\MFLD_Batch03.csv `
        -AdminUrl https://mfldclin-admin.sharepoint.com `
        -LockState ReadOnly

    Sets migration wave sites to ReadOnly prior to migration processing.

.SCENARIOS

    Migration Cutover

        Set sites to ReadOnly.

    Migration Completion

        Set sites to Unlock.

    Site Archival

        Set sites to NoAccess.

    Governance Compliance

        Enforce standard lock state policies.

.NOTES

    Requirements:

        Microsoft.Online.SharePoint.PowerShell

    Install:

        Install-Module Microsoft.Online.SharePoint.PowerShell `
            -Scope CurrentUser

    Permissions:

        SharePoint Administrator
        or
        Global Administrator

.AUTHOR
    Quisitive Migration Services

.VERSION
    2.0
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)]
    [ValidateNotNullOrEmpty()]
    [string]$JobFile,

    [Parameter(Mandatory=$true)]
    [ValidateNotNullOrEmpty()]
    [string]$AdminUrl,

    [Parameter(Mandatory=$true)]
    [ValidateSet("ReadOnly","Unlock","NoAccess")]
    [string]$LockState,

    [string]$OutputFile = ".\SiteLockStateReport-$((Get-Date).ToString('yyyyMMdd-HHmmss')).csv",

    [ValidateRange(1,10)]
    [int]$RetryCount = 3,

    [ValidateRange(1,120)]
    [int]$RetryDelaySeconds = 5,

    [switch]$Force,

    [switch]$WhatIfMode
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$LogFile = [System.IO.Path]::ChangeExtension($OutputFile,"log")

function Write-Log {
    param(
        [string]$Message,
        [string]$Level = "INFO"
    )

    $line = "[{0}] [{1}] {2}" -f `
        (Get-Date -Format "yyyy-MM-dd HH:mm:ss"),
        $Level,
        $Message

    Add-Content -Path $LogFile -Value $line -Encoding UTF8

    switch ($Level) {
        "ERROR" { Write-Host $line -ForegroundColor Red }
        "WARN"  { Write-Host $line -ForegroundColor Yellow }
        default { Write-Host $line }
    }
}

function Invoke-WithRetry {
    param(
        [scriptblock]$ScriptBlock,
        [string]$Operation
    )

    for ($attempt=1; $attempt -le $RetryCount; $attempt++) {
        try {
            return & $ScriptBlock
        }
        catch {
            if ($attempt -eq $RetryCount) {
                throw
            }

            Write-Log `
                -Level WARN `
                -Message "$Operation failed attempt $attempt/$RetryCount. Retrying."

            Start-Sleep -Seconds $RetryDelaySeconds
        }
    }
}

if (-not (Test-Path $JobFile)) {
    throw "Input CSV not found: $JobFile"
}

$rows = @(Import-Csv $JobFile)

if ($rows.Count -eq 0) {
    throw "Input CSV contains no rows."
}

if ($rows[0].PSObject.Properties.Name -notcontains "URL") {
    throw "Input CSV must contain a URL column."
}

Write-Log "Connecting to SharePoint Admin Center."

Import-Module Microsoft.Online.SharePoint.PowerShell -ErrorAction Stop

Connect-SPOService -Url $AdminUrl

if (-not $Force) {

    Write-Host ""
    Write-Host "WARNING!" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "You are about to apply lock state:"
    Write-Host "   $LockState"
    Write-Host ""
    Write-Host "To $($rows.Count) SharePoint Sites"
    Write-Host ""

    $confirmation = Read-Host "Continue? (Y/N)"

    if ($confirmation -ne "Y") {
        Write-Host "Cancelled."
        return
    }
}

$results = [System.Collections.Generic.List[object]]::new()

$total = $rows.Count
$counter = 0

foreach ($row in $rows) {

    $counter++

    Write-Progress `
        -Activity "Updating Site Lock States" `
        -Status "$counter of $total" `
        -PercentComplete (($counter / $total) * 100)

    $siteUrl = $row.URL

    try {

        Write-Log "Loading site: $siteUrl"

        $site = Invoke-WithRetry `
            -Operation "Get-SPOSite" `
            -ScriptBlock {
                Get-SPOSite `
                    -Identity $siteUrl `
                    -ErrorAction Stop
            }

        $previousState = $site.LockState

        if ($WhatIfMode) {

            Write-Log "WHATIF: Would change $siteUrl from $previousState to $LockState"

            $newState = $previousState

            $status = "WHATIF"
        }
        else {

            Invoke-WithRetry `
                -Operation "Set-SPOSite" `
                -ScriptBlock {
                    Set-SPOSite `
                        -Identity $siteUrl `
                        -LockState $LockState `
                        -ErrorAction Stop
                }

            Start-Sleep -Seconds 2

            $updatedSite = Invoke-WithRetry `
                -Operation "Verification" `
                -ScriptBlock {
                    Get-SPOSite `
                        -Identity $siteUrl `
                        -ErrorAction Stop
                }

            $newState = $updatedSite.LockState

            if ($newState -eq $LockState) {
                $status = "SUCCESS"
            }
            else {
                $status = "VERIFY_FAILED"
            }

            Write-Log "$siteUrl updated to $newState"
        }

        $results.Add(
            [PSCustomObject]@{
                SiteUrl         = $siteUrl
                RequestedState  = $LockState
                PreviousState   = $previousState
                NewState        = $newState
                Status          = $status
                Timestamp       = Get-Date
                ErrorMessage    = ""
            }
        )
    }
    catch {

        Write-Log `
            -Level ERROR `
            -Message "$siteUrl : $($_.Exception.Message)"

        $results.Add(
            [PSCustomObject]@{
                SiteUrl         = $siteUrl
                RequestedState  = $LockState
                PreviousState   = ""
                NewState        = ""
                Status          = "ERROR"
                Timestamp       = Get-Date
                ErrorMessage    = $_.Exception.Message
            }
        )
    }
}

Write-Progress `
    -Activity "Updating Site Lock States" `
    -Completed

$results |
    Export-Csv `
        -Path $OutputFile `
        -NoTypeInformation `
        -Encoding UTF8

$successCount = @(
    $results |
    Where-Object { $_.Status -eq "SUCCESS" }
).Count

$errorCount = @(
    $results |
    Where-Object { $_.Status -eq "ERROR" }
).Count

$verifyFailCount = @(
    $results |
    Where-Object { $_.Status -eq "VERIFY_FAILED" }
).Count

Write-Host ""
Write-Host "==============================================" -ForegroundColor Green
Write-Host "Site Lock State Processing Complete" -ForegroundColor Green
Write-Host "==============================================" -ForegroundColor Green
Write-Host "Sites Processed : $($results.Count)"
Write-Host "Successful      : $successCount"
Write-Host "Verification Failures : $verifyFailCount"
Write-Host "Errors          : $errorCount"
Write-Host "Report          : $OutputFile"
Write-Host "Log             : $LogFile"
Write-Host "==============================================" -ForegroundColor Green

Write-Log "Processing complete."
Write-Log "Report: $OutputFile"