<#
.SYNOPSIS
    Exports managed disks in an Azure resource group to local VHD files.

.DESCRIPTION
    Lists the managed disks in the specified resource group, temporarily grants
    read access to each disk, and uses AzCopy to download its VHD to the selected
    local directory. Existing local files are skipped and never overwritten.
    Temporary disk access is revoked after each attempted download.

    -WhatIf performs Azure authentication and read-only disk discovery, then
    reports the planned exports without creating the destination directory,
    granting disk access, or starting AzCopy. Azure PowerShell device-code
    authentication is still performed.

    AzCopy downloads from the temporary read-only SAS URL returned by Azure, so
    no separate AzCopy sign-in is needed for this transfer.

    The disk SAS URL is a bearer credential: anyone holding it can read the disk
    while it is valid. It is kept out of console output and reports. Access is
    revoked after each attempted transfer; the configured expiry is a fallback
    if the script or machine stops before cleanup runs.

    AzCopy output is kept out of the transcript. The script reports local file
    byte progress and stops a transfer if the downloaded file size makes no
    progress for TransferTimeoutMinutes. The SAS lifetime must be long enough
    for the expected transfer; increase it for very large disks or slow links.

    Dependencies:
      - PowerShell 7.2 or later
      - Az.Accounts, Az.Compute, and Az.Resources PowerShell modules
      - AzCopy v10 available on PATH
      - Azure permissions to list disks and grant/revoke disk access
      - General functions module at ..\General\Cat.GeneralFunctions.psm1
      - AzCopy writes its own diagnostic logs; protect those logs as sensitive
        because transfer diagnostics can contain disk-access details.
      - A disk transfer requires one Azure PowerShell device-code sign-in.
        AzCopy uses the temporary SAS URL and does not require a second sign-in.

.PARAMETER ResourceGroupName
    Resource group containing the managed disks.

.PARAMETER DestinationDirectory
    Local directory for the downloaded VHD files. Relative paths are resolved
    from the directory where the script is launched.

.PARAMETER SubscriptionId
    Subscription containing the resource group. If omitted, the current Az
    context's subscription is used. The script explicitly signs in again using
    device-code authentication for this subscription.

.PARAMETER TenantId
    Tenant for device-code sign-in. If omitted, the tenant from the current Az
    context is used; provide this parameter when there is no current context.

.PARAMETER SasDurationInSeconds
    Expiration duration for each temporary disk-access grant. The default is
    24 hours. Access is revoked as soon as that disk's AzCopy transfer finishes.

.PARAMETER TransferTimeoutMinutes
    Maximum time with no increase in the local VHD file size before the transfer
    is stopped. Defaults to 30 minutes. This is an inactivity timeout, not a
    maximum total duration; active transfers may run longer.

.EXAMPLE
    CD C:\GitHub\CatPowerShell\Azure
    .\Cat.Az-ExportResourceGroupDisksToLocal.ps1 `
      -SubscriptionId "15978e7f-2653-4df3-b722-f1cc406d1218" `
      -ResourceGroupName "NickPatterson-MSDN-DevTestEnvironment" `
      -DestinationDirectory "D:\DevVMHardDrives" `
      -WhatIf

Table of Contents
    1. Script Parameters
    2. Start Date / Time
    3. Snapin(s), Modules & Environment Connection(s)
    4. Input & Output Files
    5. Variables
    6. Functions
    7. Main Script
    8. Completed Script
#>

#region 1. Script Parameters
[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
param (
    [Parameter(Mandatory)]
    [ValidateNotNullOrEmpty()]
    [string]$ResourceGroupName,

    [Parameter(Mandatory)]
    [ValidateNotNullOrEmpty()]
    [string]$DestinationDirectory,

    [string]$SubscriptionId,
    [string]$TenantId,

    [ValidateRange(60, 2147483647)]
    [int]$SasDurationInSeconds = 86400,

    [ValidateRange(1, 1440)]
    [int]$TransferTimeoutMinutes = 30
)
#endregion 1. Script Parameters

#region 2. Start Date / Time
$scriptStart = $null
$transcriptStarted = $false
$runCompleted = $false
$runError = $null
$errorCaught = $false
$originalLocation = Get-Location
$reportTimestamp = Get-Date -Format 'yyMMddHHmmss'
$resolvedDestinationDirectory = [System.IO.Path]::GetFullPath($DestinationDirectory)
$logDirectory = Join-Path $PSScriptRoot '..\_Log Files'
$reportDirectory = Join-Path $PSScriptRoot '..\_Output Files'
$transcriptPath = $null
$reportPath = Join-Path $reportDirectory "Cat.Az-ExportResourceGroupDisksToLocal-$reportTimestamp.csv"
$diskResults = [System.Collections.Generic.List[object]]::new()

try {
    Set-Location -LiteralPath $PSScriptRoot
    Get-Module Cat.* | Remove-Module -Force -WhatIf:$false
    Import-Module (Join-Path $PSScriptRoot '..\General\Cat.GeneralFunctions.psm1') -Force -ErrorAction Stop
    $scriptStart = Start-CatScript
    $transcriptPath = Join-Path $logDirectory "Cat.Az-ExportResourceGroupDisksToLocal-$($scriptStart[1]).log"
    New-Item -ItemType Directory -Path $logDirectory -Force -WhatIf:$false -ErrorAction Stop | Out-Null
    Start-Transcript -Path $transcriptPath -WhatIf:$false -ErrorAction Stop | Out-Null
    $transcriptStarted = $true

    #endregion 2. Start Date / Time

    try {
        #region 3. Snapin(s), Modules & Environment Connection(s)
        Write-Host
        Write-Host ("   {0} - Authentication and module setup" -f (Get-Date -Format 'yy-MM-dd HH:mm:ss')) -ForegroundColor Magenta
        Import-Module Az.Accounts -ErrorAction Stop
        Import-Module Az.Compute -ErrorAction Stop
        Import-Module Az.Resources -ErrorAction Stop

        $azContext = Get-AzContext -ErrorAction Stop
        $tenantToAuthenticate = $TenantId
        if ([string]::IsNullOrWhiteSpace($tenantToAuthenticate) -and $azContext -and $azContext.Tenant) {
            $tenantToAuthenticate = $azContext.Tenant.Id
        }
        if ([string]::IsNullOrWhiteSpace($tenantToAuthenticate)) {
            throw 'A tenant could not be determined. Provide -TenantId to use device-code authentication.'
        }

        $subscriptionToAuthenticate = $SubscriptionId
        if ([string]::IsNullOrWhiteSpace($subscriptionToAuthenticate) -and
            $azContext -and $azContext.Subscription) {
            $subscriptionToAuthenticate = $azContext.Subscription.Id
        }
        $connectParameters = @{
            Tenant                 = $tenantToAuthenticate
            UseDeviceAuthentication = $true
            Force                  = $true
            SkipContextPopulation  = $true
            ErrorAction            = 'Stop'
        }
        if (-not [string]::IsNullOrWhiteSpace($subscriptionToAuthenticate)) {
            $connectParameters.Subscription = $subscriptionToAuthenticate
        }
        Write-Host ("      {0} - Sign in to Azure PowerShell using the device code for tenant {1}." -f (Get-Date -Format 'yy-MM-dd HH:mm:ss'), $tenantToAuthenticate)
        Connect-AzAccount @connectParameters -WhatIf:$false | Out-Null
        $azContext = Get-AzContext -ErrorAction Stop

        if (-not $azContext -or -not $azContext.Subscription -or
            [string]::IsNullOrWhiteSpace($azContext.Subscription.Id)) {
            throw 'Device-code sign-in succeeded, but no Azure subscription is selected. Provide -SubscriptionId.'
        }
        if ($azContext.Tenant.Id -ne $tenantToAuthenticate) {
            throw "Azure returned tenant '$($azContext.Tenant.Id)' instead of the requested tenant '$tenantToAuthenticate'."
        }
        if (-not [string]::IsNullOrWhiteSpace($SubscriptionId) -and
            $azContext.Subscription.Id -ne $SubscriptionId) {
            throw "Azure returned subscription '$($azContext.Subscription.Id)' instead of '$SubscriptionId'."
        }

        $azCopyCommand = Get-Command azcopy.exe -ErrorAction SilentlyContinue
        if (-not $azCopyCommand) {
            $azCopyCommand = Get-Command azcopy -ErrorAction SilentlyContinue
        }
        if (-not $azCopyCommand) {
            throw 'AzCopy v10 was not found on PATH. Install or add AzCopy to PATH, then rerun the script.'
        }
        $azCopyPath = $azCopyCommand.Source
        if ([string]::IsNullOrWhiteSpace($azCopyPath)) {
            $azCopyPath = $azCopyCommand.Path
        }
        Get-AzResourceGroup -Name $ResourceGroupName -ErrorAction Stop | Out-Null
        #endregion 3. Snapin(s), Modules & Environment Connection(s)

        #region 4. Input & Output Files
        if (-not (Test-Path -LiteralPath $reportDirectory -PathType Container)) {
            New-Item -ItemType Directory -Path $reportDirectory -Force -WhatIf:$false -ErrorAction Stop | Out-Null
        }

        if (Test-Path -LiteralPath $resolvedDestinationDirectory) {
            if (-not (Test-Path -LiteralPath $resolvedDestinationDirectory -PathType Container)) {
                throw "The destination path exists but is not a directory: $resolvedDestinationDirectory"
            }
            $destinationDirectoryReady = $true
        }
        else {
            $destinationDirectoryReady = $false
        }
        #endregion 4. Input & Output Files

        #region 5. Variables
        Write-Host
        Write-Host 'Script Parameters' -ForegroundColor Cyan
        Write-Host ("   Resource Group: {0}" -f $ResourceGroupName)
        Write-Host ("   Subscription: {0}" -f $azContext.Subscription.Id)
        Write-Host ("   Tenant: {0}" -f $azContext.Tenant.Id)
        Write-Host ("   Destination Directory: {0}" -f $resolvedDestinationDirectory)
        Write-Host ("   SAS Duration (seconds): {0}" -f $SasDurationInSeconds)
        Write-Host ("   Transfer Inactivity Timeout (minutes): {0}" -f $TransferTimeoutMinutes)
        Write-Host ("   AzCopy: {0}" -f $azCopyPath)
        Write-Host ("   Results Report: {0}" -f $reportPath)
        #endregion 5. Variables

        #region 6. Functions
        function Invoke-AzCopyDiskDownload {
            param (
                [Parameter(Mandatory)]
                [string]$AzCopyPath,

                [Parameter(Mandatory)]
                [string]$SourceSasUrl,

                [Parameter(Mandatory)]
                [string]$DestinationPath,

                [Parameter(Mandatory)]
                [string]$DiskName,

                [long]$DiskSizeBytes = 0,

                [ValidateRange(0.01, 1440)]
                [double]$InactivityTimeoutMinutes = 30
            )

            $processStartInfo = [System.Diagnostics.ProcessStartInfo]::new()
            $processStartInfo.FileName = $AzCopyPath
            $processStartInfo.UseShellExecute = $false
            $processStartInfo.CreateNoWindow = $true
            $processStartInfo.RedirectStandardOutput = $true
            $processStartInfo.RedirectStandardError = $true
            foreach ($argument in @(
                'copy'
                $SourceSasUrl
                $DestinationPath
                '--overwrite=false'
                '--output-level=essential'
                '--log-level=ERROR'
            )) {
                $processStartInfo.ArgumentList.Add($argument)
            }

            $azCopyProcess = [System.Diagnostics.Process]::new()
            $processStarted = $false
            try {
                $azCopyProcess.StartInfo = $processStartInfo
                if (-not $azCopyProcess.Start()) {
                    throw 'AzCopy could not be started.'
                }
                $processStarted = $true

                # Drain both streams without echoing output that could contain the SAS URL.
                $standardOutputTask = $azCopyProcess.StandardOutput.ReadToEndAsync()
                $standardErrorTask = $azCopyProcess.StandardError.ReadToEndAsync()
                $lastProgressBytes = 0L
                $lastProgressTime = Get-Date
                $lastDisplayTime = [datetime]::MinValue
                $transferStartTime = Get-Date
                Write-Progress `
                    -Activity "Downloading $DiskName" `
                    -Status "Waiting for the first downloaded bytes; timeout after $InactivityTimeoutMinutes minute(s) without progress." `
                    -PercentComplete 0

                while (-not $azCopyProcess.WaitForExit(1000)) {
                    $downloadedBytes = if (Test-Path -LiteralPath $DestinationPath -PathType Leaf) {
                        (Get-Item -LiteralPath $DestinationPath).Length
                    }
                    else {
                        0L
                    }

                    if ($downloadedBytes -gt $lastProgressBytes) {
                        $lastProgressBytes = $downloadedBytes
                        $lastProgressTime = Get-Date
                    }

                    if (((Get-Date) - $lastDisplayTime).TotalSeconds -ge 15) {
                        $inactiveFor = (Get-Date) - $lastProgressTime
                        $inactiveMinutes = [int][Math]::Floor($inactiveFor.TotalMinutes)
                        if ($DiskSizeBytes -gt 0) {
                            $percentComplete = [Math]::Min(99, [int](100 * $downloadedBytes / $DiskSizeBytes))
                            $progressStatus = '{0:N2} of {1:N2} GiB; no byte progress for {2} minute(s); active for {3}' -f `
                                ($downloadedBytes / 1GB), ($DiskSizeBytes / 1GB), $inactiveMinutes, ((Get-Date) - $transferStartTime)
                        }
                        else {
                            $percentComplete = 0
                            $progressStatus = '{0:N2} GiB downloaded; no byte progress for {1} minute(s); active for {2}' -f `
                                ($downloadedBytes / 1GB), $inactiveMinutes, ((Get-Date) - $transferStartTime)
                        }

                        Write-Progress `
                            -Activity "Downloading $DiskName" `
                            -Status $progressStatus `
                            -PercentComplete $percentComplete
                        $lastDisplayTime = Get-Date
                    }

                    if (((Get-Date) - $lastProgressTime).TotalMinutes -ge $InactivityTimeoutMinutes) {
                        $azCopyProcess.Kill($true)
                        $azCopyProcess.WaitForExit()
                        throw "AzCopy was stopped after $InactivityTimeoutMinutes minutes without downloaded-byte progress."
                    }
                }

                $azCopyProcess.WaitForExit()
                $null = $standardOutputTask.GetAwaiter().GetResult()
                $null = $standardErrorTask.GetAwaiter().GetResult()
                if ($azCopyProcess.ExitCode -ne 0) {
                    throw "AzCopy failed with exit code $($azCopyProcess.ExitCode). Review the AzCopy logs for details."
                }
            }
            finally {
                Write-Progress -Activity "Downloading $DiskName" -Completed
                if ($processStarted -and -not $azCopyProcess.HasExited) {
                    $azCopyProcess.Kill($true)
                    $azCopyProcess.WaitForExit()
                }
                $azCopyProcess.Dispose()
            }
        }
        #endregion 6. Functions

        #region 7. Main Script
        Write-Host
        Write-Host ("{0} - Retrieving managed disks" -f (Get-Date -Format 'yy-MM-dd HH:mm:ss')) -ForegroundColor Cyan
        $disks = @(Get-AzDisk -ResourceGroupName $ResourceGroupName -ErrorAction Stop)
        Write-Host ("   {0} - Found {1} managed disk(s)." -f (Get-Date -Format 'yy-MM-dd HH:mm:ss'), $disks.Count)

        if (-not $destinationDirectoryReady -and $disks.Count -gt 0) {
            if ($PSCmdlet.ShouldProcess($resolvedDestinationDirectory, 'Create local export directory')) {
                New-Item -ItemType Directory -Path $resolvedDestinationDirectory -Force -ErrorAction Stop | Out-Null
                $destinationDirectoryReady = $true
            }
            elseif ($WhatIfPreference) {
                Write-Host ("   {0} - WhatIf: the destination directory would be created." -f (Get-Date -Format 'yy-MM-dd HH:mm:ss')) -ForegroundColor Yellow
            }
        }

        $destinationNames = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
        $failedCount = 0
        $diskIndex = 0
        foreach ($disk in $disks) {
            $diskIndex++
            $destinationPath = Join-Path $resolvedDestinationDirectory "$($disk.Name).vhd"
            if (-not $destinationNames.Add($destinationPath)) {
                throw "Multiple managed disks resolve to the same local file path: $destinationPath"
            }

            $result = [pscustomobject][ordered]@{
                ResourceGroupName = $ResourceGroupName
                SubscriptionId    = $azContext.Subscription.Id
                DiskName          = $disk.Name
                DiskId            = $disk.Id
                DestinationPath   = $destinationPath
                Status            = 'Pending'
                Details           = ''
            }
            $diskResults.Add($result)

            Write-Progress `
                -Activity 'Exporting managed disks' `
                -Status "$($disk.Name) ($diskIndex of $($disks.Count))" `
                -PercentComplete ([int](100 * $diskIndex / [Math]::Max($disks.Count, 1)))

            if (-not $destinationDirectoryReady) {
                $result.Status = if ($WhatIfPreference) { 'WhatIf' } else { 'Skipped' }
                $result.Details = if ($WhatIfPreference) {
                    'Would export after creating the destination directory.'
                }
                else {
                    'Destination directory was not created.'
                }
                continue
            }

            if (Test-Path -LiteralPath $destinationPath) {
                $result.Status = 'Skipped'
                $result.Details = 'Destination file already exists; existing files are never overwritten.'
                continue
            }

            if ($WhatIfPreference) {
                $result.Status = 'WhatIf'
                $result.Details = 'Would grant temporary read access and download the disk VHD.'
                Write-Host ("   {0} - WhatIf: would export {1} to {2}" -f (Get-Date -Format 'yy-MM-dd HH:mm:ss'), $disk.Name, $destinationPath) -ForegroundColor Yellow
                continue
            }

            if (-not $PSCmdlet.ShouldProcess($destinationPath, "Export managed disk '$($disk.Name)'")) {
                $result.Status = 'Skipped'
                $result.Details = 'Export was declined.'
                continue
            }

            $diskAccessGranted = $false
            $temporaryPath = $null
            try {
                Write-Host ("   {0} - Granting temporary read access for {1}..." -f (Get-Date -Format 'yy-MM-dd HH:mm:ss'), $disk.Name)
                $diskAccess = Grant-AzDiskAccess `
                    -ResourceGroupName $ResourceGroupName `
                    -DiskName $disk.Name `
                    -Access Read `
                    -DurationInSecond $SasDurationInSeconds `
                    -WhatIf:$false `
                    -ErrorAction Stop
                $diskAccessGranted = $true

                if ([string]::IsNullOrWhiteSpace($diskAccess.AccessSAS)) {
                    throw "Azure did not return a disk-access URL for '$($disk.Name)'."
                }

                $temporaryPath = "$destinationPath.$([guid]::NewGuid().ToString('N')).partial"
                Write-Host ("   {0} - Downloading VHD to {1}..." -f (Get-Date -Format 'yy-MM-dd HH:mm:ss'), $destinationPath)
                Invoke-AzCopyDiskDownload `
                    -AzCopyPath $azCopyPath `
                    -SourceSasUrl $diskAccess.AccessSAS `
                    -DestinationPath $temporaryPath `
                    -DiskName $disk.Name `
                    -DiskSizeBytes ([long]$disk.DiskSizeBytes) `
                    -InactivityTimeoutMinutes $TransferTimeoutMinutes

                if (-not (Test-Path -LiteralPath $temporaryPath -PathType Leaf)) {
                    throw 'AzCopy reported success but did not create the expected local VHD file.'
                }
                Move-Item -LiteralPath $temporaryPath -Destination $destinationPath -NoClobber -ErrorAction Stop

                $result.Status = 'Exported'
                $result.Details = 'VHD downloaded successfully.'
            }
            catch {
                $result.Status = 'Failed'
                $result.Details = $_.Exception.Message
                $failedCount++
                Write-Host ("   {0} - Export failed for {1}: {2}" -f (Get-Date -Format 'yy-MM-dd HH:mm:ss'), $disk.Name, $result.Details) -ForegroundColor Red
            }
            finally {
                if ($diskAccessGranted) {
                    try {
                        Write-Host ("      {0} - Revoking temporary disk access for {1}..." -f (Get-Date -Format 'yy-MM-dd HH:mm:ss'), $disk.Name) -ForegroundColor DarkGray
                        Revoke-AzDiskAccess `
                            -ResourceGroupName $ResourceGroupName `
                            -DiskName $disk.Name `
                            -WhatIf:$false `
                            -ErrorAction Stop | Out-Null
                    }
                    catch {
                        $revokeDetails = "Failed to revoke temporary disk access: $($_.Exception.Message)"
                        $result.Details = (@($result.Details, $revokeDetails) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }) -join ' '
                        if ($result.Status -ne 'Failed') {
                            $result.Status = 'Failed'
                            $failedCount++
                        }
                        Write-Host ("   {0} - {1}" -f (Get-Date -Format 'yy-MM-dd HH:mm:ss'), $revokeDetails) -ForegroundColor Red
                    }
                }
                if ($temporaryPath -and (Test-Path -LiteralPath $temporaryPath -PathType Leaf)) {
                    try {
                        Remove-Item -LiteralPath $temporaryPath -Force -WhatIf:$false -ErrorAction Stop
                    }
                    catch {
                        $cleanupDetails = "Failed to remove incomplete temporary download '$temporaryPath': $($_.Exception.Message)"
                        $result.Details = (@($result.Details, $cleanupDetails) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }) -join ' '
                        if ($result.Status -ne 'Failed') {
                            $result.Status = 'Failed'
                            $failedCount++
                        }
                        Write-Host ("   {0} - {1}" -f (Get-Date -Format 'yy-MM-dd HH:mm:ss'), $cleanupDetails) -ForegroundColor Red
                    }
                }
            }

            if ($result.Status -eq 'Exported') {
                Write-Host ("   {0} - Exported {1} successfully." -f (Get-Date -Format 'yy-MM-dd HH:mm:ss'), $disk.Name) -ForegroundColor Green
            }
        }
        Write-Progress -Activity 'Exporting managed disks' -Completed

        $exportedCount = @($diskResults | Where-Object Status -eq 'Exported').Count
        $skippedCount = @($diskResults | Where-Object Status -eq 'Skipped').Count
        $whatIfCount = @($diskResults | Where-Object Status -eq 'WhatIf').Count
        Write-Host
        Write-Host 'Export Summary' -ForegroundColor Cyan
        Write-Host ("   Exported: {0}; Skipped: {1}; Planned (WhatIf): {2}; Failed: {3}" -f $exportedCount, $skippedCount, $whatIfCount, $failedCount)
        Write-Host ("   Results Report: {0}" -f $reportPath)

        if ($failedCount -gt 0) {
            throw "$failedCount disk export(s) failed. Review the results report and AzCopy logs."
        }
        $runCompleted = $true
        #endregion 7. Main Script
    }
    catch {
        $runError = $_
        throw
    }
}
catch {
    if (-not $runError) {
        $runError = $_
    }
    $errorCaught = $true
    throw
}
finally {
    #region 8. Completed Script
    try {
        if ($diskResults.Count -gt 0) {
            $diskResults | Export-Csv -LiteralPath $reportPath -NoTypeInformation -Force -WhatIf:$false -ErrorAction Stop
        }
        else {
            [pscustomobject][ordered]@{
                ResourceGroupName = $null
                SubscriptionId    = $null
                DiskName          = $null
                DiskId            = $null
                DestinationPath   = $null
                Status            = $null
                Details           = $null
            } | ConvertTo-Csv -NoTypeInformation | Select-Object -First 1 | Set-Content -LiteralPath $reportPath -Encoding utf8 -WhatIf:$false -ErrorAction Stop
        }
    }
    catch {
        if ($runError) {
            Write-Error "Could not write the results report '$reportPath': $($_.Exception.Message)" -ErrorAction Continue
        }
        else {
            $runError = $_
        }
    }

    if ($scriptStart) {
        try {
            Write-Host
            Write-Host 'Script Completion / Summary' -ForegroundColor Cyan
            $completionStatus = if ($runCompleted -and -not $runError) { 'completed' } else { 'failed before completion' }
            Write-Host ("   Run status: {0}" -f $completionStatus) -ForegroundColor $(if ($completionStatus -eq 'completed') { 'Green' } else { 'Red' })
            Write-Host ("   Results Report: {0}" -f $reportPath)
            if ($transcriptPath) {
                Write-Host ("   Transcript: {0}" -f $transcriptPath)
            }
            Complete-CatScript $scriptStart[0]
        }
        catch {
            if ($runError) {
                Write-Error "Could not write the completion summary: $($_.Exception.Message)" -ErrorAction Continue
            }
            else {
                $runError = $_
            }
        }
    }

    if ($transcriptStarted) {
        try {
            Stop-Transcript -WhatIf:$false -ErrorAction Stop | Out-Null
        }
        catch {
            if ($runError) {
                Write-Error "Could not stop the transcript: $($_.Exception.Message)" -ErrorAction Continue
            }
            else {
                $runError = $_
            }
        }
    }

    try {
        Set-Location -LiteralPath $originalLocation -ErrorAction Stop
    }
    catch {
        if ($runError) {
            Write-Error "Could not restore the original working directory: $($_.Exception.Message)" -ErrorAction Continue
        }
        else {
            $runError = $_
        }
    }
    if ($runError -and -not $errorCaught) {
        $PSCmdlet.ThrowTerminatingError($runError)
    }
    #endregion 8. Completed Script
}
