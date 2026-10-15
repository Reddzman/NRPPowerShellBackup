[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory)]
    [string]$Path
)

$root = Get-Item -LiteralPath $Path -Force -ErrorAction Stop
$items = @($root) + @(Get-ChildItem -LiteralPath $Path -Recurse -Force -ErrorAction Continue)

foreach ($item in $items) {
    if ($PSCmdlet.ShouldProcess($item.FullName, 'Clear read-only attribute')) {
        try {
            if ($item.IsReadOnly) {
                $item.IsReadOnly = $false
            }
        }
        catch {
            Write-Warning "Could not clear read-only on '$($item.FullName)': $($_.Exception.Message)"
        }
    }

    if (-not $item.PSIsContainer -and
        $PSCmdlet.ShouldProcess($item.FullName, 'Unblock file')) {
        try {
            Unblock-File -LiteralPath $item.FullName -ErrorAction Stop
        }
        catch {
            Write-Warning "Could not unblock '$($item.FullName)': $($_.Exception.Message)"
        }
    }
}