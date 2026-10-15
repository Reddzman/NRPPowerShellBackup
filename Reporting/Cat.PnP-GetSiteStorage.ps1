# ============================================================
#ONLY WORKS With PowerShell 7
# It will prompt you to log in to your SharePoint Admin Center and then retrieve the top 20 sites by storage usage, along with their details and owners.
# It will prompt you to authenticate twice
# ============================================================

# ============================================================
# USER CONFIGURATION — Update these variables before running
# ============================================================
$tenantDomain  = "crocsconnectbra"                        # e.g. "contoso" from contoso.sharepoint.com
$exportPath    = "C:\GitHub\CatPowerShell\_Output Files\BRASites-VersionHistory.csv"       # Full path for the exported CSV file
$clientId      = "46894dbf-6715-420b-a928-018eaa0a55f7"  # ← paste your PnP Client ID from Step 1
$adminAccount  = "sr-sharegate@crocs.com"             # ← Account to temporarily add as Site Admin

# ============================================================
# DO NOT EDIT BELOW THIS LINE
# ============================================================

$adminUrl        = "https://$tenantDomain-admin.sharepoint.com"
$pnpModuleName   = "PnP.PowerShell"
$modernTemplates = @("STS#3","GROUP#0","SITEPAGEPUBLISHING#0","TEAMCHANNEL#0","TEAMCHANNEL#1")

# ============================================================
# Step 1: Check and install PnP.PowerShell
# ============================================================
if (!(Get-Module -ListAvailable -Name $pnpModuleName)) {
    Write-Host "Downloading and installing $pnpModuleName module... " -ForegroundColor Yellow
    Install-Module $pnpModuleName -Scope CurrentUser -Confirm:$false -WarningAction SilentlyContinue -Force -AllowClobber
}
Import-Module $pnpModuleName -ErrorAction Stop

# ============================================================
# Step 2: Connect to SharePoint Admin Center
# ============================================================
Write-Host "Connecting to SharePoint Admin Center..." -ForegroundColor Cyan
Connect-PnPOnline -Url $adminUrl -Interactive -ClientId $clientId -WarningAction Ignore

# ============================================================
# Step 3: Get all site URLs (bulk call — no storage sort yet)
# ============================================================
Write-Host "Retrieving all site URLs..." -ForegroundColor Cyan
$allSiteUrls = Get-PnPTenantSite | Select-Object -ExpandProperty Url #| Select-Object -First 5

# ============================================================
# Step 4: Get full detail per site (storage pulled here)
# ============================================================
Write-Host "Retrieving detailed site info — this may take a few minutes..." -ForegroundColor Cyan
$allResults = foreach ($siteUrl in $allSiteUrls) {

    Write-Host "  Reading: $siteUrl" -ForegroundColor Gray
    $s = Get-PnPTenantSite -Url $siteUrl

    # --- Site Type: Classic or Modern ---
    $siteType = if ($modernTemplates -contains $s.Template) { "Modern" } else { "Classic" }

    # --- M365 Group Connected ---
    $m365Group = if ($s.GroupId -and $s.GroupId -ne "00000000-0000-0000-0000-000000000000") { "Yes" } else { "No" }

    # --- Version History Mode ---
    $versionMode = if ($s.EnableAutoExpirationVersionTrim) {
        "Automatic"
    } elseif ($s.MajorVersionLimit -gt 0) {
        "Manual (Max: $($s.MajorVersionLimit) versions, Expire after: $($s.ExpireVersionsAfterDays) days)"
    } else {
        "Inherited / Not Set"
    }

    # --- Site Owners (Members of the Owners SharePoint Group) ---
    $owners = "Unable to retrieve"
    try {
        Write-Host "  → Granting temporary site admin to $adminAccount" -ForegroundColor DarkYellow
        Set-PnPTenantSite -Url $siteUrl -Owners $adminAccount

        Connect-PnPOnline -Url $siteUrl -Interactive -ClientId $clientId -WarningAction Ignore
        $ownersGroup = Get-PnPGroup -AssociatedOwnerGroup
        $ownerMembers = (Get-PnPGroupMember -Group $ownersGroup).Title
        $owners = $ownerMembers -join "; "

        Write-Host "  → Removing temporary site admin from $adminAccount" -ForegroundColor DarkYellow
        Remove-PnPSiteCollectionAdmin -Owners $adminAccount

        Connect-PnPOnline -Url $adminUrl -Interactive -ClientId $clientId -WarningAction Ignore
    } catch {
        Write-Warning "Could not retrieve owners for $siteUrl : $_"
        Connect-PnPOnline -Url $adminUrl -Interactive -ClientId $clientId -WarningAction Ignore
    }

    [PSCustomObject]@{
        Title                  = $s.Title
        Url                    = $s.Url
        SiteOwners             = $owners
        SiteType               = $siteType
        Template               = $s.Template
        ConnectedToTeams       = if ($s.IsTeamsConnected) { "Yes" } else { "No" }
        ConnectedToM365Group   = $m365Group
        RetentionPolicy        = "See Microsoft Purview"
        VersionHistoryMode     = $versionMode
        StorageUsedGB          = [math]::Round($s.StorageUsageCurrent / 1024, 2)   # ✅ from per-site call
        StorageQuotaGB         = [math]::Round($s.StorageQuota / 1024, 2)
    }
}

# ============================================================
# Step 5: Sort by storage and take top 25
# ============================================================
$results = $allResults | Sort-Object StorageUsedGB -Descending #| Select-Object -First 25

# ============================================================
# Step 6: Export results to CSV
# ============================================================
$results | Export-Csv -Path $exportPath -NoTypeInformation
Write-Host "✔ Report exported to $exportPath" -ForegroundColor Green