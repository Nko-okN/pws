# Requires: Install-Module PnP.PowerShell
# Run PowerShell 7+ for best results.

param(
    [Parameter(Mandatory = $true)]
    [string] $TenantName,          # e.g., contoso  (used to build https://contoso-admin.sharepoint.com)

    [Parameter(Mandatory = $true)]
    [string] $TenantId,            # Azure AD tenant GUID or domain

    [Parameter(Mandatory = $true)]
    [string] $ClientId,

    [Parameter(Mandatory = $true)]
    [string] $Thumbprint,          # Certificate thumbprint in LocalMachine or CurrentUser\My

    [string[]] $TaskIds,           # Preferred: @("guid1","guid2",...)
    [string[]] $TaskDisplayNames,  # Optional fallback: @("Wave 1","HR Docs"...)

    [string] $ExportCsvPath        # Optional: ".\MigrationVsCurrent.csv"
)

$ErrorActionPreference = "Stop"

# --- Connect to SPO Admin (app-only cert) ---
$adminUrl = "https://$TenantName-admin.sharepoint.com"
Write-Host "Connecting to $adminUrl..." -ForegroundColor Cyan
Connect-PnPOnline -Url $adminUrl -ClientId $ClientId -Thumbprint $Thumbprint -Tenant $TenantId

# --- Migration Manager endpoints (used by the Admin UI) ---
# These are the common paths; if your tenant exposes a different version path, adjust here.
$TaskBaseApi    = "/_api/migrationmanagerservice/v1.0/tasks"
$GetTaskByIdApi = "/_api/migrationmanagerservice/v1.0/tasks('{0}')"

function Invoke-AdminApi {
    param(
        [Parameter(Mandatory=$true)][string] $Url
    )
    Invoke-PnPSPRestMethod -Url $Url -Method Get
}

function Get-MmTaskById {
    param([string] $Id)
    $url = [string]::Format($GetTaskByIdApi, $Id)
    try {
        Invoke-AdminApi -Url $url
    } catch {
        Write-Warning "Task ID $Id not found or inaccessible."
        $null
    }
}

function Get-MmTasksByDisplayName {
    param([string] $DisplayName)
    # If $filter is not supported in your tenant, fetch all and filter client-side.
    try {
        $all = Invoke-AdminApi -Url $TaskBaseApi
        $items = if ($all.value) { $all.value } else { @() }
        $items | Where-Object { $_.displayName -eq $DisplayName }
    } catch {
        Write-Warning "Failed to list tasks for display name '$DisplayName'."
        @()
    }
}

function Normalize-SiteCollectionUrl {
    param([string] $Url)
    if ([string]::IsNullOrWhiteSpace($Url)) { return $null }
    # Return the site collection root (https://tenant/sites/collection or /teams/collection)
    $m = [regex]::Match($Url, '^(https://[^/]+/(sites|teams)/[^/]+)')
    if ($m.Success) { return $m.Groups[1].Value }
    # If it's already a root or a plain site, return as-is
    return $Url.TrimEnd('/')
}

function TryGet-DestinationUrl {
    param($Task)
    # Try common property names exposed by the payload
    $candidates = @(
        $Task.destinationSiteUrl,
        $Task.destinationWebUrl,
        $Task.targetSiteUrl,
        $Task.targetUrl,
        $Task.destinationUrl
    ) | Where-Object { $_ -and $_ -is [string] -and $_.Length -gt 0 }
    if ($candidates.Count -gt 0) { return $candidates[0] }
    # Some payloads include a nested destination object
    if ($Task.destination -and $Task.destination.url) { return $Task.destination.url }
    $null
}

function Select-TaskProjection {
    param($Task)
    $dest = TryGet-DestinationUrl -Task $Task
    $root = Normalize-SiteCollectionUrl -Url $dest
    [pscustomobject]@{
        TaskId             = $Task.id
        DisplayName        = $Task.displayName
        Status             = $Task.status
        TotalBytes         = $Task.totalBytes
        TransferredBytes   = $Task.transferredBytes
        FileCount          = $Task.fileCount
        StartTime          = $Task.startTime
        EndTime            = $Task.endTime
        DestinationUrl     = $dest
        SiteCollectionUrl  = $root
        TotalKB            = if ($Task.totalBytes) { [math]::Round($Task.totalBytes / 1KB, 2) } else { $null }
        TotalGB            = if ($Task.totalBytes) { [math]::Round($Task.totalBytes / 1GB, 3) } else { $null }
        TransferredKB      = if ($Task.transferredBytes) { [math]::Round($Task.transferredBytes / 1KB, 2) } else { $null }
        TransferredGB      = if ($Task.transferredBytes) { [math]::Round($Task.transferredBytes / 1GB, 3) } else { $null }
    }
}

# --- Gather tasks (by ID first; optionally by DisplayName) ---
if (-not $TaskIds -and -not $TaskDisplayNames) {
    throw "Provide at least one of: -TaskIds or -TaskDisplayNames."
}

$taskRows = New-Object System.Collections.Generic.List[object]

if ($TaskIds) {
    foreach ($id in $TaskIds) {
        $t = Get-MmTaskById -Id $id
        if ($t) { $taskRows.Add( (Select-TaskProjection -Task $t) ) | Out-Null }
        else {
            $taskRows.Add([pscustomobject]@{
                TaskId=$id; DisplayName=$null; Status="NotFound"; TotalBytes=$null; TransferredBytes=$null; FileCount=$null;
                StartTime=$null; EndTime=$null; DestinationUrl=$null; SiteCollectionUrl=$null; TotalKB=$null; TotalGB=$null; TransferredKB=$null; TransferredGB=$null
            }) | Out-Null
        }
    }
}

if ($TaskDisplayNames) {
    foreach ($name in $TaskDisplayNames) {
        $ts = Get-MmTasksByDisplayName -DisplayName $name
        if ($ts -and $ts.Count -gt 0) {
            foreach ($t in $ts) { $taskRows.Add( (Select-TaskProjection -Task $t) ) | Out-Null }
        } else {
            $taskRows.Add([pscustomobject]@{
                TaskId=$null; DisplayName=$name; Status="NotFound"; TotalBytes=$null; TransferredBytes=$null; FileCount=$null;
                StartTime=$null; EndTime=$null; DestinationUrl=$null; SiteCollectionUrl=$null; TotalKB=$null; TotalGB=$null; TransferredKB=$null; TransferredGB=$null
            }) | Out-Null
        }
    }
}

# --- Fetch current site storage (usage/quota) for each unique destination site collection ---
$siteMap = @{}
$uniqueSites = ($taskRows | ForEach-Object { $_.SiteCollectionUrl } | Where-Object { $_ } | Select-Object -Unique)

function Get-SiteStorageInfo {
    param([string] $SiteCollectionUrl)
    try {
        # Get-PnPTenantSite -Detailed exposes StorageUsage (MB) and StorageMaximumLevel (MB) in most tenants
        $sp = Get-PnPTenantSite -Url $SiteCollectionUrl -Detailed
        if (-not $sp) { return $null }
        [pscustomobject]@{
            SiteCollectionUrl = $SiteCollectionUrl
            StorageUsageMB    = [double]$sp.StorageUsage
            StorageQuotaMB    = [double]$sp.StorageMaximumLevel
            StorageUsageGB    = [math]::Round(([double]$sp.StorageUsage) / 1024, 3)
            StorageQuotaGB    = if ($sp.StorageMaximumLevel -gt 0) { [math]::Round(([double]$sp.StorageMaximumLevel) / 1024, 3) } else { $null }
        }
    } catch {
        Write-Warning "Failed to get storage info for $SiteCollectionUrl"
        $null
    }
}

foreach ($s in $uniqueSites) {
    $info = Get-SiteStorageInfo -SiteCollectionUrl $s
    if ($info) { $siteMap[$s] = $info }
}

# --- Join tasks with site storage + compute deltas ---
$joined = $taskRows | ForEach-Object {
    $siteInfo = if ($_.SiteCollectionUrl -and $siteMap.ContainsKey($_.SiteCollectionUrl)) { $siteMap[$_.SiteCollectionUrl] } else { $null }
    $currentGB = if ($siteInfo) { $siteInfo.StorageUsageGB } else { $null }
    $migratedGB = $_.TotalGB
    [pscustomobject]@{
        TaskId            = $_.TaskId
        DisplayName       = $_.DisplayName
        Status            = $_.Status
        DestinationUrl    = $_.DestinationUrl
        SiteCollectionUrl = $_.SiteCollectionUrl

        MigratedKB        = $_.TotalKB
        MigratedGB        = $_.TotalGB
        TransferredKB     = $_.TransferredKB
        TransferredGB     = $_.TransferredGB

        CurrentSiteGB     = $currentGB
        SiteQuotaGB       = if ($siteInfo) { $siteInfo.StorageQuotaGB } else { $null }
        DeltaGB           = if ($currentGB -ne $null -and $migratedGB -ne $null) { [math]::Round($currentGB - $migratedGB, 3) } else { $null }
    }
}

# --- Output ---
$joined | Sort-Object DisplayName, TaskId | Format-Table -AutoSize

if ($ExportCsvPath) {
    $joined | Export-Csv -Path $ExportCsvPath -NoTypeInformation -Encoding UTF8
    Write-Host "Exported to $ExportCsvPath" -ForegroundColor Green
}
