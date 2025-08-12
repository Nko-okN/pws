# Requires: Install-Module PnP.PowerShell

param(
    [Parameter(Mandatory = $true)]
    [string] $TenantId,
    
    [Parameter(Mandatory = $true)]
    [string] $ClientId,

    [Parameter(Mandatory = $true)]
    [string] $Thumbprint,

    [Parameter(Mandatory = $true)]
    [string[]] $TaskIds # Example: @("GUID-1", "GUID-2")
)

$ErrorActionPreference = "Stop"

# Build SPO Admin URL
$tenantName = ($TenantId -split '\.')[0]
$adminUrl = "https://$tenantName-admin.sharepoint.com"

# Connect to SharePoint Online Admin with Certificate Auth
Connect-PnPOnline -Url $adminUrl -ClientId $ClientId -Thumbprint $Thumbprint -Tenant $TenantId

# Endpoint pattern: /_api/migrationmanagerservice/v1.0/tasks('{TaskId}')
function Get-MigrationTaskById($id) {
    $endpoint = "/_api/migrationmanagerservice/v1.0/tasks('$id')"
    try {
        $task = Invoke-PnPSPRestMethod -Url $endpoint -Method Get
        return $task
    } catch {
        Write-Warning "Failed to fetch task ID $id: $_"
        return $null
    }
}

# Gather info
$results = @()

foreach ($id in $TaskIds) {
    $task = Get-MigrationTaskById -id $id
    if ($task -ne $null) {
        $results += [pscustomobject]@{
            DisplayName       = $task.displayName
            TaskId            = $task.id
            Status            = $task.status
            TotalKB           = [math]::Round($task.totalBytes / 1KB, 2)
            TransferredKB     = [math]::Round($task.transferredBytes / 1KB, 2)
            FileCount         = $task.fileCount
            StartTime         = $task.startTime
            EndTime           = $task.endTime
        }
    }
}

# Output the table
$results | Format-Table -AutoSize

# Optional: Export to CSV
# $results | Export-Csv -Path ".\MigrationResults.csv" -NoTypeInformation -Encoding UTF8
