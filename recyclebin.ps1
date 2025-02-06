# Import the PnP PowerShell module in the parent script
Import-Module PnP.PowerShell

# Define paths
$outputCsvPath = "C:\Path\To\Output\RecycleBinReport.csv"
$logFolderPath = "C:\Path\To\Output\Logs"

# Create the log folder if it doesn't exist
if (-not (Test-Path $logFolderPath)) {
    New-Item -ItemType Directory -Path $logFolderPath | Out-Null
}

# Initialize an array to store the results
$results = @()

# Define Client ID and Thumbprint for PnP connection
$clientId = "your-client-id"
$thumbprint = "your-thumbprint"

# Connect to the SharePoint tenant (parent script)
$adminSiteUrl = "https://<your-tenant>-admin.sharepoint.com"
Connect-PnPOnline -Url $adminSiteUrl -ClientId $clientId -Thumbprint $thumbprint

# Get all SharePoint sites in the tenant
$sites = Get-PnPTenantSite

# Define batch size (e.g., 1000 sites per batch)
$batchSize = 1000
$totalSites = $sites.Count
$batches = [math]::Ceiling($totalSites / $batchSize)

# Log start of processing
"Processing started at $(Get-Date). Total sites: $totalSites, Batch size: $batchSize, Total batches: $batches" | Out-File -FilePath "$logFolderPath\OverallLog.txt" -Append

# Function to get all recycle bin items with pagination
function Get-AllRecycleBinItems {
    param ($stage)

    $allItems = @()
    $rowLimit = 5000
    do {
        $currentItems = Get-PnPRecycleBinItem -Stage $stage -RowLimit $rowLimit
        $allItems += $currentItems
    } while ($currentItems.Count -eq $rowLimit)

    return $allItems
}

# Function to process a single site
function Process-Site {
    param ($site, $batchNumber, $siteIndex, $totalSites, $logFilePath, $clientId, $thumbprint)

    # Initialize a result object
    $result = [PSCustomObject]@{
        SiteUrl                              = $site.Url
        StorageUsedGB                        = "Error"
        StorageCapacityGB                    = "Error"
        FirstStageTotalItems                 = "Error"
        FirstStageSystemItemsCount           = "Error"
        FirstStageSystemItemsSizeGB          = "Error"
        SecondStageTotalItems                = "Error"
        SecondStageSystemItemsCount          = "Error"
        SecondStageSystemItemsSizeGB         = "Error"
        TotalItemsBothStages                 = "Error"
        TotalSystemItemsBothStages           = "Error"
        TotalSystemItemsSizeBothStagesGB     = "Error"
        TotalRecycleBinSizeGB                = "Error"
    }

    try {
        # Log progress
        $logMessage = "Batch $batchNumber, Site $siteIndex of $totalSites, URL: $($site.Url)"
        $logMessage | Out-File -FilePath $logFilePath -Append

        # Connect to the site using Client ID and Thumbprint
        Connect-PnPOnline -Url $site.Url -ClientId $clientId -Thumbprint $thumbprint -ErrorAction Stop

        # Get site storage information
        $siteStorage = Get-PnPSite -Includes StorageUsage, StorageQuota
        $storageUsed = $siteStorage.StorageUsage
        $storageCapacity = $siteStorage.StorageQuota

        # Get all first-stage recycle bin items with pagination
        $firstStageRecycleBin = Get-AllRecycleBinItems -stage "FirstStage"
        $firstStageItemsDeletedBySystem = $firstStageRecycleBin | Where-Object { $_.DeletedByEmail -eq "System Account" }
        $firstStageTotalItems = $firstStageRecycleBin.Count
        $firstStageSystemItemsCount = $firstStageItemsDeletedBySystem.Count
        $firstStageSystemItemsSize = ($firstStageItemsDeletedBySystem | Measure-Object -Property Size -Sum).Sum

        # Get all second-stage recycle bin items with pagination
        $secondStageRecycleBin = Get-AllRecycleBinItems -stage "SecondStage"
        $secondStageItemsDeletedBySystem = $secondStageRecycleBin | Where-Object { $_.DeletedByEmail -eq "System Account" }
        $secondStageTotalItems = $secondStageRecycleBin.Count
        $secondStageSystemItemsCount = $secondStageItemsDeletedBySystem.Count
        $secondStageSystemItemsSize = ($secondStageItemsDeletedBySystem | Measure-Object -Property Size -Sum).Sum

        # Calculate combined totals
        $totalItemsBothStages = $firstStageTotalItems + $secondStageTotalItems
        $totalSystemItemsBothStages = $firstStageSystemItemsCount + $secondStageSystemItemsCount
        $totalSystemItemsSizeBothStages = $firstStageSystemItemsSize + $secondStageSystemItemsSize
        $totalRecycleBinSize = ($firstStageRecycleBin | Measure-Object -Property Size -Sum).Sum + ($secondStageRecycleBin | Measure-Object -Property Size -Sum).Sum

        # Update the result object
        $result.StorageUsedGB = [math]::Round($storageUsed / 1GB, 2)
        $result.StorageCapacityGB = [math]::Round($storageCapacity / 1GB, 2)
        $result.FirstStageTotalItems = $firstStageTotalItems
        $result.FirstStageSystemItemsCount = $firstStageSystemItemsCount
        $result.FirstStageSystemItemsSizeGB = [math]::Round($firstStageSystemItemsSize / 1GB, 2)
        $result.SecondStageTotalItems = $secondStageTotalItems
        $result.SecondStageSystemItemsCount = $secondStageSystemItemsCount
        $result.SecondStageSystemItemsSizeGB = [math]::Round($secondStageSystemItemsSize / 1GB, 2)
        $result.TotalItemsBothStages = $totalItemsBothStages
        $result.TotalSystemItemsBothStages = $totalSystemItemsBothStages
        $result.TotalSystemItemsSizeBothStagesGB = [math]::Round($totalSystemItemsSizeBothStages / 1GB, 2)
        $result.TotalRecycleBinSizeGB = [math]::Round($totalRecycleBinSize / 1GB, 2)
    }
    catch {
        # Log errors
        $errorMessage = "Error processing site $($site.Url): $_"
        $errorMessage | Out-File -FilePath $logFilePath -Append
    }

    # Return the result object
    return $result
}

# Process all batches in parallel
$batchJobs = @()
for ($batchNumber = 1; $batchNumber -le $batches; $batchNumber++) {
    $batchStart = ($batchNumber - 1) * $batchSize
    $batchEnd = [math]::Min($batchStart + $batchSize - 1, $totalSites - 1)
    $batchSites = $sites[$batchStart..$batchEnd]

    # Create a log file for this batch
    $batchLogFilePath = "$logFolderPath\Batch_$batchNumber.log"
    "Starting batch $batchNumber (sites $batchStart to $batchEnd)" | Out-File -FilePath $batchLogFilePath -Append

    # Process sites in parallel using jobs
    $jobs = @()
    foreach ($site in $batchSites) {
        $siteIndex = $batchSites.IndexOf($site) + $batchStart + 1
        $job = Start-Job -ScriptBlock {
            # Import the PnP module inside the job
            Import-Module PnP.PowerShell

            # Define the functions inside the job
            function Get-AllRecycleBinItems {
                param ($stage)

                $allItems = @()
                $rowLimit = 5000
                do {
                    $currentItems = Get-PnPRecycleBinItem -Stage $stage -RowLimit $rowLimit
                    $allItems += $currentItems
                } while ($currentItems.Count -eq $rowLimit)

                return $allItems
            }

            function Process-Site {
                param ($site, $batchNumber, $siteIndex, $totalSites, $logFilePath, $clientId, $thumbprint)

                # Initialize a result object
                $result = [PSCustomObject]@{
                    SiteUrl                              = $site.Url
                    StorageUsedGB                        = "Error"
                    StorageCapacityGB                    = "Error"
                    FirstStageTotalItems                 = "Error"
                    FirstStageSystemItemsCount           = "Error"
                    FirstStageSystemItemsSizeGB          = "Error"
                    SecondStageTotalItems                = "Error"
                    SecondStageSystemItemsCount          = "Error"
                    SecondStageSystemItemsSizeGB         = "Error"
                    TotalItemsBothStages                 = "Error"
                    TotalSystemItemsBothStages           = "Error"
                    TotalSystemItemsSizeBothStagesGB     = "Error"
                    TotalRecycleBinSizeGB                = "Error"
                }

                try {
                    # Log progress
                    $logMessage = "Batch $batchNumber, Site $siteIndex of $totalSites, URL: $($site.Url)"
                    $logMessage | Out-File -FilePath $logFilePath -Append

                    # Connect to the site using Client ID and Thumbprint
                    Connect-PnPOnline -Url $site.Url -ClientId $clientId -Thumbprint $thumbprint -ErrorAction Stop

                    # Get site storage information
                    $siteStorage = Get-PnPSite -Includes StorageUsage, StorageQuota
                    $storageUsed = $siteStorage.StorageUsage
                    $storageCapacity = $siteStorage.StorageQuota

                    # Get all first-stage recycle bin items with pagination
                    $firstStageRecycleBin = Get-AllRecycleBinItems -stage "FirstStage"
                    $firstStageItemsDeletedBySystem = $firstStageRecycleBin | Where-Object { $_.DeletedByEmail -eq "System Account" }
                    $firstStageTotalItems = $firstStageRecycleBin.Count
                    $firstStageSystemItemsCount = $firstStageItemsDeletedBySystem.Count
                    $firstStageSystemItemsSize = ($firstStageItemsDeletedBySystem | Measure-Object -Property Size -Sum).Sum

                    # Get all second-stage recycle bin items with pagination
                    $secondStageRecycleBin = Get-AllRecycleBinItems -stage "SecondStage"
                    $secondStageItemsDeletedBySystem = $secondStageRecycleBin | Where-Object { $_.DeletedByEmail -eq "System Account" }
                    $secondStageTotalItems = $secondStageRecycleBin.Count
                    $secondStageSystemItemsCount = $secondStageItemsDeletedBySystem.Count
                    $secondStageSystemItemsSize = ($secondStageItemsDeletedBySystem | Measure-Object -Property Size -Sum).Sum

                    # Calculate combined totals
                    $totalItemsBothStages = $firstStageTotalItems + $secondStageTotalItems
                    $totalSystemItemsBothStages = $firstStageSystemItemsCount + $secondStageSystemItemsCount
                    $totalSystemItemsSizeBothStages = $firstStageSystemItemsSize + $secondStageSystemItemsSize
                    $totalRecycleBinSize = ($firstStageRecycleBin | Measure-Object -Property Size -Sum).Sum + ($secondStageRecycleBin | Measure-Object -Property Size -Sum).Sum

                    # Update the result object
                    $result.StorageUsedGB = [math]::Round($storageUsed / 1GB, 2)
                    $result.StorageCapacityGB = [math]::Round($storageCapacity / 1GB, 2)
                    $result.FirstStageTotalItems = $firstStageTotalItems
                    $result.FirstStageSystemItemsCount = $firstStageSystemItemsCount
                    $result.FirstStageSystemItemsSizeGB = [math]::Round($firstStageSystemItemsSize / 1GB, 2)
                    $result.SecondStageTotalItems = $secondStageTotalItems
                    $result.SecondStageSystemItemsCount = $secondStageSystemItemsCount
                    $result.SecondStageSystemItemsSizeGB = [math]::Round($secondStageSystemItemsSize / 1GB, 2)
                    $result.TotalItemsBothStages = $totalItemsBothStages
                    $result.TotalSystemItemsBothStages = $totalSystemItemsBothStages
                    $result.TotalSystemItemsSizeBothStagesGB = [math]::Round($totalSystemItemsSizeBothStages / 1GB, 2)
                    $result.TotalRecycleBinSizeGB = [math]::Round($totalRecycleBinSize / 1GB, 2)
                }
                catch {
                    # Log errors
                    $errorMessage = "Error processing site $($site.Url): $_"
                    $errorMessage | Out-File -FilePath $logFilePath -Append
                }

                # Return the result object
                return $result
            }

            # Call the Process-Site function inside the job
            Process-Site -site $using:site -batchNumber $using:batchNumber -siteIndex $using:siteIndex -totalSites $using:totalSites -logFilePath $using:batchLogFilePath -clientId $using:clientId -thumbprint $using:thumbprint
        }
        $jobs += $job
    }

    # Store the batch job for later
    $batchJobs += [PSCustomObject]@{
        BatchNumber = $batchNumber
        Jobs = $jobs
        LogFilePath = $batchLogFilePath
    }
}

# Wait for all batch jobs to complete
foreach ($batchJob in $batchJobs) {
    $batchJob.Jobs | Wait-Job

    # Collect results from jobs
    foreach ($job in $batchJob.Jobs) {
        $result = Receive-Job -Job $job
        if ($result) {
            $results += $result
        }
    }

    # Clean up jobs
    $batchJob.Jobs | Remove-Job

    # Log batch completion
    "Completed batch $($batchJob.BatchNumber)" | Out-File -FilePath $batchJob.LogFilePath -Append
}

# Export the results to a CSV file
$results | Export-Csv -Path $outputCsvPath -NoTypeInformation

Write-Host "Report exported to $outputCsvPath"
Write-Host "Log files saved to $logFolderPath"
