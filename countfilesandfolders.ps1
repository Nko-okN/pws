# Import necessary modules
Import-Module PnP.PowerShell

# SharePoint site and folder details
$siteUrl = "https://your-sharepoint-site-url"
$folderRelativeUrl = "/sites/your-site/Shared Documents/YourFolder" # Replace with your target folder path

# Connect to SharePoint site
Write-Host "Connecting to SharePoint site: $siteUrl" -ForegroundColor Cyan
Connect-PnPOnline -Url $siteUrl -UseWebLogin
Write-Host "Successfully connected to SharePoint site." -ForegroundColor Green

# Initialize counters
$totalFolders = 0
$totalFiles = 0

# Initialize a list to store folder and file details
$treeSizeReport = @()

# Function to recursively process folders and files
function Get-FolderTree {
    param (
        [string]$folderRelativeUrl,
        [string]$parentPath = ""
    )

    Write-Host "Processing folder: $folderRelativeUrl" -ForegroundColor Cyan

    # Get all items in the current folder
    $folderItems = Get-PnPFolderItem -FolderSiteRelativeUrl $folderRelativeUrl

    # Initialize counters for the current folder
    $folderFileCount = 0
    $folderSize = 0

    foreach ($item in $folderItems) {
        if ($item.FileSystemObjectType -eq "Folder") {
            # If the item is a folder, recursively process it
            $global:totalFolders++
            $folderPath = "$parentPath/$($item.Name)"
            Write-Host "Found subfolder: $folderPath" -ForegroundColor Yellow
            Get-FolderTree -folderRelativeUrl $item.ServerRelativeUrl -parentPath $folderPath
        } elseif ($item.FileSystemObjectType -eq "File") {
            # If the item is a file, add it to the report
            $global:totalFiles++
            $folderFileCount++
            $fileSize = [math]::Round($item.Length / 1KB, 2) # Convert size to KB
            $folderSize += $fileSize

            $treeSizeReport += [PSCustomObject]@{
                FolderPath  = $parentPath
                FileName    = $item.Name
                FileSizeKB  = $fileSize
                FileType    = $item.Name.Split(".")[-1] # Extract file extension
            }
        }
    }

    # Add folder summary to the report
    if ($parentPath -ne "") {
        $treeSizeReport += [PSCustomObject]@{
            FolderPath  = $parentPath
            FileName    = "Folder Summary"
            FileSizeKB  = $folderSize
            FileType    = "Folder"
            FileCount   = $folderFileCount
        }
    }
}

# Start processing the root folder
Write-Host "Starting folder traversal..." -ForegroundColor Cyan
Get-FolderTree -folderRelativeUrl $folderRelativeUrl
Write-Host "Folder traversal complete." -ForegroundColor Green

# Display total counts
Write-Host "Total folders found: $totalFolders" -ForegroundColor Cyan
Write-Host "Total files found: $totalFiles" -ForegroundColor Cyan

# Export the TreeSize report to a CSV file
$reportPath = "$env:USERPROFILE\Desktop\TreeSizeReport.csv"
Write-Host "Exporting TreeSize report to: $reportPath" -ForegroundColor Cyan
$treeSizeReport | Export-Csv -Path $reportPath -NoTypeInformation
Write-Host "TreeSize report saved to $reportPath" -ForegroundColor Green
