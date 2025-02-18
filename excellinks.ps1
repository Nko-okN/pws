# Import necessary modules
Import-Module PnP.PowerShell

# SharePoint site and folder details
$siteUrl = "https://your-sharepoint-site-url"
$folderRelativeUrl = "/sites/your-site/Shared Documents/YourFolder"

# Connect to SharePoint site
Write-Host "Connecting to SharePoint site: $siteUrl" -ForegroundColor Cyan
Connect-PnPOnline -Url $siteUrl -UseWebLogin
Write-Host "Successfully connected to SharePoint site." -ForegroundColor Green

# Get all Excel files in the specified folder
Write-Host "Retrieving Excel files from folder: $folderRelativeUrl" -ForegroundColor Cyan
$files = Get-PnPFolderItem -FolderSiteRelativeUrl $folderRelativeUrl | Where-Object { $_.Name -like "*.xlsx" }
Write-Host "Found $($files.Count) Excel files in the folder." -ForegroundColor Green

# Initialize a list to store link information
$linkReport = @()

# Function to extract links from Excel file
function Get-ExcelLinks {
    param (
        [string]$filePath
    )

    Write-Host "Analyzing file: $filePath" -ForegroundColor Cyan

    # Create an Excel application object
    Write-Host "Starting Excel application..." -ForegroundColor Cyan
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    Write-Host "Excel application started." -ForegroundColor Green

    # Open the Excel workbook
    Write-Host "Opening workbook: $filePath" -ForegroundColor Cyan
    $workbook = $excel.Workbooks.Open($filePath)
    Write-Host "Workbook opened successfully." -ForegroundColor Green

    # Initialize a list to store links
    $links = @()

    # Loop through each worksheet
    Write-Host "Processing worksheets..." -ForegroundColor Cyan
    foreach ($worksheet in $workbook.Worksheets) {
        Write-Host "Processing worksheet: $($worksheet.Name)" -ForegroundColor Cyan

        # Search for hyperlinks
        Write-Host "Searching for hyperlinks..." -ForegroundColor Cyan
        foreach ($hyperlink in $worksheet.Hyperlinks) {
            Write-Host "Found hyperlink: $($hyperlink.Address)" -ForegroundColor Yellow
            $links += [PSCustomObject]@{
                FilePath = $filePath
                SheetName = $worksheet.Name
                LinkType = "Hyperlink"
                LinkAddress = $hyperlink.Address
                LinkSubAddress = $hyperlink.SubAddress
            }
        }

        # Search for formulas containing links (e.g., VLOOKUP)
        Write-Host "Searching for links in formulas..." -ForegroundColor Cyan
        $usedRange = $worksheet.UsedRange
        foreach ($cell in $usedRange) {
            if ($cell.Formula -like "*http*") {
                Write-Host "Found link in formula: $($cell.Formula) at cell $($cell.Address)" -ForegroundColor Yellow
                $links += [PSCustomObject]@{
                    FilePath = $filePath
                    SheetName = $worksheet.Name
                    LinkType = "Formula"
                    LinkAddress = $cell.Formula
                    LinkSubAddress = $cell.Address
                }
            }
        }
    }

    # Close the workbook and quit Excel
    Write-Host "Closing workbook and quitting Excel..." -ForegroundColor Cyan
    $workbook.Close($false)
    $excel.Quit()
    Write-Host "Excel application closed." -ForegroundColor Green

    # Release COM objects
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()

    Write-Host "Finished analyzing file: $filePath" -ForegroundColor Green
    return $links
}

# Loop through each file and extract links
Write-Host "Starting analysis of Excel files..." -ForegroundColor Cyan
foreach ($file in $files) {
    Write-Host "Processing file: $($file.Name)" -ForegroundColor Cyan
    $filePath = "$env:TEMP\$($file.Name)"

    # Download the file to a temporary location
    Write-Host "Downloading file to temporary location: $filePath" -ForegroundColor Cyan
    Get-PnPFile -Url $file.ServerRelativeUrl -Path $env:TEMP -Filename $file.Name -AsFile
    Write-Host "File downloaded successfully." -ForegroundColor Green

    # Extract links from the Excel file
    Write-Host "Extracting links from file: $filePath" -ForegroundColor Cyan
    $links = Get-ExcelLinks -filePath $filePath

    # Add the links to the report
    $linkReport += $links
    Write-Host "Added $($links.Count) links from file: $($file.Name)" -ForegroundColor Green

    # Delete the temporary file
    Write-Host "Deleting temporary file: $filePath" -ForegroundColor Cyan
    Remove-Item -Path $filePath -Force
    Write-Host "Temporary file deleted." -ForegroundColor Green
}

# Export the report to a CSV file
$reportPath = "$env:USERPROFILE\Desktop\ExcelLinksReport.csv"
Write-Host "Exporting link report to: $reportPath" -ForegroundColor Cyan
$linkReport | Export-Csv -Path $reportPath -NoTypeInformation
Write-Host "Link analysis complete. Report saved to $reportPath" -ForegroundColor Green
