# Import the PnP PowerShell module
Import-Module PnP.PowerShell

# Define your SharePoint site URL
$siteUrl = Read-Host "Enter the SharePoint site URL (or type 'exit' to quit)"

# Exit if the user types 'exit'
if ($siteUrl -eq 'exit') {
    Write-Host "Exiting script."
    exit
}

# Connect to the SharePoint site
Connect-PnPOnline -Url $siteUrl -Interactive

# Function to display the main menu
function Show-MainMenu {
    Write-Host "`nSelect the storage quota increase:"
    Write-Host "1. Increase by 500 GB"
    Write-Host "2. Increase by 1 TB"
    Write-Host "3. Enter a custom amount"
    Write-Host "Type 'exit' to quit"
}

# Function to handle custom amount input
function Get-CustomAmount {
    while ($true) {
        Write-Host "`nEnter the amount (e.g., '250 GB' or '1 TB') or type 'back' to return to the main menu:"
        $input = Read-Host
        if ($input -eq 'back') {
            return $null
        }
        if ($input -match '^\d+\s*(GB|TB)$') {
            $amount = [double]($input -replace '[^0-9.]', '')
            $unit = $matches[1]
            if ($unit -eq 'TB') {
                $amount = $amount * 1TB
            } else {
                $amount = $amount * 1GB
            }
            return $amount
        } else {
            Write-Host "Invalid input. Please enter a valid amount (e.g., '250 GB' or '1 TB')."
        }
    }
}

# Main script logic
while ($true) {
    Show-MainMenu
    $choice = Read-Host "Enter your choice (1, 2, 3, or 'exit')"

    # Exit if the user types 'exit'
    if ($choice -eq 'exit') {
        Write-Host "Exiting script."
        Disconnect-PnPOnline
        exit
    }

    # Process the user's choice
    switch ($choice) {
        1 {
            $increaseAmount = 500GB
            Write-Host "Increasing storage quota by 500 GB..."
            break
        }
        2 {
            $increaseAmount = 1TB
            Write-Host "Increasing storage quota by 1 TB..."
            break
        }
        3 {
            $increaseAmount = Get-CustomAmount
            if ($increaseAmount -eq $null) {
                continue
            }
            Write-Host "Increasing storage quota by $([math]::Round($increaseAmount / 1GB, 2)) GB..."
            break
        }
        default {
            Write-Host "Invalid choice. Please enter 1, 2, 3, or 'exit'."
            continue
        }
    }

    # Get the current storage quota and warning level
    $site = Get-PnPTenantSite -Url $siteUrl
    $currentQuota = $site.StorageQuota
    $currentWarningLevel = $site.StorageQuotaWarningLevel

    # Calculate the new storage quota and warning level
    $newQuota = $currentQuota + $increaseAmount
    $newWarningLevel = $currentWarningLevel + $increaseAmount

    # Apply the new storage quota and warning level
    Set-PnPTenantSite -Url $siteUrl -StorageQuota $newQuota -StorageQuotaWarningLevel $newWarningLevel

    # Display the updated storage quota
    Write-Host "`nNew Storage Quota: $([math]::Round($newQuota / 1GB, 2)) GB"
    Write-Host "New Warning Level: $([math]::Round($newWarningLevel / 1GB, 2)) GB"
    Write-Host "Storage quota has been increased successfully.`n"
}
