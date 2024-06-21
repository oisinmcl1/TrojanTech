$serialNumber = "5CG021392W"
$productNumber = "6XE66ET"

# Check if the HPWarranty module is installed
if (!(Get-Module -ListAvailable -Name HPWarranty)) {
    try {
        # Install the HPWarranty module
        Install-Module -Name HPWarranty -Force -Scope CurrentUser -AllowClobber
    } catch {
        Write-Error "Failed to install the HPWarranty module: $_"
        exit
    }
}

# Import the HPWarranty module
try {
    Import-Module HPWarranty -ErrorAction Stop
} catch {
    Write-Error "Failed to import the HPWarranty module: $_"
    exit
}

# Get the warranty information
try {
    $warrantyInfo = Get-HPIncWarrantyEntitlement -SerialNumber $serialNumber -ProductNumber $productNumber
} catch {
    Write-Error "Failed to get warranty information: $_"
    exit
}

foreach ($info in $warrantyInfo) {
    # Get the warranty end date
    $endDate = $info.EndDate

    # Check if the warranty has expired
    $hasExpired = (Get-Date) -gt $endDate

    # Output the information
    Write-Output "Serial Number: $serialNumber"
    Write-Output "Warranty End Date: $endDate"
    Write-Output "Has Warranty Expired: $hasExpired"
    Write-Output "-------------------------"
}
