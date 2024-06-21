$serialNumber = "5CG021392W"

# Check if the HPWarranty module is installed
if (!(Get-Module -ListAvailable -Name HPWarranty)) {
    try {
        # Install the HPWarranty module
        Install-Module -Name HPWarranty -Force -Scope CurrentUser
    } catch {
        Write-Error "Failed to install the HPWarranty module: $_"
        exit
    }
}

# Import the HPWarranty module
try {
    Import-Module HPWarranty
} catch {
    Write-Error "Failed to import the HPWarranty module: $_"
    exit
}

# Check if the cmdlet exists
if (-not (Get-Command -Name Get-HPIncWarrantyEntitlement -Module HPWarranty -ErrorAction SilentlyContinue)) {
    Write-Error "The Get-HPIncWarrantyEntitlement cmdlet is not available."
    exit
}

# Get the warranty information
try {
    $warrantyInfo = Get-HPIncWarrantyEntitlement -SerialNumber $serialNumber
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
