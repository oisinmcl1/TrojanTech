# Import the module
Import-Module PSWarranty

# Define your device's serial number and product number

$deviceSerial = "5CG021392W"

#$deviceSerial = "5CD3476QKC"

# Get the warranty information
$warrantyInfo = Get-WarrantyInfo -DeviceSerial $deviceSerial

# Output the warranty information
Write-Output $warrantyInfo
