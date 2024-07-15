# Import the Active Directory module
Import-Module ActiveDirectory

# Retrieve all user accounts in Active Directory and select the required properties
$users = Get-ADUser -Filter * -Properties SamAccountName, DisplayName

# Ensure the C:\temp directory exists
$directory = "C:\temp"
if (-not (Test-Path -Path $directory)) {
    New-Item -ItemType Directory -Path $directory
}

# Export the selected properties to a CSV file in C:\temp
$filePath = Join-Path -Path $directory -ChildPath "usernames.csv"
$users | Select-Object SamAccountName, DisplayName | Export-Csv $filePath -NoTypeInformation

# Output the location of the exported CSV file
Write-Host "Usernames and display names have been exported to $filePath"