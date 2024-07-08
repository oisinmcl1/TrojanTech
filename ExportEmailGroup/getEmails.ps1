# Check if ad module is installed.
if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
    Write-Host ("Active Directory Module not installed")
    exit
}

# Prompt for name of AD
$groupName = Read-Host -Prompt ("Enter AD group name")

<#
if (!$groupName) {
    Write-Host ("No AD group entered")
    exit
}
#>

# Attempt to retrieve AD group obj based on entered group name
try {
    $group = Get-ADGroup -Identity $groupName -ErrorAction Stop
}
# Output message and exit if group doesn't exist
catch {
    Write-Host ("Group does not exist in AD")
    exit
}

# Retrieve members array
$members = Get-ADGroupMemeber -Identity $groupName


# Array of emails retrieved
$emailAddr = @()