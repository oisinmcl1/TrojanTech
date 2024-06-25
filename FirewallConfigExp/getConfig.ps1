function Get-APIKey {
    param (
        [string]$firewallIp,
        [securestring]$password
    )

    $username = "admin"
}

function Export-FirewallConfig {
    param (
        [string]$firewallIp,
        [string]$apiKey,
        [string]$outputFile
    )
}

$firewallIp = Read-Host "Enter the IP address of the firewall"
$password = Read-Host "Enter your password" -AsSecureString
$outputFile = "C:"
