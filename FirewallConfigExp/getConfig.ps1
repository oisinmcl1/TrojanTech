# Gets default gateway ip (firewall)
$defaultGateway = (Get-NetIPConfiguration).IPv4DefaultGateway[0].NextHop

# Checks that default gateway ends in ".1" (indicates that it is firewall)
if ($defaultGateway.SubString($defaultGateway.length - 2) -eq ".1") {
    $firewallIp = $defaultGateway
}
# Otherwise prompt the user to enter firewall ip
else {
    $firewallIp = Read-Host "Enter firewall IP"
}

function Get-APIKey {
    param (
        [string]$firewallIp,
        [securestring]$password
    )
    $username = "admin"

    # Convert SecureString to plain text
    $passwordPlainText = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($password))

    $uri = "https://$firewallIp/api/?type=keygen&user=$username&password=$passwordPlainText"
    $response = Invoke-RestMethod -Uri $uri -Method Post

    return $response.response.result.key
}

function Export-FirewallConfig {
    param (
        [string]$firewallIp,
        [string]$apiKey,
        [string]$outputFile
    )

    $uri = "https://$firewallIp/api/?type=export&category=configuration&key=$apiKey"
    Invoke-WebRequest -Uri $uri -OutFile $outputFile
}

$password = Read-Host "Password" -AsSecureString
$outputFile = "C:\PaloAltoConfig\config.xml"

try {
    $apiKey = Get-APIKey -firewallIp $firewallIp -password $password
    Write-Output "API Key retrieved successfully."
} catch {
    Write-Error "Failed to retrieve API Key: $_"
}

try {
    Export-FirewallConfig -firewallIp $firewallIp -apiKey $apiKey -outputFile $outputFile
    Write-Output "Configuration exported successfully."
} catch {
    Write-Error "Failed to export configuration: $_"
}
