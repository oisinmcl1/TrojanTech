function Get-APIKey {
    param (
        [string]$firewallIp,
        [securestring]$password = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($password))
    )
    $username = "admin"

    $uri = "https://$firewallIp/api/?type=keygen&user=$username&password=$password"
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

$firewallIp = Read-Host "IP of firewall"
$password = Read-Host "Password" -AsSecureString
$outputFile = "C:\PaloAltoConfig\config.xml"

$apiKey = Get-APIKey -firewallIp $firewallIp -password $password
Export-FirewallConfig -firewallIp $firewallIp -apiKey $apiKey -outputFile $outputFile

# Hopefully gets API key using palo alto api?
# Untested though!!!