# Ignore SSL certificate errors
Add-Type @"
using System.Net;
using System.Net.Security;
using System.Security.Cryptography.X509Certificates;
public static class TrustAllCertsPolicy {
    public static void Ignore() {
        ServicePointManager.ServerCertificateValidationCallback = 
            new RemoteCertificateValidationCallback(
                delegate { return true; }
            );
    }
}
"@
[TrustAllCertsPolicy]::Ignore()

# Ensure using all SSL/TLS protocols
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 -bor [Net.SecurityProtocolType]::Tls11 -bor [Net.SecurityProtocolType]::Tls -bor [Net.SecurityProtocolType]::Ssl3

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

$firewallIp = Read-Host "IP of firewall"
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
