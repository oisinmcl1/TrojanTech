# Ignore SSL certificate errors somehow?
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

# Get the default gateway IP address (firewall)
$defaultGateway = (Get-NetIPConfiguration).IPv4DefaultGateway[0].NextHop

# Debug output to confirm default gateway extraction
Write-Output "Default Gateway: $defaultGateway"

# Check if the default gateway ends in ".1" (indicating it might be the firewall IP)
if ($defaultGateway.EndsWith(".1")) {
    $firewallIp = $defaultGateway
} else {
    # Prompt the user to enter the firewall IP if the default gateway does not end in ".1"
    $firewallIp = Read-Host "Enter firewall IP"
}

# Debug output to confirm firewall IP selection
Write-Output "Using Firewall IP: $firewallIp"

function ConvertFrom-SecureStringToPlainText {
    param (
        [System.Security.SecureString]$secureString
    )
    return [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($secureString))
}

function Get-APIKey {
    param (
        [string]$firewallIp,
        [securestring]$password
    )
    $username = "admin"

    # Convert SecureString to plain text for the API request
    $passwordPlainText = ConvertFrom-SecureStringToPlainText -secureString $password

    # Debug output to confirm API request URL
    $uri = "https://$firewallIp/api/?type=keygen&user=$username&password=$passwordPlainText"
    Write-Output "Requesting API Key with URL: $uri"
    $response = Invoke-RestMethod -Uri $uri -Method Post -ErrorAction Stop

    # Return the API key from the response
    if ($response.response.status -ne "success") {
        throw "API Key retrieval failed. Response: $($response | ConvertTo-Json -Compress)"
    }
    return $response.response.result.key
}

function Export-FirewallConfig {
    param (
        [string]$firewallIp,
        [string]$apiKey,
        [string]$outputFile
    )

    # Create export configuration request URL
    $uri = "https://$firewallIp/api/?type=export&category=configuration&key=$apiKey"
    
    # Debug output to confirm export URL
    Write-Output "Exporting config with URL: $uri"
    Invoke-WebRequest -Uri $uri -OutFile $outputFile -ErrorAction Stop
}

$password = Read-Host "Password" -AsSecureString
$outputFile = "C:\PaloAltoConfig\config.xml"

try {
    # Try to retrieve the API key using the firewall IP and password
    $apiKey = Get-APIKey -firewallIp $firewallIp -password $password
    Write-Output "API Key retrieved successfully."
} catch {
    # Catch and display errors if the API key retrieval fails
    Write-Error "Failed to retrieve API Key: $_"
    exit 1
}

try {
    # Try to export the firewall configuration using the API key
    Export-FirewallConfig -firewallIp $firewallIp -apiKey $apiKey -outputFile $outputFile
    Write-Output "Configuration exported successfully to $outputFile."
} catch {
    # Catch and display errors if the configuration export fails
    Write-Error "Failed to export configuration: $_"
    exit 1
}
