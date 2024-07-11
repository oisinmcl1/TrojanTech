# I'm getting an error when I Invoke-RestMethod to get the API Key, this first part attempts to bypass it.
# Thank you Jamiker from stackoverflow

# Check if ServerCertificateValidationCallback type already exists before adding
if (-not ([System.Management.Automation.PSTypeName]'ServerCertificateValidationCallback').Type) {
    $certCallback = @"
        using System;
        using System.Net;
        using System.Net.Security;
        using System.Security.Cryptography.X509Certificates;
        
        public class ServerCertificateValidationCallback
        {
            public static void Ignore()
            {
                if (ServicePointManager.ServerCertificateValidationCallback == null)
                {
                    ServicePointManager.ServerCertificateValidationCallback += 
                        delegate(object sender, X509Certificate certificate, X509Chain chain, SslPolicyErrors sslPolicyErrors)
                        {
                            return true;
                        };
                }
            }
        }
"@
    Add-Type $certCallback
}

# Set all necessary SSL/TLS protocols
[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12 -bor `
                                                     [System.Net.SecurityProtocolType]::Tls11 -bor `
                                                     [System.Net.SecurityProtocolType]::Tls

# Ignore SSL certificate errors
[ServerCertificateValidationCallback]::Ignore()



# Get the default gateway IP address (firewall)
$defaultGateway = (Get-NetIPConfiguration).IPv4DefaultGateway[0].NextHop

# Debug output to confirm default gateway extraction
Write-Output "Default Gateway: $defaultGateway"

# Check if the default gateway ends in ".1" (indicating it might be the firewall IP)
if ($defaultGateway.SubString($defaultGateway.Length - 2) -eq ".1") {
    $firewallIp = $defaultGateway
} else {
    # Prompt the user to enter the firewall IP if the default gateway does not end in ".1"
    $firewallIp = Read-Host "Enter firewall IP"
}

# Debug output to confirm firewall IP selection
Write-Output "Using Firewall IP: $firewallIp"

function Get-APIKey {
    param (
        [string]$firewallIp,
        [securestring]$password
    )
    $username = "admin"

    # Convert SecureString to plain text for the API request
    $passwordPlainText = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($password))

    # Debug output to confirm API request URL
    $uri = "https://$firewallIp/api/?type=keygen&user=$username&password=$passwordPlainText"
    Write-Output "Requesting API Key with URL: $uri"
    $response = Invoke-RestMethod -Uri $uri -Method Post

    # Debug output to display API key retrieved
    Write-Output "API Key retrieved: $($response.response.result.key)"

    # Return the API key from the response
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
    
    try {
        Invoke-WebRequest -Uri $uri -OutFile $outputFile
        Write-Output "Configuration exported successfully to: $outputFile"
    } catch {
        Write-Error "Failed to export configuration: $_"
    }
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
}

try {
    # Try to export the firewall configuration using the API key
    Export-FirewallConfig -firewallIp $firewallIp -apiKey $apiKey -outputFile $outputFile
} catch {
    # Catch and display errors if the configuration export fails
    Write-Error "Failed to export configuration: $_"
}