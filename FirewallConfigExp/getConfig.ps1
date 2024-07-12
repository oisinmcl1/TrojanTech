if (-not ([System.Management.Automation.PSTypeName]'TrustAllCertsPolicy').Type) {
    Add-Type @"
    using System.Net;
    using System.Security.Cryptography.X509Certificates;

    public class TrustAllCertsPolicy : ICertificatePolicy {
        public bool CheckValidationResult(
            ServicePoint srvPoint, X509Certificate certificate,
            WebRequest request, int certificateProblem) {
            return true;
        }
    }
"@
}
[System.Net.ServicePointManager]::CertificatePolicy = New-Object TrustAllCertsPolicy



# Get the default gateway IP address (firewall)
$defaultGateway = (Get-NetIPConfiguration).IPv4DefaultGateway[0].NextHop

# Debug output to confirm default gateway extraction
## Write-Output "Default Gateway: $defaultGateway"

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

    # URL encode the password to handle special characters
    $encodedPassword = [System.Net.WebUtility]::UrlEncode($passwordPlainText)

    # Define the parameters for API key generation
    $uri = "https://$firewallIp/api/?type=keygen"
    $contentType = "application/x-www-form-urlencoded"
    $body = "user=$username&password=$encodedPassword"

    # Send the POST request to generate API key
    try {
        $response = Invoke-WebRequest -Uri $uri -Method Post -ContentType $contentType -Body $body
    
        # Parse the response to extract the API key
        $responseXml = [xml]$response.Content
        $apiKey = $responseXml.response.result.key

        ## Write-Output "API Key retrieved"
    
    } 
    catch {
        Write-Error "Error retrieving API Key: $_"
        exit
    }

    return $apiKey
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
    ## Write-Output "Exporting config with URL: $uri"
    
    try {
        Invoke-WebRequest -Uri $uri -OutFile $outputFile
        Write-Output "Configuration exported successfully to: $outputFile"
    } 
    catch {
        Write-Error "Failed to export configuration: $_"
    }
}


# Check if the directory exists
if (-not (Test-Path -Path "C:\PaloAltoConfig\")) {
    # If the directory does not exist, create it
    New-Item -ItemType directory -Path "C:\PaloAltoConfig\"
}

$password = Read-Host "Password" -AsSecureString

$outputFile = "C:\PaloAltoConfig\config.xml"

try {
    # Try to retrieve the API key using the firewall IP and password
    $apiKey = Get-APIKey -firewallIp $firewallIp -password $password
    Write-Output "API Key retrieved successfully."
    ## Write-Output $apiKey
} 
catch {
    # Catch and display errors if the API key retrieval fails
    Write-Error "Failed to retrieve API Key: $_"
}

try {
    # Try to export the firewall configuration using the API key
    Export-FirewallConfig -firewallIp $firewallIp -apiKey $apiKey -outputFile $outputFile
} 
catch {
    # Catch and display errors if the configuration export fails
    Write-Error "Failed to export configuration: $_"
}
