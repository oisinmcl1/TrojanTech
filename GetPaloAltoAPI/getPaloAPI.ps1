param (
    [string]$username,
    [string]$encodedPassword
)

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

# Check if the default gateway ends in ".1" (indicating it might be the firewall IP)
if ($defaultGateway.SubString($defaultGateway.Length - 2) -eq ".1") {
    $firewallIp = $defaultGateway
} else {
    Write-Error "Default gateway does not end in '.1'. Unable to determine firewall IP."
    exit
}

Write-Output "Using Firewall IP: $firewallIp"

# Decode the base64-encoded password and convert it to a secure string
$decodedPasswordBytes = [Convert]::FromBase64String($encodedPassword)
$passwordPlainText = [Text.Encoding]::UTF8.GetString($decodedPasswordBytes)
$password = ConvertTo-SecureString -String $passwordPlainText -AsPlainText -Force

function Get-APIKey {
    param (
        [string]$firewallIp,
        [string]$username,
        [securestring]$password
    )

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

    } 
    catch {
        Write-Error "Error retrieving API Key: $_"
        exit
    }

    return $apiKey
}


# Try to retrieve the API key using the firewall IP, username, and password
try {
    $apiKey = Get-APIKey -firewallIp $firewallIp -username $username -password $password
    Write-Output "API Key retrieved successfully: $apiKey"
} 
catch {
    Write-Error "Failed to retrieve API Key: $_"
}