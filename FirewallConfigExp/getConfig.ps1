param (
    [string]$username,
    [string]$encodedPassword,
    [string]$outputDir
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

function Export-FirewallConfig {
    param (
        [string]$firewallIp,
        [string]$apiKey,
        [string]$outputFile
    )

    # Create export configuration request URL
    $uri = "https://$firewallIp/api/?type=export&category=configuration&key=$apiKey"
    
    try {
        Invoke-WebRequest -Uri $uri -OutFile $outputFile
        Write-Output "Configuration exported successfully to: $outputFile"
    } 
    catch {
        Write-Error "Failed to export configuration: $_"
    }
}

# Ensure the output directory exists
if (-not (Test-Path -Path $outputDir)) {
    # If the directory does not exist, create it
    try {
        New-Item -ItemType directory -Path $outputDir
        Write-Output "Directory created: $outputDir"
    } catch {
        Write-Error "Failed to create directory: $_"
        exit
    }
}

# Get device name and substring of client's name
$device = (Get-ComputerInfo).CsName
$device = $device.SubString(4, 3)

# Get the current date and format it as yyyyMMdd
$date = Get-Date -Format "yyyyMMdd"

# Create the output file path with the date prefix
$outputFile = Join-Path -Path $outputDir -ChildPath "${date}_${device}_config.xml"

try {
    # Try to retrieve the API key using the firewall IP, username, and password
    $apiKey = Get-APIKey -firewallIp $firewallIp -username $username -password $password
    Write-Output "API Key retrieved successfully."
} 
catch {
    # Catch and display errors if the API key retrieval fails
    Write-Error "Failed to retrieve API Key: $_"
    exit
}

try {
    # Try to export the firewall configuration using the API key
    Export-FirewallConfig -firewallIp $firewallIp -apiKey $apiKey -outputFile $outputFile
} 
catch {
    # Catch and display errors if the configuration export fails
    Write-Error "Failed to export configuration: $_"
}

<#

.\getConfig.ps1 -username "admin" -encodedPassword "base64EncodedPassword" -outputDir "C:\temp"

#>