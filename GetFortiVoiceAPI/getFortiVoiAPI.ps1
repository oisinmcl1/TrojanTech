param(
    [string]$ip,
    [string]$username,
    [string]$password
)

# Function to check if module is installed and import it
function ModuleInstall {
    param(
        [string]$moduleName
    )

    # NuGet needed for these modules, check if installed, if not install
    if (!(Get-PackageProvider -Name NuGet)) {
        try {
            Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force
        }
        catch {
            Write-Error "Failed to install NuGet: `n $_"
        }
    }

    # Install module if not installed already
    if(!(Get-Module -ListAvailable -Name $moduleName)) {
        try {
            Install-Module -Name $moduleName -Repository PSGallery -Force
        }
        catch {
            Write-Error "Failed to Install $moduleName :`n $_"
            exit
        }
    }
    
    # Import module
    try{
        Import-Module $moduleName -ErrorAction Stop
    }
    catch {
        Write-Error "Failed to Import $moduleName :`n $_"
        exit
    }    
}

## ModuleInstall -moduleName "Posh-SSH"
ModuleInstall -moduleName "PSWriteHTML"
ModuleInstall -moduleName "Invoke-RestMethod"


# Function to enable rest API through SSH
function Enable-RestAPI {
    param(
    [string]$ip,
    [string]$username,
    [string]$password
)

    # Establish Connection to FortiVoice System
    $credential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $username, (ConvertTo-SecureString -String $password -AsPlainText -Force)
    $sshSesh = New-SSHSession -ComputerName $ip -Credential $credential

    # Run CLI commands to enable REST API
    Invoke-SSHCommand -SSHSession $sshSesh -Command "config system global"
    Invoke-SSHCommand -SSHSession $sshSesh -Command "set rest-api enable"
    Invoke-SSHCommand -SSHSession $sshSesh -Command "end"

    # Close SSH
    Remove-SSHSession -SSHSession $sshSesh
}

# THIS SECTION AUTHENTICATES FOR FORTIVOICE API AND RETRIEVES APSSCOOKIE

function Get-FortiCookie {
    param(
    [string]$ip,
    [string]$username,
    [string]$password
    )

    $url = "https://$ip/api/v1/VoiceAdminLogin"
}



# Enable Rest API
try {
    Write-Output "Enabling Rest-API`n"
    Enable-RestAPI -ip $ip -username $username -password $password
}
catch {
    Write-Error "Failed to enable Rest API: `n $_"
}