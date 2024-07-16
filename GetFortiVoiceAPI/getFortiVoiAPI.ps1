param(
    [string]$ip,
    [string]$username
    ## [string]$password
)

# Function to check if module is installed and import it
function ModuleInstall {
    param(
        [string]$moduleName
    )

    # Update NuGet to install needed modules
    Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force

    # Install module if not installed already
    if(!(Get-Module -ListAvailable -Name $moduleName)) {
        try {
            Install-Module -Name $moduleName -Repository PSGallery -Force
        }
        catch {
            Write-Error "Failed to install $moduleName :`n $_"
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

<#

# THIS SECTION ENABLES RESTFUL API THROUGH SSH

# Establish Connection to FortiVoice System
$sshSesh = New-SSHSession -ComputerName $ip -Credential (Get-Credential)

# Run CLI commands to enable REST API
Invoke-SSHCommand -SSHSession $sshSesh -Command "config system global"
Invoke-SSHCommand -SSHSession $sshSesh -Command "set rest-api enable"
Invoke-SSHCommand -SSHSession $sshSesh -Command "end"

# Close SSH
Remove-SSHSession -SSHSession $sshSesh

#>