param(
    [string]$ip,
    [string]$username
)



# THIS SECTION ENABLES RESTFUL API THROUGH SSH

# Check if SSH module is installed
if(!(Get-Module -ListAvailable -Name Posh-SSH)) {
    try {
        Install-Module -Name Posh-SSH -Repository PSGallery -Force
    }
    catch {
        Write-Error "Failed to install Posh-SSH $_"
        exit
    }
}

# Import SSH module
try{
    Import-Module Posh-SSH -ErrorAction Stop
}
catch {
    Write-Error "Failed to Import Posh-SSH $_"
    exit
}

# Establish Connection to FortiVoice System
$sshSesh = New-SSHSession -ComputerName $ip -Credential (Get-Credential)

# Run CLI commands to enable REST API
Invoke-SSHCommand -SSHSession $sshSesh -Command "config system global"
Invoke-SSHCommand -SSHSession $sshSesh -Command "set rest-api enable"
Invoke-SSHCommand -SSHSession $sshSesh -Command "end"

# Close SSH
Remove-SSHSession -SSHSession $sshSesh
