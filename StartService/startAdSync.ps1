# Define the service name
$serviceName = "ADSync"

# Write output with timestamp
function Write-Log {
    param (
        [string]$message
    )
    Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss") - $message"
}

# Get the service status
$service = Get-Service -Name $serviceName -ErrorAction SilentlyContinue

if ($service) {
    if ($service.Status -ne 'Running') {
        Write-Log "Service $serviceName is not running. Attempting to restart..."
        try {
            # Check if the service can be stopped
            if ($service.CanStop) {
                # Attempt to restart the service
                Restart-Service -Name $serviceName -Force -ErrorAction Stop
                Write-Log "Service $serviceName restarted successfully."
            } else {
                Write-Log "Service $serviceName cannot be stopped."
            }
        } catch {
            Write-Log "Failed to restart service $serviceName. Error: $_"
        }
    } else {
        Write-Log "Service $serviceName is running."
    }
} else {
    Write-Log "Service $serviceName not found."
}
