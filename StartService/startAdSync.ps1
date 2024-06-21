# Define the service name
$serviceName = "ADSync"

# Get the service status
$service = Get-Service -Name $serviceName -ErrorAction SilentlyContinue

if ($service) {
    if ($service.Status -ne 'Running') {
        Write-Output "Service $serviceName is not running. Attempting to restart..."
        try {
            # Attempt to restart the service
            Restart-Service -Name $serviceName -Force -ErrorAction Stop
            Write-Output "Service $serviceName restarted successfully."
        } catch {
            Write-Output "Failed to restart service $serviceName. Error: $_"
        }
    } else {
        Write-Output "Service $serviceName is running."
    }
} else {
    Write-Output "Service $serviceName not found."
}
