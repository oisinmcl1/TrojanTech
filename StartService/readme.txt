PowerShell Script for Service Monitoring and Restart

Originally intended for starting ADSync when not running but can be changed to any other service.

This PowerShell script is designed to monitor and restart a specific service on a Windows machine. The service is defined by the variable `$serviceName`.

The script performs the following operations:

1. Retrieves the status of the service defined by `$serviceName`.
2. If the service is found but not running, it attempts to restart the service.
3. If the service restarts successfully, it outputs a success message.
4. If the service fails to restart, it outputs an error message with the details of the failure.
5. If the service is already running, it outputs a message indicating that the service is running.
6. If the service is not found, it outputs a message indicating that the service was not found.

To use this script, replace `$serviceName` with the name of the service you want to monitor and restart.

Please note: This script requires PowerShell and appropriate administrative permissions to run.


Oisín Mc Laughlin
10/06/2024