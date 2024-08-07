param(
    [string]$ip,
    [string]$interval
)

# If no ip or interval provided as parameters, prompt using read-host
if (-not $ip) {
    $ip = Read-Host "Ip of Device"
}
if (-not $interval) {
    $interval = Read-Host "Interval of ping in seconds"
}

# Date for naming of log file
$dateID = Get-Date -Format yyyyMMdd-HHmmss
$logfile = "pinglog-$dateID.txt"
#Location for cli output
$dir = Get-Location

Write-Output "=================================================`nStarting Ping Monitor at $dateID for:`n$ip`n`nLog file location:`n$dir\$logfile`n`nInterval:`n$interval seconds`n`nPress Control-C to stop.`n================================================="
while ($true) {
    # Get timestamp and ping device
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $ping = Test-Connection -ComputerName $ip -Count 1 -ErrorAction SilentlyContinue
    
    # Log status if ping is online or offline
    if ($ping) {
        $status = "Online"
    } 
    else {
        $status = "Offline"
    }

    # Append log to txt file and sleep for duration of interval
    "$timestamp - $status" | Out-File -FilePath $logfile -Append
    Start-Sleep -Seconds $interval
}
