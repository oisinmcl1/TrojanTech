param(
    [string]$ip,
    [string]$interval
)

$logfile = "pinglog.txt"
$dir = Get-Location

Write-Output "=======================================`nStarting Ping Monitor for:`n$ip`n`nLog file location:`n$dir`n`nInterval:`n$interval seconds`n`nPress Control-C to stop.`n======================================="
while ($true) {
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $ping = Test-Connection -ComputerName $ip -Count 1 -ErrorAction SilentlyContinue
    
    if ($ping) {
        $status = "Online"
    } else {
        $status = "Offline"
    }

    "$timestamp - $status" | Out-File -FilePath $logfile -Append
    Start-Sleep -Seconds $interval
}
