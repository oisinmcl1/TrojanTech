param(
    [string]$ip,
    [string]$interval
)

$dateID = Get-Date -Format yyyyMMdd-HHmmss
$logfile = "pinglog-$dateID.txt"
$dir = Get-Location

Write-Output "=======================================`nStarting Ping Monitor at $dateID for:`n$ip`n`nLog file location:`n$dir\$logfile`n`nInterval:`n$interval seconds`n`nPress Control-C to stop.`n======================================="
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
