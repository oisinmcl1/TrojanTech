$logfile = "pinglog.txt"
$deviceIP = Read-Host "Device IP"

while ($true) {
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $ping = Test-Connection -ComputerName $devicerIP -Count 1 -ErrorAction SilentlyContinue
    
    if ($ping) {
        $status = "Online"
    } else {
        $status = "Offline"
    }

    "$timestamp - $status" | Out-File -FilePath $logfile -Append
    Start-Sleep -Seconds 10
}
