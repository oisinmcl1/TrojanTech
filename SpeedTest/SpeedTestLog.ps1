# Define the path where Speedtest CLI is located
$speedtestDir = "C:\Speedtest"

# Define the paths for the CSV file and the speed text files
$csvFilePath = Join-Path $speedtestDir "speedslog.csv"
$downloadSpeedFilePath = Join-Path $speedtestDir "download_speed.txt"
$uploadSpeedFilePath = Join-Path $speedtestDir "upload_speed.txt"

# Run speed test and store result in JSON format
$speedtestResult = & "$speedtestDir\speedtest.exe" --format=json
$jsonResult = ConvertFrom-Json $speedtestResult

# Extract download and upload speeds in Mbps
$downloadSpeed = [math]::round($jsonResult.download.bandwidth * 8 / 1000000, 2)
$uploadSpeed = [math]::round($jsonResult.upload.bandwidth * 8 / 1000000, 2)

# Extract additional data
$timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
$ping = $jsonResult.ping.latency
$serverName = $jsonResult.server.name
$serverLocation = $jsonResult.server.location
$serverCountry = $jsonResult.server.country
$isp = $jsonResult.isp
$resultID = $jsonResult.result.id

# Write the latest download and upload speeds to text files
"$downloadSpeed Mbps" | Out-File $downloadSpeedFilePath -Encoding utf8
"$uploadSpeed Mbps" | Out-File $uploadSpeedFilePath -Encoding utf8

# Create or append to the CSV file
if (-Not (Test-Path $csvFilePath)) {
    # Create CSV with headers if it doesn't exist
    $headers = "Timestamp,DownloadSpeedMbps,UploadSpeedMbps,PingLatency,ServerName,ServerLocation,ServerCountry,ISP,ResultID"
    $headers | Out-File $csvFilePath -Encoding utf8
}

# Append the new result to the CSV file
$csvEntry = "$timestamp,$downloadSpeed,$uploadSpeed,$ping,$serverName,$serverLocation,$serverCountry,$isp,$resultID"
$csvEntry | Out-File $csvFilePath -Append -Encoding utf8

Write-Output "Speed test results appended to $csvFilePath."
Write-Output "Latest download speed: $downloadSpeed Mbps"
Write-Output "Latest upload speed: $uploadSpeed Mbps"
