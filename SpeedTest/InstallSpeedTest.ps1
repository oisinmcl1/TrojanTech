# Define the URL for Speedtest CLI installer
$url = "https://install.speedtest.net/app/cli/ookla-speedtest-1.2.0-win64.zip"

# Define the path where Speedtest CLI will be installed
$installDir = "C:\Speedtest"

# Create the directory if it doesn't exist
if (-not (Test-Path $installDir)) {
    New-Item -ItemType Directory -Path $installDir | Out-Null
}

# Define the path for the downloaded installer
$zipFilePath = Join-Path $installDir "ookla-speedtest-1.2.0-win64.zip"

# Download the Speedtest CLI installer
Invoke-WebRequest -Uri $url -OutFile $zipFilePath

# Extract the contents of the installer ZIP file
Expand-Archive -Path $zipFilePath -DestinationPath $installDir

# Accept the license agreements automatically
$process = Start-Process -FilePath (Join-Path $installDir "speedtest.exe") -ArgumentList "--accept-license --accept-gdpr" -NoNewWindow -PassThru
$process.WaitForExit()

# Cleanup: Remove the downloaded ZIP file
Remove-Item -Path $zipFilePath

Write-Output "Speedtest CLI installed successfully and license agreements accepted."
