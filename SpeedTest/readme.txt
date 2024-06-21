Note

Permissions: Ensure correct permissions to download and install the Speedtest CLI.

Firewall and Security: Make sure firewall or security software are configured to allow the Speedtest CLI download and installation.

Non-Commercial Use: The Speedtest CLI is intended for personal, non-commercial use only. Exercise caution if considering commercial use (careful).

Makes use of the Speedtest Cli: https://www.speedtest.net/apps/cli



InstallSpeedTest.ps1

Download Installer: Downloads the Speedtest CLI installer from Ookla's official website.

Extract Installer: Extracts the contents of the downloaded ZIP file to a temporary directory.

Install Speedtest CLI: Copies the Speedtest CLI executable (speedtest.exe) to the specified installation directory (C:\Speedtest by default).

Cleanup: Removes temporary files and directories to tidy up the system after installation.

Completion: Notifies the user when the installation process is complete.



SpeedTestLog.ps1

Run Speed Test: Executes the Speedtest CLI and captures the result in JSON format.

Extract Speeds: Extracts and converts the download and upload speeds from bits per second to Mbps.

Write Speeds to Text Files: Writes the latest download and upload speeds to separate text files (download_speed.txt and upload_speed.txt).

Update CSV Log: Appends the new result to the existing CSV file (speedslog.csv), creating headers if the file does not exist.

Output: Prints messages indicating the completion of the logging process and the latest speeds.



speedlog.csv

Timestamp: The date and time when the speed test was performed.

DownloadSpeedMbps: The download speed measured in Mbps.

UploadSpeedMbps: The upload speed measured in Mbps.

PingLatency: The latency (ping time) in milliseconds.

ServerName: The name of the server used for the speed test.

ServerLocation: The location of the server used for the speed test.

ServerCountry: The country where the server is located.

ISP: The Internet Service Provider used for the speed test.

ResultID: A unique identifier for the specific speed test result.



download_speed.txt

Latest Download Speed: Contains the latest download speed in Mbps.



upload_speed.txt

Latest Upload Speed: Contains the latest upload speed in Mbps.



Oisin Mc Laughlin
24/05/2024
