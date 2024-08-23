# SMTP Server address, mail from address and rcpt to address
param (
    [string]$smtpServer,
    [string]$fromEmail,
    [string]$toEmail
)

# If no function parameters, read through cli
if (-not $smtpServer) {
    $smtpServer = Read-Host "SMTP Server address"
}
if (-not $fromEmail) {
    $fromEmail = Read-Host "Mail from address"
}
if (-not $toEmail) {
    $toEmail = Read-Host "Rcpt to address"
}

# Create a TCP connection to SMTP server on port 25
$tcpClient = New-Object System.Net.Sockets.TcpClient($smtpServer, 25)
$stream = $tcpClient.GetStream()
$reader = New-Object System.IO.StreamReader($stream)
$writer = New-Object System.IO.StreamWriter($stream)
$writer.AutoFlush = $true

# Read server response
$reader.ReadLine()

# Send Helo
$writer.WriteLine("HELO")
$reader.ReadLine()

# Mail From
$writer.WriteLine("mail from:<$fromEmail>")
$reader.ReadLine()

# Rcpt To
$writer.WriteLine("rcpt to:<$toEmail>")
$reader.ReadLine()

# Close Connection
$reader.Close()
$writer.Close()
$stream.Close()
$tcpClient.Close()

Write-Host "All SMTP commands sent successfully!"
